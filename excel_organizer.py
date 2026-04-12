import os
import pandas as pd
import json
import requests
import lark_oapi as lark
from lark_oapi.api.bitable.v1 import *
import sys

def get_excel_files(folder_path):
    """获取文件夹中所有Excel文件"""
    excel_files = []
    for file in os.listdir(folder_path):
        # 跳过临时文件（以~$开头）
        if file.startswith('~$'):
            continue
        # 支持多种Excel格式
        if file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.xlsm'):
            excel_files.append(os.path.join(folder_path, file))
    return excel_files

def process_excel_file(file_path):
    """处理单个Excel文件，提取所需字段"""
    try:
        # 获取文件名作为第一列
        file_name = os.path.basename(file_path)
        
        # 尝试读取所有工作表
        xl = pd.ExcelFile(file_path)
        print(f"文件 {file_name} 中的工作表: {xl.sheet_names}")
        
        # 存储所有工作表的数据
        all_sheet_data = []
        
        # 遍历所有工作表，只处理包含TestCase的工作表
        for sheet_name in xl.sheet_names:
            # 只处理工作表名中含有TestCase的表
            if 'TestCase' not in sheet_name:
                print(f"\n跳过工作表: {sheet_name} (不包含TestCase)")
                continue
            
            print(f"\n处理工作表: {sheet_name}")
            
            # 读取工作表，不指定表头
            df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            
            # 特殊处理TestCase-EV工作表
            if 'TestCase-EV' in sheet_name:
                print("特殊处理TestCase-EV工作表")
                try:
                    # 尝试使用不同的表头行
                    for header_row in range(5):  # 尝试前5行作为表头
                        print(f"尝试将第 {header_row+1} 行作为表头")
                        # 读取数据
                        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
                        
                        # 过滤掉全为空的列
                        df = df.dropna(axis=1, how='all')
                        
                        # 打印列信息
                        print(f"列名: {list(df.columns)}")
                        
                        # 直接使用固定列位置
                        if len(df.columns) >= 3:
                            print("使用固定列位置提取数据")
                            # 创建结果数据框
                            result = pd.DataFrame()
                            # 尝试不同的列位置
                            for case_col in [1, 2]:  # 尝试第二列和第三列作为用例ID
                                try:
                                    result['用例ID CaseID'] = df.iloc[:, case_col]
                                    # 找到需求ID列（通常在用例ID列的下一列）
                                    req_col = case_col + 1
                                    if req_col < len(df.columns):
                                        result['需求ID Requirement ID'] = df.iloc[:, req_col]
                                    else:
                                        result['需求ID Requirement ID'] = ''
                                    
                                    # 对于计划否和测试结果，使用AB列和AC列
                                    # AB列是第28列（索引27），AC列是第29列（索引28）
                                    if len(df.columns) >= 29:
                                        result['计划否Plan or not'] = df.iloc[:, 27]  # AB列
                                        result['测试结果Test Results'] = df.iloc[:, 28]  # AC列
                                    else:
                                        result['计划否Plan or not'] = ''
                                        result['测试结果Test Results'] = ''
                                    
                                    result['FileName'] = file_name
                                    result['工作表'] = sheet_name
                                    
                                    # 打印提取的数据预览
                                    print("提取的数据预览:")
                                    print(result.head())
                                    
                                    # 过滤掉空行
                                    result = result.dropna(subset=['用例ID CaseID'])
                                    
                                    # 过滤掉表头行被当作数据的情况
                                    result = result[~result['用例ID CaseID'].astype(str).str.strip().isin(['用例ID', 'CaseID', '用例ID CaseID', 'Case ID', 'Test Case ID'])]
                                    
                                    # 过滤掉明显不是用例ID的值
                                    result = result[result['用例ID CaseID'].astype(str).str.contains('TestCase', case=False, na=False)]
                                    
                                    # 打印过滤后的数据预览
                                    print("过滤后的数据预览:")
                                    print(result.head())
                                    print(f"过滤后的数据行数: {len(result)}")
                                    
                                    # 处理需求ID的一对多关系
                                    if '需求ID Requirement ID' in result.columns:
                                        # 先替换所有的换行符和空格为逗号，然后按逗号拆分
                                        result['需求ID Requirement ID'] = result['需求ID Requirement ID'].astype(str).str.replace('\n', ',').str.replace(' ', ',')
                                        # 拆分需求ID
                                        result = result.assign(**{'需求ID Requirement ID': result['需求ID Requirement ID'].str.split(',')})
                                        result = result.explode('需求ID Requirement ID')
                                        # 去除空白
                                        result['需求ID Requirement ID'] = result['需求ID Requirement ID'].str.strip()
                                        # 过滤掉空的需求ID
                                        result = result[result['需求ID Requirement ID'] != '']
                                        
                                        # 打印处理后的数据预览
                                        print("处理需求ID后的数据预览:")
                                        print(result.head())
                                        print(f"处理后的数据行数: {len(result)}")
                                    
                                    # 只添加非空数据，即使计划否和测试结果列没有数据
                                    if not result.empty:
                                        # 确保所有必要的列都存在
                                        required_columns = ['FileName', '用例ID CaseID', '需求ID Requirement ID', '计划否Plan or not', '测试结果Test Results']
                                        for col in required_columns:
                                            if col not in result.columns:
                                                result[col] = ''
                                        # 调整列顺序
                                        result = result[required_columns]
                                        all_sheet_data.append(result)
                                        print(f"添加了 {len(result)} 条数据")
                                        break  # 找到合适的列后停止尝试
                                except Exception as e:
                                    print(f"尝试用例ID列 {case_col} 时出错: {str(e)}")
                                    continue
                            if not result.empty:
                                break  # 找到合适的表头后停止尝试
                except Exception as e:
                    print(f"处理TestCase-EV工作表时出错: {str(e)}")
                    import traceback
                    traceback.print_exc()
                continue  # 处理完TestCase-EV工作表后，继续处理下一个工作表
            
            # 尝试不同的表头行
            for header_row in range(10):  # 尝试前10行作为表头
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
                    
                    # 过滤掉全为空的列
                    df = df.dropna(axis=1, how='all')
                    
                    # 检查列数
                    if len(df.columns) < 4:
                        continue
                    
                    print(f"尝试将第 {header_row+1} 行作为表头:")
                    print(f"列名: {list(df.columns)}")
                    
                    # 尝试不同的列名变体
                    column_mapping = {
                        'CaseID': ['CaseID', '用例ID', 'Case ID', 'Test Case ID', 'TestCase', 'Test case', 'testcase'],
                        '需求ID Requirement ID': ['需求ID Requirement ID', '需求ID', 'Requirement ID', '需求编号', '需求', 'REQ', 'Requirement', 'requirement'],
                        '计划否Plan or not': ['计划否Plan or not', '计划否', 'Plan or not', '是否计划', '计划', 'Plan', 'plan', '是否'],
                        '测试结果Test Results': ['测试结果Test Results', '测试结果', 'Test Results', '结果', 'Result', '测试', 'Test', 'result', 'test']
                    }
                    
                    # 匹配实际列名
                    actual_columns = {}
                    print("\n开始匹配列名:")
                    
                    # 直接匹配已知的列名格式
                    known_columns = {
                        'CaseID': ['用例ID\nCaseID', '用例ID CaseID'],
                        '需求ID Requirement ID': ['需求ID\nRequirement ID', '需求ID Requirement ID'],
                        '计划否Plan or not': ['计划否\nPlan or not', '计划否Plan or not'],
                        '测试结果Test Results': ['测试结果\nTest Results', '测试结 果\nTest Results', '测试结果Test Results']
                    }
                    
                    # 首先尝试直接匹配已知列名
                    for target_col, possible_cols in known_columns.items():
                        for col in possible_cols:
                            if col in df.columns:
                                actual_columns[target_col] = col
                                print(f"直接匹配到 {target_col}: {col}")
                                break
                    
                    # 如果直接匹配失败，尝试更灵活的匹配
                    if len(actual_columns) < 4:
                        print("直接匹配失败，尝试灵活匹配:")
                        for target_col, possible_cols in column_mapping.items():
                            if target_col not in actual_columns:
                                for col in possible_cols:
                                    for df_col in df.columns:
                                        # 去除换行符和空格后进行匹配
                                        df_col_str = str(df_col).replace('\n', ' ').strip().lower()
                                        col_str = col.replace('\n', ' ').strip().lower()
                                        # 避免匹配到空格列
                                        if df_col_str.strip() == '':
                                            continue
                                        # 更严格的匹配
                                        if col_str in df_col_str:
                                            actual_columns[target_col] = df_col
                                            print(f"灵活匹配到 {target_col}: {df_col} (匹配关键词: {col})")
                                            break
                                    if target_col in actual_columns:
                                        break
                    
                    # 打印匹配结果
                    print(f"匹配结果: {actual_columns}")
                    
                    # 检查是否匹配到了空格列
                    if ' ' in actual_columns.values():
                        print("警告: 匹配到了空格列，可能会导致数据提取错误")
                        # 尝试使用固定列位置
                        if len(df.columns) >= 4:
                            print("尝试使用固定列位置提取数据")
                            actual_columns = {
                                'CaseID': df.columns[1],  # 第二列
                                '需求ID Requirement ID': df.columns[2],  # 第三列
                                '计划否Plan or not': df.columns[18],  # 第19列
                                '测试结果Test Results': df.columns[19]  # 第20列
                            }
                            print(f"使用固定列位置: {actual_columns}")
                    
                    # 特殊处理：根据列的位置和内容来识别
                    if not actual_columns:
                        # 打印所有列的详细信息，帮助调试
                        print("详细列信息:")
                        for i, col in enumerate(df.columns):
                            print(f"列 {i}: {col} (类型: {type(col)})")
                            # 尝试读取该列的前几个值
                            try:
                                sample_values = df[col].dropna().head(3).tolist()
                                print(f"  示例值: {sample_values}")
                            except:
                                pass
                    
                    # 检查是否找到所有必要列
                    missing_columns = [col for col in column_mapping.keys() if col not in actual_columns]
                    if not missing_columns:
                        print(f"找到所有必要列，使用第 {header_row+1} 行作为表头")
                        try:
                            # 提取数据
                            print(f"提取列: {[actual_columns[col] for col in column_mapping.keys()]}")
                            result = df[[actual_columns[col] for col in column_mapping.keys()]].copy()
                            
                            # 打印提取的数据预览
                            print("提取的数据预览:")
                            print(result.head())
                            
                            # 重命名列
                            column_rename = {
                                'CaseID': '用例ID CaseID',
                                '需求ID Requirement ID': '需求ID Requirement ID',
                                '计划否Plan or not': '计划否Plan or not',
                                '测试结果Test Results': '测试结果Test Results'
                            }
                            result.columns = [column_rename[col] for col in column_mapping.keys()]
                            result['FileName'] = file_name
                            result['工作表'] = sheet_name
                            
                            # 打印重命名后的数据预览
                            print("重命名后的数据预览:")
                            print(result.head())
                            
                            # 过滤掉表头行被当作数据的情况
                            # 检查用例ID CaseID列是否包含非空值且不是表头文本
                            if not result['用例ID CaseID'].isnull().all():
                                # 只过滤掉明显是表头文本的行，而不是所有包含Case或ID的行
                                result = result[~result['用例ID CaseID'].astype(str).str.strip().isin(['用例ID', 'CaseID', '用例ID CaseID', 'Case ID', 'Test Case ID'])]
                                # 过滤掉空行
                                result = result.dropna(subset=['用例ID CaseID'])
                                
                                # 打印过滤后的数据预览
                                print("过滤后的数据预览:")
                                print(result.head())
                                print(f"过滤后的数据行数: {len(result)}")
                            
                            # 处理需求ID的一对多关系
                            if '需求ID Requirement ID' in result.columns:
                                # 先替换所有的换行符和空格为逗号，然后按逗号拆分
                                result['需求ID Requirement ID'] = result['需求ID Requirement ID'].astype(str).str.replace('\n', ',').str.replace(' ', ',')
                                # 拆分需求ID
                                result = result.assign(**{'需求ID Requirement ID': result['需求ID Requirement ID'].str.split(',')})
                                result = result.explode('需求ID Requirement ID')
                                # 去除空白
                                result['需求ID Requirement ID'] = result['需求ID Requirement ID'].str.strip()
                                # 过滤掉空的需求ID
                                result = result[result['需求ID Requirement ID'] != '']
                                
                                # 打印处理后的数据预览
                                print("处理需求ID后的数据预览:")
                                print(result.head())
                                print(f"处理后的数据行数: {len(result)}")
                            
                            # 只添加非空数据，即使计划否和测试结果列没有数据
                            if not result.empty:
                                # 确保所有必要的列都存在
                                required_columns = ['FileName', '用例ID CaseID', '需求ID Requirement ID', '计划否Plan or not', '测试结果Test Results']
                                for col in required_columns:
                                    if col not in result.columns:
                                        result[col] = ''
                                # 调整列顺序
                                result = result[required_columns]
                                all_sheet_data.append(result)
                                print(f"添加了 {len(result)} 条数据")
                            else:
                                print("警告: 过滤后没有数据")
                            break  # 找到合适的表头后，停止尝试其他行
                        except Exception as e:
                            print(f"处理数据时出错: {str(e)}")
                            import traceback
                            traceback.print_exc()
                            continue
                except Exception as e:
                    print(f"尝试第 {header_row+1} 行作为表头时出错: {str(e)}")
        
        # 如果找到数据，合并所有工作表的数据
        if all_sheet_data:
            return pd.concat(all_sheet_data, ignore_index=True)
        else:
            # 如果所有尝试都失败
            print(f"文件 {file_name} 无法找到必要的列")
            return pd.DataFrame()
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def merge_all_data(folder_path):
    """合并所有Excel文件的数据"""
    excel_files = get_excel_files(folder_path)
    all_data = []
    
    if not excel_files:
        print("未找到Excel文件")
        return pd.DataFrame()
    
    print(f"找到 {len(excel_files)} 个Excel文件")
    
    for file in excel_files:
        data = process_excel_file(file)
        if not data.empty:
            all_data.append(data)
    
    if all_data:
        merged_df = pd.concat(all_data, ignore_index=True)
        # 重新排列列顺序，将文件名放在第一列
        columns = ['FileName', '用例ID CaseID', '需求ID Requirement ID', '计划否Plan or not', '测试结果Test Results']
        # 确保所有必要的列都存在
        for col in columns:
            if col not in merged_df.columns:
                merged_df[col] = ''
        merged_df = merged_df[columns]
        return merged_df
    else:
        return pd.DataFrame()

def get_feishu_token(app_id, app_secret):
    """获取飞书API令牌"""
    url = "https://open.feishu.cn/open-apis/auth/v3/app_access_token/internal/"
    headers = {"Content-Type": "application/json"}
    data = {
        "app_id": app_id,
        "app_secret": app_secret
    }
    response = requests.post(url, headers=headers, json=data)
    if response.status_code == 200:
        return response.json().get("app_access_token")
    else:
        print(f"获取飞书令牌失败: {response.text}")
        return None

def export_to_feishu(merged_df, feishu_url):
    """将数据导入到飞书多维表格"""
    try:
        # 打印数据预览
        print("数据预览:")
        print(merged_df.head())
        print(f"总共 {len(merged_df)} 条记录")
        print(f"准备导入到飞书多维表格: {feishu_url}")
        
        # 从API地址中提取app_token和table_id
        # 格式: https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records
        import re
        match = re.search(r'apps/(.*?)/tables/(.*?)/records', feishu_url)
        if not match:
            print("无效的飞书多维表格API地址格式")
            return False
        
        app_token = match.group(1)
        table_id = match.group(2)
        
        # 飞书AppID和secret
        app_id = "cli_a948b44ee4f99cd1"
        app_secret = "3Yb5DsFHxuXnv1DhEsvfPhTvBJtvUxsQ"
        
        # 创建client
        client = lark.Client.builder() \
            .app_id(app_id) \
            .app_secret(app_secret) \
            .log_level(lark.LogLevel.DEBUG) \
            .build()
        
        # 处理NaN值，转换为空字符串
        def process_value(value):
            if pd.isna(value):
                return ""
            elif isinstance(value, float) and (value == float('inf') or value == float('-inf')):
                return ""
            else:
                return str(value)
        
        # 准备导入数据
        success_count = 0
        total_count = len(merged_df)  # 导入全部数据
        
        for i in range(total_count):
            row = merged_df.iloc[i]
            
            # 构建字段数据
            fields = {
                "FileName": process_value(row["FileName"]),
                "STCaseID": process_value(row["用例ID CaseID"]),
                "StoryID": process_value(row["需求ID Requirement ID"]),
                "计划否": process_value(row["计划否Plan or not"]),
                "测试结果": process_value(row["测试结果Test Results"])
            }
            
            # 确保字段不为空
            if not any(fields.values()):
                continue
            
            # 构造请求对象
            request: CreateAppTableRecordRequest = CreateAppTableRecordRequest.builder() \
                .app_token(app_token) \
                .table_id(table_id) \
                .user_id_type("open_id") \
                .request_body(AppTableRecord.builder()
                    .fields(fields)
                    .build()) \
                .build()
            
            # 发起请求
            response: CreateAppTableRecordResponse = client.bitable.v1.app_table_record.create(request)
            
            # 处理失败返回
            if not response.success():
                lark.logger.error(
                    f"导入第 {i+1} 条记录失败, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
                print(f"导入第 {i+1} 条记录失败: {response.msg}")
            else:
                print(f"成功导入第 {i+1} 条记录")
                success_count += 1
        
        print(f"导入完成: 成功 {success_count} 条, 失败 {total_count - success_count} 条")
        return success_count > 0
        
    except Exception as e:
        print(f"导入到飞书时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def main(folder_path, feishu_url):
    """主函数"""
    if not folder_path:
        print("请设置文件夹路径")
        return
    
    if not os.path.exists(folder_path):
        print(f"文件夹不存在: {folder_path}")
        return
    
    # 合并数据
    merged_df = merge_all_data(folder_path)
    
    if merged_df.empty:
        print("没有找到有效数据")
        return
    
    # 导入到飞书多维表格
    if feishu_url:
        # 检查是否是API地址
        if "open.feishu.cn/open-apis" not in feishu_url:
            print("警告: 您提供的可能是飞书多维表格的访问URL，而不是API接口URL")
            print("正确的API接口URL格式应为: https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records")
            print("请参考飞书开放平台文档获取正确的API接口URL")
        
        success = export_to_feishu(merged_df, feishu_url)
        if success:
            print("成功导入到飞书多维表格")
        else:
            print("导入到飞书多维表格失败")
    else:
        print("未设置飞书多维表格API地址，跳过导入步骤")
        print("提示: 飞书多维表格API地址格式应为: https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records")
        print("请参考飞书开放平台文档获取正确的API接口URL")

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
