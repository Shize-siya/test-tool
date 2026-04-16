#!/usr/bin/python
# -*- coding:UTF-8 -*-

from math import e
import xml.etree.ElementTree as ET
import argparse
import os
import re
import pandas as pd  # 新增导入

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from feishu import Feishu

def parse_config(elem, list_type):
    app_token = elem.find('app_token').text
    list_config = elem.find(list_type)
    table_id = list_config.find('table_id').text
    view_id = list_config.find('view_id').text
    return Feishu(app_token, table_id, view_id)

def do_or_list(read, write):
    fs_read = parse_config(read, 'or_list')
    or_list = fs_read.get_or_list()

    fs_write = parse_config(write, 'or_list')
    fs_write.clear_list()

    print(f'共解析到 {len(or_list)} 条需求记录')
    fs_write.insert_or_list(or_list)

def do_feature_list(read, write):
    fs_read = parse_config(read, 'feature_list')
    feature_list = fs_read.get_feature_list()

    fs_write = parse_config(write, 'feature_list')
    fs_write.clear_list()

    print(f'共解析到 {len(feature_list)} 条需求记录')
    fs_write.insert_feature_list(feature_list)

def do_prd(read, write):

    fs_write = parse_config(write, 'prd')
    fs_write.clear_list()

    # 遍历当前目录
    base_path = read.find('prd').find('path').text
    for parent, dirnames, filenames in os.walk(base_path):
        prd_list = []
        for filename in filenames:
            # 过滤已打开文件 ~$开头
            if filename[0:2] == '~$':
                continue
            
            # 过滤docx文件
            filetype = filename.split(".")[-1]
            if filetype != 'docx':
                continue
        
            rpath = os.path.join(parent, filename)
            print('读取文件：' + filename)

            # 读取docx文件
            doc = Document(rpath)

            # 遍历表格
            prd_list = get_tables_with_headings(doc)
            if len(prd_list) > 0:
                print(f'共解析到 {len(prd_list)} 条PRD记录')
                # 注意：insert_prd_list 需要第二个参数 filename，这里传入最后一个处理的文件名
                fs_write.insert_prd_list(prd_list, filename)

def do_code(read, write):
    code_list = []

    fs_write = parse_config(write, 'code')
    fs_write.clear_list()

    # 遍历Jacoco生成的HTML目录
    base_path = read.find('code').find('path').text
    for parent, dirnames, filenames in os.walk(base_path):
        for filename in filenames:
            # 只处理.html文件，排除index.html和jacoco-sessions.html
            if not filename.endswith('.html') or filename in ['index.html', 'jacoco-sessions.html']:
                continue
            
            # 排除source文件（包含源码的行号）
            if filename.endswith('.source.html'):
                continue
            
            rpath = os.path.join(parent, filename)
            # print('解析文件：' + filename)
            
            # 读取HTML文件
            with open(rpath, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 提取package名称
            package_match = re.search(r'class="el_package">(.*?)</a>', content)
            package = package_match.group(1) if package_match else ''
            
            # 提取class名称
            class_match = re.search(r'<span class="el_class">(.*?)</span>', content)
            classname = class_match.group(1) if class_match else ''
            
            # 提取所有method名称
            method_pattern = r'class="el_method"[^>]*>([^<]+)</a>'
            methods = re.findall(method_pattern, content)
            
            # 如果没有method信息，从文件名提取
            if not methods and filename.endswith('.html') and not filename.endswith('.kt.html'):
                # 对于纯class页面，提取方法名
                method_matches = re.findall(r'<a href="[^"]+\.kt\.html#L(\d+)"[^>]*class="el_method"[^>]*>([^<]+)</a>', content)
                for _, method_name in method_matches:
                    methods.append(method_name)
            
            # 如果仍没有方法，尝试从kt.html文件提取
            if not methods:
                kt_file = rpath.replace('.html', '.kt.html')
                if os.path.exists(kt_file):
                    with open(kt_file, 'r', encoding='utf-8') as f:
                        kt_content = f.read()
                    method_pattern = r'<a[^>]*href="#L(\d+)"[^>]*>([^<]+)</a>'
                    methods = re.findall(method_pattern, kt_content)
                    methods = [m[1] for m in methods]  # 只取方法名
            
            # 添加到列表
            for method in methods:
                code_list.append([package, classname, method])
    
    print(f'共解析到 {len(code_list)} 条代码记录')
    fs_write.insert_code_list(code_list)

def do_ut(read, write):
    ut_list = []

    fs_write = parse_config(write, 'ut')
    fs_write.clear_list()

    # 遍历Kotlin代码库
    base_path = read.find('ut').find('path').text
    for parent, dirnames, filenames in os.walk(base_path):
        for filename in filenames:
            # 只处理Test.kt结尾的文件
            if not filename.endswith('Test.kt'):
                continue

            # 相对路径
            relative_path = os.path.relpath(parent, base_path)
            
            # 读取文件内容
            rpath = os.path.join(parent, filename)
            with open(rpath, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 提取所有@Test注解的函数
            # 匹配 @Test 注解后的 fun 函数定义
            test_method_pattern = r'@Test[^\n]*\n\s*fun\s+(\w+)\s*\('
            matches = re.findall(test_method_pattern, content)
            
            for utc_case_name in matches:
                # 提取MethodName: 从函数名截取test_之后到第一个_之前的内容
                # 例如: test_getFirstLocMatchInfo_withNullMatchInfo_returnsNull -> getFirstLocMatchInfo
                method_name = utc_case_name
                if utc_case_name.startswith('test_'):
                    # 去掉test_前缀，然后取第一个_之前的内容
                    method_part = utc_case_name[5:]  # 去掉test_
                    underscore_idx = method_part.find('_')
                    if underscore_idx > 0:
                        method_name = method_part[:underscore_idx]

                        ut_list.append([relative_path, filename, utc_case_name, method_name])
    
    print(f'共解析到 {len(ut_list)} 条UT记录')
    fs_write.insert_ut_list(ut_list)

def get_excel_files(folder_path):
    """获取文件夹中所有Excel文件"""
    excel_files = []
    for file in os.listdir(folder_path):
        if file.startswith('~$'):
            continue
        if file.endswith(('.xlsx', '.xls', '.xlsm')):
            excel_files.append(os.path.join(folder_path, file))
    return excel_files

def process_excel_file(file_path):
    """处理单个Excel文件，提取所需字段，返回DataFrame"""
    try:
        file_name = os.path.basename(file_path)
        xl = pd.ExcelFile(file_path)
        all_sheet_data = []

        for sheet_name in xl.sheet_names:
            if 'TestCase' not in sheet_name:
                continue

            # 特殊处理 TestCase-EV 工作表
            if 'TestCase-EV' in sheet_name:
                try:
                    for header_row in range(5):
                        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
                        df = df.dropna(axis=1, how='all')
                        if len(df.columns) >= 3:
                            result = pd.DataFrame()
                            for case_col in [1, 2]:
                                try:
                                    result['用例ID CaseID'] = df.iloc[:, case_col]
                                    req_col = case_col + 1
                                    if req_col < len(df.columns):
                                        result['需求ID Requirement ID'] = df.iloc[:, req_col]
                                    else:
                                        result['需求ID Requirement ID'] = ''
                                    if len(df.columns) >= 29:
                                        result['计划否Plan or not'] = df.iloc[:, 27]
                                        result['测试结果Test Results'] = df.iloc[:, 28]
                                    else:
                                        result['计划否Plan or not'] = ''
                                        result['测试结果Test Results'] = ''
                                    result['FileName'] = file_name
                                    result['工作表'] = sheet_name

                                    result = result.dropna(subset=['用例ID CaseID'])
                                    result = result[~result['用例ID CaseID'].astype(str).str.strip().isin(
                                        ['用例ID', 'CaseID', '用例ID CaseID', 'Case ID', 'Test Case ID'])]
                                    result = result[result['用例ID CaseID'].astype(str).str.contains('TestCase', case=False, na=False)]

                                    if '需求ID Requirement ID' in result.columns:
                                        result['需求ID Requirement ID'] = result['需求ID Requirement ID'].astype(str).str.replace('\n', ',').str.replace(' ', ',')
                                        result = result.assign(**{'需求ID Requirement ID': result['需求ID Requirement ID'].str.split(',')})
                                        result = result.explode('需求ID Requirement ID')
                                        result['需求ID Requirement ID'] = result['需求ID Requirement ID'].str.strip()
                                        result = result[result['需求ID Requirement ID'] != '']

                                    if not result.empty:
                                        required = ['FileName', '用例ID CaseID', '需求ID Requirement ID', '计划否Plan or not', '测试结果Test Results']
                                        for col in required:
                                            if col not in result.columns:
                                                result[col] = ''
                                        all_sheet_data.append(result[required])
                                        break
                                except:
                                    continue
                            if not result.empty:
                                break
                except Exception as e:
                    print(f"处理TestCase-EV工作表出错: {e}")
                continue

            # 普通TestCase工作表
            for header_row in range(10):
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
                    df = df.dropna(axis=1, how='all')
                    if len(df.columns) < 4:
                        continue

                    # 列名匹配
                    known = {
                        'CaseID': ['用例ID\nCaseID', '用例ID CaseID', 'CaseID', '用例ID', 'Case ID', 'Test Case ID'],
                        '需求ID Requirement ID': ['需求ID\nRequirement ID', '需求ID Requirement ID', '需求ID', 'Requirement ID', '需求编号'],
                        '计划否Plan or not': ['计划否\nPlan or not', '计划否Plan or not', '计划否', 'Plan or not'],
                        '测试结果Test Results': ['测试结果\nTest Results', '测试结果Test Results', '测试结果', 'Test Results']
                    }
                    actual = {}
                    for target, candidates in known.items():
                        for col in candidates:
                            if col in df.columns:
                                actual[target] = col
                                break
                    if len(actual) == 4:
                        result = df[[actual['CaseID'], actual['需求ID Requirement ID'], actual['计划否Plan or not'], actual['测试结果Test Results']]].copy()
                        rename = {
                            'CaseID': '用例ID CaseID',
                            '需求ID Requirement ID': '需求ID Requirement ID',
                            '计划否Plan or not': '计划否Plan or not',
                            '测试结果Test Results': '测试结果Test Results'
                        }
                        result.columns = [rename[k] for k in actual.keys()]
                        result['FileName'] = file_name
                        result['工作表'] = sheet_name

                        result = result.dropna(subset=['用例ID CaseID'])
                        result = result[~result['用例ID CaseID'].astype(str).str.strip().isin(
                            ['用例ID', 'CaseID', '用例ID CaseID', 'Case ID', 'Test Case ID'])]

                        if '需求ID Requirement ID' in result.columns:
                            result['需求ID Requirement ID'] = result['需求ID Requirement ID'].astype(str).str.replace('\n', ',').str.replace(' ', ',')
                            result = result.assign(**{'需求ID Requirement ID': result['需求ID Requirement ID'].str.split(',')})
                            result = result.explode('需求ID Requirement ID')
                            result['需求ID Requirement ID'] = result['需求ID Requirement ID'].str.strip()
                            result = result[result['需求ID Requirement ID'] != '']

                        if not result.empty:
                            required = ['FileName', '用例ID CaseID', '需求ID Requirement ID', '计划否Plan or not', '测试结果Test Results']
                            for col in required:
                                if col not in result.columns:
                                    result[col] = ''
                            all_sheet_data.append(result[required])
                            break
                except:
                    continue

        if all_sheet_data:
            return pd.concat(all_sheet_data, ignore_index=True)
        else:
            return pd.DataFrame()
    except Exception as e:
        print(f"处理文件 {file_path} 出错: {e}")
        return pd.DataFrame()

def merge_all_data(folder_path):
    """合并所有Excel文件的数据"""
    excel_files = get_excel_files(folder_path)
    if not excel_files:
        print("未找到Excel文件")
        return pd.DataFrame()
    all_data = []
    for file in excel_files:
        data = process_excel_file(file)
        if not data.empty:
            all_data.append(data)
    if all_data:
        return pd.concat(all_data, ignore_index=True)
    else:
        return pd.DataFrame()

def do_excel(read, write):
    """解析Excel测试用例并写入飞书多维表格"""
    # 从read配置中获取Excel文件夹路径
    excel_node = read.find('excel')
    if excel_node is None or excel_node.find('path') is None:
        print("配置文件中未找到 <read><excel><path> 节点，跳过Excel处理")
        return
    folder_path = excel_node.find('path').text
    if not os.path.exists(folder_path):
        print(f"Excel文件夹不存在: {folder_path}")
        return

    # 从write配置中获取飞书写入对象
    fs_write = parse_config(write, 'excel')
    fs_write.clear_list()

    # 解析Excel
    merged_df = merge_all_data(folder_path)
    if merged_df.empty:
        print("没有从Excel中提取到有效数据")
        return

    # 转换为列表格式 [FileName, STCaseID, StoryID, 计划否, 测试结果]
    data_list = []
    for _, row in merged_df.iterrows():
        data_list.append([
            str(row['FileName']) if pd.notna(row['FileName']) else '',
            str(row['用例ID CaseID']) if pd.notna(row['用例ID CaseID']) else '',
            str(row['需求ID Requirement ID']) if pd.notna(row['需求ID Requirement ID']) else '',
            str(row['计划否Plan or not']) if pd.notna(row['计划否Plan or not']) else '',
            str(row['测试结果Test Results']) if pd.notna(row['测试结果Test Results']) else ''
        ])

    print(f"共解析到 {len(data_list)} 条Excel测试记录")
    fs_write.insert_excel_list(data_list)


def get_tables_with_headings(doc):
    results = []
    """获取文档中所有标题"""
    headings = []

    # 遍历表格
    for idx, element in enumerate(doc.element.body):
        if element.tag.endswith('p'):
            para = Paragraph(element, doc)
            # 判断段落样式是否为标题
            if para.style.name.startswith('Heading'):
                headings.append({
                    'index': idx,
                    'level': int(para.style.name.split()[-1]) if para.style.name.split() else 0,  # 标题级别数字
                    'text': para.text.strip()      # 标题文本
                })

        elif element.tag.endswith('tbl'):
            table = Table(element, doc)
            
            # 第一行第一列
            first_row = table.rows[0]
            cells = first_row.cells

            # 检查表头，过滤表格
            if cells[0].text != 'User Story ID':
                continue

            # 遍历表格，从第二行开始
            for row in table.rows[1:]:
                list = []
                list.append(row.cells[0].text)
                list.append(row.cells[1].text)
                list.append(row.cells[2].text)
                list.append(row.cells[3].text)

                # 遍历标题，查找所有父级段落，保存到list中
                pre_level = 0
                text_list = []
                for i in range(len(headings)):
                    h = headings[i]

                    if h['level'] == 1:
                        text_list = [h['text']]
                    else:
                        for j in range(pre_level - h['level'] + 1):
                            text_list.pop()
                        text_list.append(h['text'])

                    pre_level = h['level']

                for text in text_list:
                    list.append(text)

                # 添加到结果列表
                results.append(list)
    
    return results

def do_itcase(read, write):
    """解析Excel测试用例（时序ID HLD_SequenceName）并写入飞书多维表格
    读取第18列（索引17）作为时序ID，字段名为“时序ID HLD_SequenceName”
    """
    # 从read配置中获取Excel文件夹路径（使用itcase节点）
    itcase_node = read.find('itcase')
    if itcase_node is None or itcase_node.find('path') is None:
        print("配置文件中未找到 <read><itcase><path> 节点，跳过Excel处理")
        return
    folder_path = itcase_node.find('path').text
    if not os.path.exists(folder_path):
        print(f"Excel文件夹不存在: {folder_path}")
        return

    fs_write = parse_config(write, 'itcase')  
    fs_write.clear_list()
    
    def get_excel_files_itcase(folder_path):
        excel_files = []
        for file in os.listdir(folder_path):
            if file.startswith('~$'):
                continue
            if file.endswith(('.xlsx', '.xls', '.xlsm')):
                excel_files.append(os.path.join(folder_path, file))
        return excel_files

    def process_excel_file_itcase(file_path):
        try:
            file_name = os.path.basename(file_path)
            xl = pd.ExcelFile(file_path)
            all_sheet_data = []

            for sheet_name in xl.sheet_names:
                if 'TestCase' not in sheet_name:
                    continue

                # 特殊处理 TestCase-EV 工作表
                if 'TestCase-EV' in sheet_name:
                    try:
                        for header_row in range(5):
                            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
                            df = df.dropna(axis=1, how='all')
                            if len(df.columns) >= 3:
                                result = pd.DataFrame()
                                for case_col in [1, 2]:
                                    try:
                                        result['用例ID CaseID'] = df.iloc[:, case_col]
                                        # 时序ID：第13列（索引12）
                                        if len(df.columns) >= 12:
                                            result['时序ID HLD_SequenceName'] = df.iloc[:, 11]
                                        else:
                                            result['时序ID HLD_SequenceName'] = ''
                                        if len(df.columns) >= 29:
                                            result['计划否Plan or not'] = df.iloc[:, 27]
                                            result['测试结果Test Results'] = df.iloc[:, 28]
                                        else:
                                            result['计划否Plan or not'] = ''
                                            result['测试结果Test Results'] = ''
                                        result['FileName'] = file_name
                                        result['工作表'] = sheet_name

                                        result = result.dropna(subset=['用例ID CaseID'])
                                        result = result[~result['用例ID CaseID'].astype(str).str.strip().isin(
                                            ['用例ID', 'CaseID', '用例ID CaseID', 'Case ID', 'Test Case ID'])]
                                        result = result[result['用例ID CaseID'].astype(str).str.contains('TestCase', case=False, na=False)]

                                        if '时序ID HLD_SequenceName' in result.columns:
                                            result['时序ID HLD_SequenceName'] = result['时序ID HLD_SequenceName'].astype(str).str.replace('\n', ',')
                                            result = result.assign(**{'时序ID HLD_SequenceName': result['时序ID HLD_SequenceName'].str.split(',')})
                                            result = result.explode('时序ID HLD_SequenceName')
                                            result['时序ID HLD_SequenceName'] = result['时序ID HLD_SequenceName'].str.strip()
                                            result = result[result['时序ID HLD_SequenceName'] != '']

                                        if not result.empty:
                                            required = ['FileName', '用例ID CaseID', '时序ID HLD_SequenceName', '计划否Plan or not', '测试结果Test Results']
                                            for col in required:
                                                if col not in result.columns:
                                                    result[col] = ''
                                            all_sheet_data.append(result[required])
                                            break
                                    except:
                                        continue
                                if not result.empty:
                                    break
                    except Exception as e:
                        print(f"处理TestCase-EV工作表出错: {e}")
                    continue

                # 普通 TestCase 工作表
                for header_row in range(10):
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
                        df = df.dropna(axis=1, how='all')
                        if len(df.columns) < 4:
                            continue

                        known = {
                            'CaseID': ['用例ID\nCaseID', '用例ID CaseID', 'CaseID', '用例ID', 'Case ID', 'Test Case ID'],
                            '计划否Plan or not': ['计划否\nPlan or not', '计划否Plan or not', '计划否', 'Plan or not'],
                            '测试结果Test Results': ['测试结果\nTest Results', '测试结果Test Results', '测试结果', 'Test Results']
                        }
                        actual = {}
                        for target, candidates in known.items():
                            for col in candidates:
                                if col in df.columns:
                                    actual[target] = col
                                    break
                        if len(actual) == 3:
                            result = df[[actual['CaseID'], actual['计划否Plan or not'], actual['测试结果Test Results']]].copy()
                            # 添加时序ID列
                            if len(df.columns) >= 12:
                                result['时序ID HLD_SequenceName'] = df.iloc[:, 11]
                            else:
                                result['时序ID HLD_SequenceName'] = ''
                            rename = {
                                'CaseID': '用例ID CaseID',
                                '计划否Plan or not': '计划否Plan or not',
                                '测试结果Test Results': '测试结果Test Results'
                            }
                            result.columns = [rename[k] for k in actual.keys()] + ['时序ID HLD_SequenceName']
                            result['FileName'] = file_name
                            result['工作表'] = sheet_name

                            result = result.dropna(subset=['用例ID CaseID'])
                            result = result[~result['用例ID CaseID'].astype(str).str.strip().isin(
                                ['用例ID', 'CaseID', '用例ID CaseID', 'Case ID', 'Test Case ID'])]

                            if '时序ID HLD_SequenceName' in result.columns:
                                result['时序ID HLD_SequenceName'] = result['时序ID HLD_SequenceName'].astype(str).str.replace('\n', ',')
                                result = result.assign(**{'时序ID HLD_SequenceName': result['时序ID HLD_SequenceName'].str.split(',')})
                                result = result.explode('时序ID HLD_SequenceName')
                                result['时序ID HLD_SequenceName'] = result['时序ID HLD_SequenceName'].str.strip()
                                result = result[result['时序ID HLD_SequenceName'] != '']

                            if not result.empty:
                                required = ['FileName', '用例ID CaseID', '时序ID HLD_SequenceName', '计划否Plan or not', '测试结果Test Results']
                                for col in required:
                                    if col not in result.columns:
                                        result[col] = ''
                                all_sheet_data.append(result[required])
                                break
                    except:
                        continue

            if all_sheet_data:
                return pd.concat(all_sheet_data, ignore_index=True)
            else:
                return pd.DataFrame()
        except Exception as e:
            print(f"处理文件 {file_path} 出错: {e}")
            return pd.DataFrame()

    def merge_all_data_itcase(folder_path):
        excel_files = get_excel_files_itcase(folder_path)
        if not excel_files:
            print("未找到Excel文件")
            return pd.DataFrame()
        all_data = []
        for file in excel_files:
            data = process_excel_file_itcase(file)
            if not data.empty:
                all_data.append(data)
        if all_data:
            return pd.concat(all_data, ignore_index=True)
        else:
            return pd.DataFrame()

    merged_df = merge_all_data_itcase(folder_path)
    if merged_df.empty:
        print("没有从Excel中提取到有效数据（时序ID模式）")
        return

    data_list = []
    for _, row in merged_df.iterrows():
        data_list.append([
            str(row['FileName']) if pd.notna(row['FileName']) else '',
            str(row['用例ID CaseID']) if pd.notna(row['用例ID CaseID']) else '',
            str(row['时序ID HLD_SequenceName']) if pd.notna(row['时序ID HLD_SequenceName']) else '',
            str(row['计划否Plan or not']) if pd.notna(row['计划否Plan or not']) else '',
            str(row['测试结果Test Results']) if pd.notna(row['测试结果Test Results']) else ''
        ])

    print(f"共解析到 {len(data_list)} 条Excel测试记录（时序ID模式）")
    fs_write.insert_it_list(data_list)

# 读取配置文件
def read_config():
    tree = ET.parse('config.xml')
    root = tree.getroot()
    return root.find('read'), root.find('write')

if __name__ == "__main__":
    read, write = read_config()

    # 优先从命令行读取，否则交互式输入
    parser = argparse.ArgumentParser()
    parser.add_argument('--action', choices=['0', '1', '2', '3', '4', '5', '6', '7'], help='操作: 0=ALL, 1=ORlist, 2=Featurelist, 3=PRD, 4=Code, 5=UT, 6=STcase, 7=IT')
    args = parser.parse_args()

    if args.action:
        action = args.action
    else:
        # 等待接收用户输入
        action = input("请输入要执行的操作  0：ALL  1：ORlist  2：Featurelist  3：PRD  4：Code  5：UT  6：STcase, 7：IT")

    if action == "0":
        print("执行ORlist")
        do_or_list(read, write)
        print("执行Featurelist")
        do_feature_list(read, write)
        print("执行PRD")
        do_prd(read, write)
        print("执行Code")
        do_code(read, write)
        print("执行UT")
        do_ut(read, write)
        print("执行STcase")
        do_excel(read, write)
        print("执行IT")
        do_itcase(read, write)
    elif action == "1":
        do_or_list(read, write)
    elif action == "2":
        do_feature_list(read, write)
    elif action == "3":
        do_prd(read, write)
    elif action == "4":
        do_code(read, write)
    elif action == "5":
        do_ut(read, write)
    elif action == "6":
        do_excel(read, write)
    elif action == "7":
        do_itcase(read, write)
    else:
        print("输入错误")