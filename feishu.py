#!/usr/bin/python
# -*- coding:UTF-8 -*-

import xml.etree.ElementTree as ET
import json
import requests

import lark_oapi as lark
from lark_oapi.api.bitable.v1 import *

class Feishu:
    def __init__(self, app_token, table_id, view_id):

        self.app_id = "cli_a948b44ee4f99cd1"
        self.app_secret = "3Yb5DsFHxuXnv1DhEsvfPhTvBJtvUxsQ"

        self.app_token = app_token
        self.table_id = table_id
        self.view_id = view_id

        # 创建client
        self.client = lark.Client.builder() \
            .app_id(self.app_id) \
            .app_secret(self.app_secret) \
            .log_level(lark.LogLevel.ERROR) \
            .build()

    # 通用的批量插入方法
    def _batch_insert(self, data_list, field_map):
        """
        通用批量插入方法
        :param data_list: 数据列表，每个元素是列表或元组
        :param field_map: 字段映射字典，如 {"字段名": lambda record, idx: record[idx]}
        """
        page_size = 500
        for i in range(0, len(data_list), page_size):
            batch = data_list[i:i + page_size]
            records = []
            for record in batch:
                fields = {}
                for field_name, extractor in field_map.items():
                    try:
                        fields[field_name] = extractor(record)
                    except IndexError:
                        fields[field_name] = ''
                records.append(AppTableRecord.builder().fields(fields).build())

            request = BatchCreateAppTableRecordRequest.builder() \
                .app_token(self.app_token) \
                .table_id(self.table_id) \
                .request_body(BatchCreateAppTableRecordRequestBody.builder()
                    .records(records)
                    .build()) \
                .build()

            response: BatchCreateAppTableRecordResponse = self.client.bitable.v1.app_table_record.batch_create(request)

            if response.success():
                print(f"已写入 {min(i + page_size, len(data_list))}/{len(data_list)} 条记录")
            else:
                print(f"写入失败: {response.code} - {response.msg}")

    # 通用的分页查询方法
    def _search_records(self, field_names=None):
        """
        通用分页查询方法
        :param field_names: 要查询的字段列表，None表示查询所有字段
        :return: 记录列表
        """
        records = []
        page_token = None
        has_more = True

        while has_more:
            if page_token is None:
                request = SearchAppTableRecordRequest.builder() \
                    .app_token(self.app_token) \
                    .table_id(self.table_id) \
                    .page_size(500) \
                    .request_body(SearchAppTableRecordRequestBody.builder()
                        .view_id(self.view_id)
                        .field_names(field_names)
                        .build()) \
                    .build()
            else:
                request = SearchAppTableRecordRequest.builder() \
                    .app_token(self.app_token) \
                    .table_id(self.table_id) \
                    .page_token(page_token) \
                    .page_size(500) \
                    .request_body(SearchAppTableRecordRequestBody.builder()
                        .view_id(self.view_id)
                        .field_names(field_names)
                        .build()) \
                    .build()

            response: SearchAppTableRecordResponse = self.client.bitable.v1.app_table_record.search(request)

            if not response.success():
                lark.logger.error(
                    f"client.bitable.v1.app_table_record.search failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}, resp: \n{json.dumps(json.loads(response.raw.content), indent=4, ensure_ascii=False)}")
                return []

            data = lark.JSON.marshal(response.data, indent=4)
            res = json.loads(data)
            has_more = res['has_more']
            if has_more:
                page_token = res['page_token']

            records.extend(res['items'])

        return records

    # 清空数据表
    def clear_list(self):
        all_records = []
        items = self._search_records()
        all_records.extend([item['record_id'] for item in items])

        page_size = 500
        for i in range(0, len(all_records), page_size):
            del_records = all_records[i:i + page_size]
            
            request = BatchDeleteAppTableRecordRequest.builder() \
                .app_token(self.app_token) \
                .table_id(self.table_id) \
                .request_body(BatchDeleteAppTableRecordRequestBody.builder()
                    .records(del_records)
                    .build()) \
                .build()
            
            response: BatchDeleteAppTableRecordResponse = self.client.bitable.v1.app_table_record.batch_delete(request)
            
            if response.success():
                print(f"已删除 {min(i + page_size, len(all_records))}/{len(all_records)} 条记录")
            else:
                print(f"删除失败: {response.code} - {response.msg}")

    # 获取ORlist
    def get_or_list(self):
        orlist = []
        items = self._search_records(["RFQ ID", "Featurelist text", "OR Status", "Category"])

        for item in items:
            if 'Featurelist text' in item['fields']:
                for feature in item['fields']['Featurelist text']['value'][0]['text'].split(','):
                    orlist.append([
                        item['fields']['RFQ ID']['value'][0]['text'],
                        feature,
                        item['fields']['OR Status'],
                        item['fields']['Category']
                    ])
            else:
                orlist.append([
                    item['fields']['RFQ ID']['value'][0]['text'],
                    "",
                    item['fields']['OR Status'],
                    item['fields']['Category']
                ])

        return orlist

    # 获取Featurelist
    def get_feature_list(self):
        featurelist = []
        field_names = ["Feature ID", "RFQ ID text", "JIRA ID", "1st Level", "2nd Level", "Reference PRD", "Reference UE", "UE ScreenID", "Release", "Phases", "Coding"]
        items = self._search_records(field_names)

        for item in items:
            fields = item['fields']
            record = [
                fields.get('Feature ID', {}).get('value', [{}])[0].get('text', ''),
                fields.get('RFQ ID text', {}).get('value', [{}])[0].get('text', '') if 'RFQ ID text' in fields else '',
                fields.get('JIRA ID', {}).get('value', [{}])[0].get('text', '') if 'JIRA ID' in fields else '',
                fields.get('1st Level', ''),
                fields.get('2nd Level', ''),
                ' '.join([prd['text'] for prd in fields['Reference PRD']]) if 'Reference PRD' in fields else '',
                fields.get('Reference UE', [{}])[0].get('text', '') if 'Reference UE' in fields else '',
                fields.get('UE ScreenID', [{}])[0].get('text', '') if 'UE ScreenID' in fields else '',
                fields.get('Release', ''),
                fields.get('Phases', ''),
                fields.get('Coding', '')
            ]
            featurelist.append(record)

        return featurelist

    # 写入ORlist
    def insert_or_list(self, orlist):
        self._batch_insert(orlist, {
            "RFQID": lambda r: r[0],
            "FeatureID": lambda r: r[1],
            "OR Status": lambda r: r[2],
            "Category": lambda r: r[3]
        })

    # 写入Featurelist
    def insert_feature_list(self, featurelist):
        self._batch_insert(featurelist, {
            "FeatureID": lambda r: r[0],
            "RFQID": lambda r: r[1],
            "JIRAID": lambda r: r[2],
            "1 st Level": lambda r: r[3],
            "2 nd Level": lambda r: r[4],
            "Reference PRD": lambda r: r[5],
            "Reference UE": lambda r: r[6],
            "UE ScreenID": lambda r: r[7],
            "Release": lambda r: r[8],
            "Phases": lambda r: r[9],
            "Coding": lambda r: r[10]
        })

    # 写入PRD
    def insert_prd_list(self, prdlist, filename):
        self._batch_insert(prdlist, {
            "文档名": lambda r: filename,
            "User Story ID": lambda r: r[0],
            "RFQ ID": lambda r: r[1],
            "Feature Name": lambda r: r[2],
            "JIRA-ID": lambda r: r[3],
            "一级章节名": lambda r: r[4] if len(r) > 4 else '',
            "二级章节名": lambda r: r[5] if len(r) > 5 else '',
            "三级章节名": lambda r: r[6] if len(r) > 6 else '',
            "四级章节名": lambda r: r[7] if len(r) > 7 else '',
            "五级章节名": lambda r: r[8] if len(r) > 8 else ''
        })

    # 写入Code列表
    def insert_code_list(self, codelist):
        self._batch_insert(codelist, {
            "Core": lambda r: 'EngineCore',
            "Package": lambda r: r[0],
            "ClassName": lambda r: r[1],
            "MethodName": lambda r: r[2]
        })

    # 写入UT列表
    def insert_ut_list(self, utlist):
        self._batch_insert(utlist, {
            "Path": lambda r: r[0],
            "FileName": lambda r: r[1],
            "UTCaseName": lambda r: r[2],
            "MethodName": lambda r: r[3]
        })

    # ========== 新增：写入Excel测试用例列表 ==========
    def insert_excel_list(self, excel_list):
        self._batch_insert(excel_list, {
            "FileName": lambda r: r[0],
            "STCaseID": lambda r: r[1],
            "StoryID": lambda r: r[2],
            "计划否": lambda r: r[3],
            "测试结果": lambda r: r[4]
        })