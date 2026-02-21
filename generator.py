#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel直接生成Draw.io格式转换器 - 完整还原原版功能
包含：投资者分类、合并策略、管理/托管标签、正确的标签位置
"""

import pandas as pd
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple
import os
from datetime import datetime
import math


class InputValidationError(Exception):
    """输入Excel格式错误"""


class ExcelToDrawioGenerator:
    """Excel到Draw.io格式生成 - 完整功能版本"""

    def __init__(self):
        # 完全按照fund_generator_v3.py的参数设置
        self.width = 3200  # 增加画布宽度
        self.height = 1400  # 增加高度以适应更大的层级间距
        self.fund_center_y = 900  # 基金位置下移

        # 不同类型节点的尺寸设置 - 与原版完全一致
        self.node_sizes = {
            'investor': {'width': 140, 'height': 120},      # 投资者节点
            'fund': {'width': 200, 'height': 90},          # 基金节点
            'manager': {'width': 200, 'height': 90},        # 管理人节点
            'custodian': {'width': 200, 'height': 90},      # 托管人节点
            'supervisor': {'width': 200, 'height': 90},     # 募集监督机构节点
            'target': {'width': 200, 'height': 90}          # 投资领域节点
        }

        # Draw.io样式定义
        self.styles = {
            'fund': 'rounded=1;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;fontStyle=1;fontSize=14;',
            'investor': 'rounded=1;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;fontSize=11;',
            'manager': 'rounded=1;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;fontSize=11;',
            'custodian': 'rounded=1;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;fontSize=11;',
            'supervisor': 'rounded=1;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;fontSize=11;',
            'target': 'rounded=1;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;fontSize=11;',
            'connection': 'edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;endArrow=classic;strokeColor=#000000;strokeWidth=2;',
            'management_line': 'edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;endArrow=classic;strokeColor=#333333;strokeWidth=2;'
        }

        self.cell_id_counter = 2  # Draw.io cell ID计数器（0,1为保留ID）

    def get_node_size(self, category: str) -> Dict[str, int]:
        """获取指定类型节点的尺寸"""
        return self.node_sizes.get(category, self.node_sizes['investor'])

    def get_next_cell_id(self) -> str:
        """获取下一个cell ID"""
        current_id = str(self.cell_id_counter)
        self.cell_id_counter += 1
        return current_id

    @staticmethod
    def _safe_get_value(df, row, col, default=''):
        """安全获取DataFrame值，处理NaN"""
        try:
            value = df.iloc[row, col]
            if pd.isna(value):
                return default
            return str(value).strip()
        except (IndexError, KeyError):
            return default

    @staticmethod
    def _validate_investor_columns(df: pd.DataFrame):
        """校验模板列，避免出现不友好的KeyError"""
        required = ['投资者层级', '名称', '投资至', '投资者分类']
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise InputValidationError(f"Excel 模板缺少列: {', '.join(missing)}")

    def load_excel(self, filepath: str) -> Tuple[Dict, pd.DataFrame]:
        """
        加载Excel文件 - 完全按照原版逻辑
        返回：(基本信息字典, 投资者DataFrame)
        """
        # 读取所有数据
        df_all = pd.read_excel(filepath, header=None)

        fund_info = {
            'chart_title': self._safe_get_value(df_all, 0, 0, '基金结构图'),
            'fund_name': self._safe_get_value(df_all, 1, 1, '基金产品'),
            'manager': self._safe_get_value(df_all, 2, 1, ''),
            'custodian': self._safe_get_value(df_all, 3, 1, ''),
            'supervisor': self._safe_get_value(df_all, 4, 1, ''),
            'investment_field': self._safe_get_value(df_all, 5, 1, ''),
        }

        # 第7行是表头，第8行开始是数据
        df_investors = pd.read_excel(filepath, header=6)  # header=6表示第7行（0-indexed）
        df_investors = df_investors.dropna(how='all').reset_index(drop=True)

        # 模板列校验（不影响生成逻辑）
        self._validate_investor_columns(df_investors)

        # 去掉说明行
        df_investors = df_investors[~df_investors['投资者层级'].astype(str).str.contains('说明', na=False)]
        if df_investors.empty:
            raise InputValidationError("投资者数据为空，请检查模板第 7 行表头和第 8 行开始的数据。")

        return fund_info, df_investors

    def build_graph(self, fund_info: Dict, df: pd.DataFrame) -> Tuple[Dict, str]:
        """构建节点图 - 完全按照原版逻辑，包含合并策略"""
        nodes = {}
        node_counter = 0

        # 1. 创建基金产品节点
        fund_id = f"N{node_counter}"
        node_counter += 1
        nodes[fund_id] = {
            'id': fund_id,
            'name': str(fund_info['fund_name']),
            'category': 'fund',
            'type': '基金产品',
            'level': 3,
            'parent_id': None,
            'amount': None,
            'ratio': None,
            'remark': '',
            'cell_id': self.get_next_cell_id()
        }

        # 2. 创建管理人节点
        if fund_info['manager']:
            mgr_id = f"N{node_counter}"
            node_counter += 1
            nodes[mgr_id] = {
                'id': mgr_id,
                'name': str(fund_info['manager']),
                'category': 'manager',
                'type': '管理人',
                'level': 3,
                'parent_id': None,
                'amount': None,
                'ratio': None,
                'remark': '',
                'cell_id': self.get_next_cell_id()
            }

        # 3. 创建托管人/募集监督机构节点 - 完全按照原版逻辑
        custodian = str(fund_info.get('custodian', '')) if fund_info.get('custodian') else ''
        supervisor = str(fund_info.get('supervisor', '')) if fund_info.get('supervisor') else ''

        # 处理可能的NaN值
        if custodian == 'nan':
            custodian = ''
        if supervisor == 'nan':
            supervisor = ''

        custodian = custodian.strip()
        supervisor = supervisor.strip()

        if supervisor:
            # 有募集监督机构，根据是否有托管决定显示文本
            cust_id = f"N{node_counter}"
            node_counter += 1
            if custodian:
                # 有托管：显示"募集监督&托管机构"
                nodes[cust_id] = {
                    'id': cust_id,
                    'name': supervisor,  # 显示募集监督机构的名称
                    'category': 'supervisor',
                    'type': '募集监督&托管机构',
                    'level': 3,
                    'parent_id': None,
                    'amount': None,
                    'ratio': None,
                    'remark': custodian,  # 将托管机构名称存在remark中
                    'cell_id': self.get_next_cell_id()
                }
            else:
                # 无托管：显示"募集监督机构"
                nodes[cust_id] = {
                    'id': cust_id,
                    'name': supervisor,
                    'category': 'supervisor',
                    'type': '募集监督机构',
                    'level': 3,
                    'parent_id': None,
                    'amount': None,
                    'ratio': None,
                    'remark': '',
                    'cell_id': self.get_next_cell_id()
                }
        elif custodian:
            # 只有托管人，没有募集监督机构（这种情况应该很少见）
            cust_id = f"N{node_counter}"
            node_counter += 1
            nodes[cust_id] = {
                'id': cust_id,
                'name': custodian,
                'category': 'custodian',
                'type': '托管人',
                'level': 3,
                'parent_id': None,
                'amount': None,
                'ratio': None,
                'remark': '',
                'cell_id': self.get_next_cell_id()
            }

        # 4. 创建投资领域节点
        if fund_info['investment_field']:
            target_id = f"N{node_counter}"
            node_counter += 1
            nodes[target_id] = {
                'id': target_id,
                'name': str(fund_info['investment_field']),
                'category': 'target',
                'type': '投资领域',
                'level': 4,
                'parent_id': fund_id,
                'amount': None,
                'ratio': None,
                'remark': '',
                'cell_id': self.get_next_cell_id()
            }

        # 5. 创建投资者节点（完全按照原版合并逻辑）
        groups = {}
        individual_nodes = {}
        row_to_id = {}
        name_to_ids = {}

        for idx, row in df.iterrows():
            name = str(row['名称']).strip()
            if not name or name == 'nan':
                continue

            # 解析层级
            level_str = str(row['投资者层级'])
            if '一级' in level_str:
                level = 1
            elif '二级' in level_str:
                level = 2
            elif '三级' in level_str:
                level = 3
            elif '四级' in level_str:
                level = 4
            else:
                level = 1

            invest_to = str(row['投资至']).strip() if pd.notna(row['投资至']) else ''
            partner_type = str(row['合伙人类型']).strip() if '合伙人类型' in row and pd.notna(row['合伙人类型']) else 'LP'
            investor_type = str(row['投资者分类']).strip() if pd.notna(row['投资者分类']) else ''

            # 判断是否需要合并：只有自然人和境内法人机构需要合并
            should_merge = ('自然人' in investor_type) or ('境内法人' in investor_type)

            if should_merge:
                # 需要合并的投资者
                parent_key = invest_to if invest_to else '__FUND__'
                key = (level, parent_key, partner_type, investor_type)

                if key not in groups:
                    groups[key] = {
                        'members': [],
                        'amount_sum': 0.0,
                        'ratio_sum': 0.0,
                        'has_amount': False,
                        'has_ratio': False
                    }

                amount = row['出资金额(万)'] if '出资金额(万)' in row and pd.notna(row['出资金额(万)']) else None
                # 比例由代码自动计算，不再从Excel读取
                ratio = None

                if amount is not None:
                    groups[key]['amount_sum'] += float(amount)
                    groups[key]['has_amount'] = True
                if ratio is not None:
                    groups[key]['ratio_sum'] += float(ratio)
                    groups[key]['has_ratio'] = True

                groups[key]['members'].append(name)
            else:
                # 不需要合并的投资者，保持独立节点
                investor_id = f"N{node_counter}"
                node_counter += 1
                row_to_id[idx] = investor_id

                amount = row['出资金额(万)'] if '出资金额(万)' in row and pd.notna(row['出资金额(万)']) else None
                # 比例由代码自动计算，不再从Excel读取
                ratio = None

                individual_nodes[investor_id] = {
                    'id': investor_id,
                    'name': name,
                    'category': 'investor',
                    'type': investor_type,
                    'partner_type': partner_type,
                    'level': level,
                    'parent_id': None,
                    'amount': amount,
                    'ratio': ratio,
                    'invest_to': invest_to,
                    'remark': '',
                    'cell_id': self.get_next_cell_id()
                }

                if name not in name_to_ids:
                    name_to_ids[name] = []
                name_to_ids[name].append((idx, investor_id))

        # 创建合并的组节点
        group_key_to_id = {}
        for key, info in groups.items():
            level, parent_key, partner_type, investor_type = key
            investor_id = f"N{node_counter}"
            node_counter += 1
            group_key_to_id[key] = investor_id

            amount_value = round(info['amount_sum'], 2) if info['has_amount'] else None
            ratio_value = round(info['ratio_sum'], 2) if info['has_ratio'] else None

            nodes[investor_id] = {
                'id': investor_id,
                'name': f"{partner_type} {investor_type}".strip(),
                'category': 'investor',
                'type': investor_type,
                'partner_type': partner_type,
                'level': level,
                'parent_id': None,
                'amount': amount_value,
                'ratio': ratio_value,
                'members': info['members'],
                'is_group': True,
                'remark': '',
                'cell_id': self.get_next_cell_id()
            }

        # 添加独立节点到nodes
        for investor_id, node_info in individual_nodes.items():
            nodes[investor_id] = node_info

        # 设置父子关系 - 完全按照原版逻辑
        # 1. 处理合并组的父子关系
        for key, investor_id in group_key_to_id.items():
            level, parent_key, partner_type, investor_type = key
            if level == 1:
                nodes[investor_id]['parent_id'] = fund_id
            else:
                if parent_key and parent_key != '__FUND__':
                    # 寻找父节点
                    parent_found = False
                    # 先在独立节点中找
                    if parent_key in name_to_ids:
                        for parent_idx, parent_id in name_to_ids[parent_key]:
                            if parent_id in nodes and nodes[parent_id]['level'] == level - 1:
                                nodes[investor_id]['parent_id'] = parent_id
                                parent_found = True
                                break
                    # 再在组节点中找
                    if not parent_found:
                        for parent_key_candidate, parent_group_id in group_key_to_id.items():
                            if parent_key_candidate[0] == level - 1:
                                # 检查是否有成员名称匹配
                                parent_members = nodes[parent_group_id].get('members', [])
                                if parent_key in parent_members:
                                    nodes[investor_id]['parent_id'] = parent_group_id
                                    parent_found = True
                                    break
                    if not parent_found:
                        nodes[investor_id]['parent_id'] = fund_id
                else:
                    nodes[investor_id]['parent_id'] = fund_id

        # 2. 处理独立节点的父子关系
        for investor_id, node_info in individual_nodes.items():
            invest_to = node_info['invest_to']
            level = node_info['level']

            if level == 1:
                nodes[investor_id]['parent_id'] = fund_id
            else:
                if invest_to and invest_to != '__FUND__':
                    parent_found = False
                    # 先在独立节点中找
                    if invest_to in name_to_ids:
                        for parent_idx, parent_id in name_to_ids[invest_to]:
                            if parent_id in nodes and nodes[parent_id]['level'] == level - 1:
                                nodes[investor_id]['parent_id'] = parent_id
                                parent_found = True
                                break
                    # 再在组节点中找
                    if not parent_found:
                        for group_key, group_id in group_key_to_id.items():
                            if group_key[0] == level - 1:
                                group_members = nodes[group_id].get('members', [])
                                if invest_to in group_members:
                                    nodes[investor_id]['parent_id'] = group_id
                                    parent_found = True
                                    break
                    if not parent_found:
                        nodes[investor_id]['parent_id'] = fund_id
                else:
                    nodes[investor_id]['parent_id'] = fund_id

        # 自动计算各级投资者的比例
        self._calculate_ratios(nodes, fund_id)

        return nodes, fund_id

    def _calculate_ratios(self, nodes: Dict, fund_id: str):
        """计算各级投资者相对于父节点的出资比例"""

        # 按父节点分组
        children_by_parent = {}

        for nid, node in nodes.items():
            if node['category'] == 'investor':
                parent_id = node.get('parent_id')
                if parent_id:
                    if parent_id not in children_by_parent:
                        children_by_parent[parent_id] = []
                    children_by_parent[parent_id].append(node)

        # 对每个父节点，计算其所有子节点的比例
        for parent_id, children in children_by_parent.items():
            # 计算该父节点下所有子节点的金额总和
            total_amount = sum(float(c['amount'] or 0) for c in children)

            # 计算每个子节点的比例（保留4位小数）
            for child in children:
                if total_amount > 0 and child['amount']:
                    child['ratio'] = round(float(child['amount']) / total_amount * 100, 4)
                else:
                    child['ratio'] = None

    def calculate_layout(self, nodes: Dict, fund_id: str) -> Dict:
        """计算节点布局 - 完全按照原版逻辑"""
        positions = {}

        # 按层级分组
        level_groups = {1: [], 2: [], 3: [], 4: []}
        for nid, node in nodes.items():
            if node['category'] == 'investor':
                level_groups[node['level']].append(nid)

        fund_size = self.get_node_size('fund')
        fund_center_x = (self.width - fund_size['width']) / 2 + fund_size['width'] / 2

        # 1. 基金产品
        positions[fund_id] = {
            'x': (self.width - fund_size['width']) / 2,
            'y': self.fund_center_y,
            'width': fund_size['width'],
            'height': fund_size['height']
        }

        # 2. 管理人、托管人和募集监督机构 - 完全按照原版逻辑
        for nid, node in nodes.items():
            if node['category'] == 'manager':
                mgr_size = self.get_node_size('manager')
                positions[nid] = {
                    'x': (self.width - fund_size['width']) / 2 - mgr_size['width'] - 80,
                    'y': self.fund_center_y,
                    'width': mgr_size['width'],
                    'height': mgr_size['height']
                }
            elif node['category'] == 'custodian':
                custodian_size = self.get_node_size('custodian')
                positions[nid] = {
                    'x': (self.width - fund_size['width']) / 2 + fund_size['width'] + 80,
                    'y': self.fund_center_y,
                    'width': custodian_size['width'],
                    'height': custodian_size['height']
                }
            elif node['category'] == 'supervisor':
                supervisor_size = self.get_node_size('supervisor')
                positions[nid] = {
                    'x': (self.width - fund_size['width']) / 2 + fund_size['width'] + 80,
                    'y': self.fund_center_y,
                    'width': supervisor_size['width'],
                    'height': supervisor_size['height']
                }
            elif node['category'] == 'target':
                target_size = self.get_node_size('target')
                positions[nid] = {
                    'x': fund_center_x - target_size['width'] / 2,
                    'y': 1050,
                    'width': target_size['width'],
                    'height': target_size['height']
                }

        # 3. 一级投资者（动态扩展画布宽度） - 完全按照原版逻辑
        level1_nodes = level_groups[1]
        level1_y = self.fund_center_y - 200  # 700
        if level1_nodes:
            investor_size = self.get_node_size('investor')
            num = len(level1_nodes)
            gap = 40

            # 计算所需总宽度
            required_width = num * investor_size['width'] + (num - 1) * gap + 200  # 200px边距

            # 动态扩展画布宽度
            if required_width > self.width:
                self.width = required_width
                # 重新计算基金和其他节点的位置
                fund_size = self.get_node_size('fund')
                fund_center_x = (self.width - fund_size['width']) / 2 + fund_size['width'] / 2

                # 更新基金位置
                positions[fund_id]['x'] = (self.width - fund_size['width']) / 2

                # 更新管理人、托管人等位置
                for nid, node in nodes.items():
                    if node['category'] == 'manager':
                        mgr_size = self.get_node_size('manager')
                        positions[nid]['x'] = (self.width - fund_size['width']) / 2 - mgr_size['width'] - 80
                    elif node['category'] in ['custodian', 'supervisor']:
                        custodian_size = self.get_node_size('custodian')
                        positions[nid]['x'] = (self.width - fund_size['width']) / 2 + fund_size['width'] + 80
                    elif node['category'] == 'target':
                        target_size = self.get_node_size('target')
                        positions[nid]['x'] = fund_center_x - target_size['width'] / 2

            # 单行布局
            total_w = num * investor_size['width'] + (num - 1) * gap
            start_x = (self.width - total_w) / 2

            for i, nid in enumerate(level1_nodes):
                positions[nid] = {
                    'x': start_x + i * (investor_size['width'] + gap),
                    'y': level1_y,
                    'width': investor_size['width'],
                    'height': investor_size['height']
                }

        # 4. 二级投资者（在对应父节点正上方，统一间距：200）
        level2_y = level1_y - 200  # 500
        self._layout_children_above_parent(nodes, positions, level_groups[2], level2_y)

        # 5. 三级投资者（统一间距：200，保持一致）
        level3_y = level2_y - 200  # 300
        self._layout_children_above_parent(nodes, positions, level_groups[3], level3_y)

        # 4级投资者（如果有的话，统一间距：200）
        if 4 in level_groups and level_groups[4]:
            level4_y = level3_y - 200  # 100
            self._layout_children_above_parent(nodes, positions, level_groups[4], level4_y)

        return positions

    def _layout_children_above_parent(self, nodes, positions, child_ids, y_pos):
        """将子节点布局在父节点正上方，严格保持父子对应关系 - 完全按照原版逻辑"""

        # 按父节点分组
        by_parent = {}
        for nid in child_ids:
            node = nodes[nid]
            parent_id = node.get('parent_id')
            if parent_id:
                if parent_id not in by_parent:
                    by_parent[parent_id] = []
                by_parent[parent_id].append(nid)

        investor_size = self.get_node_size('investor')
        gap = 20  # 子节点间距

        # 记录所有已分配的位置，用于冲突检测
        allocated_ranges = []  # [(start_x, end_x), ...]

        # 按父节点X坐标排序
        sorted_parents = sorted(by_parent.items(),
                               key=lambda x: positions[x[0]]['x'] if x[0] in positions else 0)

        # 为每个父节点的子节点分配位置
        for parent_id, children in sorted_parents:
            if parent_id not in positions:
                continue

            parent_pos = positions[parent_id]
            parent_center_x = parent_pos['x'] + parent_pos['width'] / 2
            num_children = len(children)

            if num_children == 1:
                # 单个子节点：优先紧凑排列
                nid = children[0]

                # 如果已有节点，尝试紧凑排列
                if allocated_ranges:
                    # 找到最右边已占用的位置，紧贴排列
                    max_occupied_x = max(end for start, end in allocated_ranges)
                    x_pos = max_occupied_x + gap

                else:
                    # 如果没有其他节点，使用父节点上方
                    x_pos = parent_center_x - investor_size['width'] / 2

                # 检查冲突并调整
                while any(not (x_pos + investor_size['width'] <= start or x_pos >= end)
                         for start, end in allocated_ranges):
                    x_pos += investor_size['width'] + gap

                positions[nid] = {
                    'x': x_pos,
                    'y': y_pos,
                    'width': investor_size['width'],
                    'height': investor_size['height']
                }
                allocated_ranges.append((x_pos, x_pos + investor_size['width']))
            else:
                # 多个子节点：优先紧凑排列，其次考虑父节点居中
                total_width = num_children * investor_size['width'] + (num_children - 1) * gap

                # 尝试紧凑排列
                if allocated_ranges:
                    max_occupied_x = max(end for start, end in allocated_ranges)
                    compact_start_x = max_occupied_x + gap
                    ideal_start_x = parent_center_x - total_width / 2

                    # 如果紧凑位置更左且不会太远离父节点，使用紧凑位置
                    if compact_start_x < ideal_start_x and abs((compact_start_x + total_width/2) - parent_center_x) <= 300:
                        start_x = compact_start_x
                    else:
                        start_x = ideal_start_x
                else:
                    start_x = parent_center_x - total_width / 2

                # 检查整组是否冲突并调整
                while any(not (start_x + total_width <= start or start_x >= end)
                         for start, end in allocated_ranges):
                    start_x += investor_size['width'] + gap

                for i, nid in enumerate(children):
                    x_pos = start_x + i * (investor_size['width'] + gap)
                    positions[nid] = {
                        'x': x_pos,
                        'y': y_pos,
                        'width': investor_size['width'],
                        'height': investor_size['height']
                    }

                allocated_ranges.append((start_x, start_x + total_width))

        # 检查是否需要扩展画布宽度
        if allocated_ranges:
            max_x = max(end for start, end in allocated_ranges)
            if max_x + 100 > self.width:
                self.width = max_x + 100

    def format_node_text(self, node: Dict) -> str:
        """格式化节点文本 - 完全按照原版逻辑"""
        lines = []

        if node['category'] == 'investor':
            # 投资者节点显示分类信息
            if node.get('is_group', False):
                # 合并组节点：显示类型标题 + 成员列表（完全按照原版逻辑）
                partner_type = node.get('partner_type', 'LP')
                investor_type = node.get('type', '')

                # 简化分类显示
                if '自然人' in investor_type:
                    suffix = '自然人'
                elif '境内法人' in investor_type:
                    suffix = '境内法人机构'
                else:
                    suffix = investor_type[:8]

                header = f"{partner_type} {suffix}".strip()
                lines.append(header)

                # 添加空行（分类和成员列表之间的空行）
                lines.append('')

                # 添加成员列表，每个成员单独换行，不进行内部换行
                members = node.get('members', [])
                for name in members:
                    lines.append(name)
            else:
                # 独立投资者节点：使用合伙人类型 + 分类
                partner_type = node.get('partner_type', 'LP')
                investor_type = node['type']

                # 简化分类显示
                if '自然人' in investor_type:
                    suffix = '自然人'
                elif '境内法人' in investor_type:
                    suffix = '境内法人机构'
                elif '境内非法人' in investor_type:
                    suffix = '境内非法人机构'
                elif '管理人跟投' in investor_type:
                    suffix = '管理人跟投'
                elif '私募基金' in investor_type:
                    suffix = '私募基金'
                elif '政府类引导基金' in investor_type:
                    suffix = '政府引导基金'
                elif '境外' in investor_type:
                    suffix = '境外机构'
                else:
                    suffix = investor_type[:8]  # 截取前8字

                prefix = f"{partner_type} {suffix}"
                lines.append(prefix)

                # 添加空行（投资者分类和名称之间的空行）
                lines.append('')

                # 添加投资者名称（不换行，直接显示完整名称）
                name = node['name']
                lines.append(name)

        elif node['category'] == 'manager':
            lines.append('基金管理人：')
            # 添加空行（分类和名称之间的空行）
            lines.append('')
            name = node['name']
            # 管理人名称不换行，直接显示完整名称
            lines.append(name)

        elif node['category'] == 'custodian':
            # 托管人节点
            lines.append('托管人：')
            # 添加空行（分类和名称之间的空行）
            lines.append('')
            name = node['name']
            if len(name) > 15:
                if '（' in name:
                    parts = name.split('（', 1)
                    lines.append(parts[0])
                    lines.append('（' + parts[1])
                else:
                    mid = len(name) // 2
                    lines.append(name[:mid])
                    lines.append(name[mid:])
            else:
                lines.append(name)

        elif node['category'] == 'supervisor':
            # 募集监督机构节点，根据type决定显示文本
            if node['type'] == '募集监督&托管机构':
                lines.append('募集监督&托管机构：')
            else:
                lines.append('募集监督机构：')
            # 添加空行（分类和名称之间的空行）
            lines.append('')
            name = node['name']
            if len(name) > 15:
                if '（' in name:
                    parts = name.split('（', 1)
                    lines.append(parts[0])
                    lines.append('（' + parts[1])
                else:
                    mid = len(name) // 2
                    lines.append(name[:mid])
                    lines.append(name[mid:])
            else:
                lines.append(name)

        else:
            # 其他节点类型
            name = str(node['name']) if node['name'] is not None else '未知节点'
            if len(name) > 15:
                if '（' in name:
                    parts = name.split('（', 1)
                    lines.append(parts[0])
                    lines.append('（' + parts[1])
                else:
                    mid = len(name) // 2
                    lines.append(name[:mid])
                    lines.append(name[mid:])
            else:
                lines.append(name)

        return '<br/>'.join(lines)

    def generate_drawio_xml(self, fund_info: Dict, nodes: Dict, positions: Dict, fund_id: str, output_file: str):
        """生成Draw.io XML文件"""

        # 创建根元素
        mxfile = ET.Element('mxfile', {
            'host': 'app.diagrams.net',
            'modified': datetime.now().isoformat(),
            'agent': 'Excel转Draw.io生成',
            'version': '21.7.5',
            'etag': f'fund_structure_{datetime.now().strftime("%Y%m%d_%H%M%S")}'
        })

        # 创建diagram元素
        diagram = ET.SubElement(mxfile, 'diagram', {
            'id': 'fund_structure_diagram',
            'name': fund_info.get('chart_title', '基金结构图')
        })

        # 创建mxGraphModel - 使用动态计算的画布大小
        graph_model = ET.SubElement(diagram, 'mxGraphModel', {
            'dx': '1422',
            'dy': '794',
            'grid': '1',
            'gridSize': '10',
            'guides': '1',
            'tooltips': '1',
            'connect': '1',
            'arrows': '1',
            'fold': '1',
            'page': '1',
            'pageScale': '1',
            'pageWidth': str(self.width),
            'pageHeight': str(self.height),
            'math': '0',
            'shadow': '0'
        })

        # 创建root元素
        root = ET.SubElement(graph_model, 'root')

        # 添加默认的两个cell（Draw.io要求）
        ET.SubElement(root, 'mxCell', {'id': '0'})
        ET.SubElement(root, 'mxCell', {'id': '1', 'parent': '0'})

        # 添加产品结构图标题
        title_text = fund_info.get('chart_title', '基金结构图')

        # 动态计算标题位置：在最高层节点上方60像素
        min_y = min(pos['y'] for pos in positions.values())
        title_y = max(10, min_y - 60)  # 确保不小于10

        title_cell = ET.SubElement(root, 'mxCell', {
            'id': str(self.cell_id_counter),
            'value': title_text,
            'style': 'text;html=1;strokeColor=none;fillColor=none;align=center;verticalAlign=middle;whiteSpace=nowrap;rounded=0;fontSize=24;fontStyle=1;',
            'vertex': '1',
            'parent': '1'
        })
        title_geometry = ET.SubElement(title_cell, 'mxGeometry', {
            'x': str(int(self.width / 2 - 300)),  # 居中显示，增加宽度
            'y': str(int(title_y)),  # 动态计算的Y位置
            'width': '600',  # 增加宽度以适应长标题
            'height': '40',  # 增加高度以适应24号字体
            'as': 'geometry'
        })
        self.cell_id_counter += 1

        # 添加节点
        for nid, node in nodes.items():
            if nid not in positions:
                continue

            pos = positions[nid]
            text = self.format_node_text(node)
            style = self.styles.get(node['category'], self.styles['investor'])

            # 确保所有值都是字符串，避免NaN问题
            cell_id = str(node['cell_id'])
            text = str(text) if text is not None else ''
            style = str(style) if style is not None else ''

            # 创建节点cell
            cell = ET.SubElement(root, 'mxCell', {
                'id': cell_id,
                'value': text,
                'style': style,
                'vertex': '1',
                'parent': '1'
            })

            # 添加几何信息
            ET.SubElement(cell, 'mxGeometry', {
                'x': str(int(pos['x'])),
                'y': str(int(pos['y'])),
                'width': str(int(pos['width'])),
                'height': str(int(pos['height'])),
                'as': 'geometry'
            })

        # 添加连接线和标签 - 完整功能版本
        self._add_connections_and_labels(root, nodes, positions, fund_id)

        # 写入文件
        tree = ET.ElementTree(mxfile)
        ET.indent(tree, space="  ", level=0)

        with open(output_file, 'w', encoding='utf-8', newline='\n') as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            tree.write(f, encoding='unicode', xml_declaration=False)

        print(f"[OK] Draw.io文件已生成: {output_file}")
        print(f"   画布大小: {self.width} x {self.height}")
        print(f"   节点数量: {len([n for n in nodes.values() if n['id'] in positions])}")

    def _add_connections_and_labels(self, root, nodes: Dict, positions: Dict, fund_id: str):
        """添加L型连接线和完整标签 - 包含管理/托管/投资标签"""

        fund_pos = positions[fund_id]
        fund_center_x = fund_pos['x'] + fund_pos['width'] / 2
        fund_top_y = fund_pos['y']

        # 1. 投资者之间的连接线和金额标签
        for nid, node in nodes.items():
            if node['category'] == 'investor' and nid in positions:
                parent_id = node.get('parent_id')
                if parent_id and parent_id in positions:
                    # 子节点位置
                    child_pos = positions[nid]
                    child_center_x = child_pos['x'] + child_pos['width'] / 2
                    child_bottom_y = child_pos['y'] + child_pos['height']

                    # 父节点位置
                    parent_pos = positions[parent_id]
                    parent_center_x = parent_pos['x'] + parent_pos['width'] / 2
                    parent_top_y = parent_pos['y']

                    # 创建L型连接线
                    conn_id = self.get_next_cell_id()

                    # 计算L型路径的中间点
                    if abs(child_center_x - parent_center_x) < 5:
                        # 垂直连接 - 绑定到节点
                        cell = ET.SubElement(root, 'mxCell', {
                            'id': conn_id,
                            'value': '',
                            'style': self.styles['connection'] + 'exitX=0.5;exitY=1;entryX=0.5;entryY=0;',
                            'edge': '1',
                            'parent': '1',
                            'source': str(node['cell_id']),  # 绑定源节点（子节点）
                            'target': str(nodes[parent_id]['cell_id'])  # 绑定目标节点（父节点）
                        })

                        geometry = ET.SubElement(cell, 'mxGeometry', {
                            'width': '50',
                            'height': '50',
                            'relative': '1',
                            'as': 'geometry'
                        })
                    else:
                        # L型连接：垂直向下，然后水平，再垂直向上
                        mid_y = (child_bottom_y + parent_top_y) / 2

                        # 创建带中间点的连接线 - 绑定到节点
                        cell = ET.SubElement(root, 'mxCell', {
                            'id': conn_id,
                            'value': '',
                            'style': self.styles['connection'] + 'exitX=0.5;exitY=1;entryX=0.5;entryY=0;',
                            'edge': '1',
                            'parent': '1',
                            'source': str(node['cell_id']),  # 绑定源节点（子节点）
                            'target': str(nodes[parent_id]['cell_id'])  # 绑定目标节点（父节点）
                        })

                        geometry = ET.SubElement(cell, 'mxGeometry', {
                            'width': '50',
                            'height': '50',
                            'relative': '1',
                            'as': 'geometry'
                        })

                        # 使用Draw.io标准正交样式，让它自动处理L型路径
                        # orthogonalEdgeStyle应该能自动创建完美的L型连接

                    # 添加金额和比例标签（在连接线起点右侧）
                    if (node.get('amount') is not None and pd.notna(node['amount'])) or \
                       (node.get('ratio') is not None and pd.notna(node['ratio'])):

                        label_lines = []

                        # 添加金额
                        if node.get('amount') is not None and pd.notna(node['amount']):
                            try:
                                amount_val = float(node['amount'])
                                if amount_val > 0:
                                    label_lines.append(f"{amount_val}万")
                            except (ValueError, TypeError):
                                pass

                        # 添加比例（格式：4位小数）
                        if node.get('ratio') is not None and pd.notna(node['ratio']):
                            try:
                                ratio_val = float(node['ratio'])
                                if ratio_val > 0:
                                    label_lines.append(f"{ratio_val:.4f}%")
                            except (ValueError, TypeError):
                                pass

                        if label_lines:
                            # 标签位置：完全按照原版SVG的精确位置
                            # 原版：连接线起点右侧25px，下方15px
                            # 连接线起点 = 子节点中心底部
                            label_x = child_center_x + 15  # 连接线起点右侧25px
                            label_y = child_bottom_y + 5  # 连接线起点下方15px

                            label_text = '<br/>'.join(label_lines)

                            # 创建标签
                            label_id = self.get_next_cell_id()
                            label_cell = ET.SubElement(root, 'mxCell', {
                                'id': label_id,
                                'value': label_text,
                                'style': 'text;html=1;strokeColor=none;fillColor=none;align=left;verticalAlign=middle;whiteSpace=wrap;rounded=0;fontSize=11;fontColor=#666666;',
                                'vertex': '1',
                                'parent': '1'
                            })

                            ET.SubElement(label_cell, 'mxGeometry', {
                                'x': str(int(label_x)),
                                'y': str(int(label_y)),
                                'width': '60',
                                'height': '40',
                                'as': 'geometry'
                            })

        # 2. 管理人连接线和"管理"标签
        for nid, node in nodes.items():
            if node['category'] == 'manager' and nid in positions:
                mgr_pos = positions[nid]

                x1 = mgr_pos['x'] + mgr_pos['width']
                y1 = self.fund_center_y + 50  # 使用固定的水平线Y坐标
                x2 = fund_pos['x']
                y2 = self.fund_center_y + 50

                # 创建管理连接线 - 绑定到节点
                conn_id = self.get_next_cell_id()
                cell = ET.SubElement(root, 'mxCell', {
                    'id': conn_id,
                    'value': '',
                    'style': self.styles['management_line'] + 'exitX=1;exitY=0.5;entryX=0;entryY=0.5;',
                    'edge': '1',
                    'parent': '1',
                    'source': str(node['cell_id']),  # 绑定管理人节点
                    'target': str(nodes[fund_id]['cell_id'])  # 绑定基金节点
                })

                geometry = ET.SubElement(cell, 'mxGeometry', {
                    'width': '50',
                    'height': '50',
                    'relative': '1',
                    'as': 'geometry'
                })

                # 添加"管理"标签
                label_id = self.get_next_cell_id()
                label_cell = ET.SubElement(root, 'mxCell', {
                    'id': label_id,
                    'value': '管理',
                    'style': 'text;html=1;strokeColor=none;fillColor=none;align=center;verticalAlign=middle;whiteSpace=wrap;rounded=0;fontSize=11;fontColor=#333333;fontStyle=1;',
                    'vertex': '1',
                    'parent': '1'
                })

                ET.SubElement(label_cell, 'mxGeometry', {
                    'x': str(int((x1 + x2) / 2 - 15)),
                    'y': str(int(y1 - 25)),
                    'width': '30',
                    'height': '20',
                    'as': 'geometry'
                })

        # 3. 托管人连接线和"托管"标签
        for nid, node in nodes.items():
            if node['category'] == 'custodian' and nid in positions:
                cust_pos = positions[nid]

                x1 = cust_pos['x']
                y1 = self.fund_center_y + 50
                x2 = fund_pos['x'] + fund_pos['width']
                y2 = self.fund_center_y + 50

                # 创建托管连接线 - 绑定到节点（从托管人指向基金）
                conn_id = self.get_next_cell_id()
                cell = ET.SubElement(root, 'mxCell', {
                    'id': conn_id,
                    'value': '',
                    'style': self.styles['management_line'] + 'exitX=0;exitY=0.5;entryX=1;entryY=0.5;',
                    'edge': '1',
                    'parent': '1',
                    'source': str(node['cell_id']),  # 绑定托管人节点（起点）
                    'target': str(nodes[fund_id]['cell_id'])  # 绑定基金节点（终点，箭头指向这里）
                })

                geometry = ET.SubElement(cell, 'mxGeometry', {
                    'width': '50',
                    'height': '50',
                    'relative': '1',
                    'as': 'geometry'
                })

                # 添加"托管"标签
                label_id = self.get_next_cell_id()
                label_cell = ET.SubElement(root, 'mxCell', {
                    'id': label_id,
                    'value': '托管',
                    'style': 'text;html=1;strokeColor=none;fillColor=none;align=center;verticalAlign=middle;whiteSpace=wrap;rounded=0;fontSize=11;fontColor=#333333;fontStyle=1;',
                    'vertex': '1',
                    'parent': '1'
                })

                ET.SubElement(label_cell, 'mxGeometry', {
                    'x': str(int((x1 + x2) / 2 - 15)),
                    'y': str(int(y1 - 25)),
                    'width': '30',
                    'height': '20',
                    'as': 'geometry'
                })

            elif node['category'] == 'supervisor' and nid in positions:
                # 只有募集监督机构的情况
                sup_pos = positions[nid]

                x1 = sup_pos['x']
                y1 = self.fund_center_y + 50
                x2 = fund_pos['x'] + fund_pos['width']
                y2 = self.fund_center_y + 50

                # 创建募集监督连接线（从募集监督机构指向基金）
                conn_id = self.get_next_cell_id()
                cell = ET.SubElement(root, 'mxCell', {
                    'id': conn_id,
                    'value': '',
                    'style': self.styles['management_line'] + 'exitX=0;exitY=0.5;entryX=1;entryY=0.5;',
                    'edge': '1',
                    'parent': '1',
                    'source': str(node['cell_id']),  # 绑定募集监督机构节点（起点）
                    'target': str(nodes[fund_id]['cell_id'])  # 绑定基金节点（终点，箭头指向这里）
                })

                geometry = ET.SubElement(cell, 'mxGeometry', {
                    'width': '50',
                    'height': '50',
                    'relative': '1',
                    'as': 'geometry'
                })

                ET.SubElement(geometry, 'mxPoint', {
                    'x': str(int(x1)),
                    'y': str(int(y1)),
                    'as': 'sourcePoint'
                })

                ET.SubElement(geometry, 'mxPoint', {
                    'x': str(int(x2)),
                    'y': str(int(y2)),
                    'as': 'targetPoint'
                })

                # 添加"募集监督"标签
                label_id = self.get_next_cell_id()
                label_cell = ET.SubElement(root, 'mxCell', {
                    'id': label_id,
                    'value': '募集监督',
                    'style': 'text;html=1;strokeColor=none;fillColor=none;align=center;verticalAlign=middle;whiteSpace=wrap;rounded=0;fontSize=11;fontColor=#333333;fontStyle=1;',
                    'vertex': '1',
                    'parent': '1'
                })

                ET.SubElement(label_cell, 'mxGeometry', {
                    'x': str(int((x1 + x2) / 2 - 25)),
                    'y': str(int(y1 - 25)),
                    'width': '50',
                    'height': '20',
                    'as': 'geometry'
                })

                # 如果同时是托管机构，在下方添加"托管"标签
                if node['type'] == '募集监督&托管机构':
                    label_id2 = self.get_next_cell_id()
                    label_cell2 = ET.SubElement(root, 'mxCell', {
                        'id': label_id2,
                        'value': '托管',
                        'style': 'text;html=1;strokeColor=none;fillColor=none;align=center;verticalAlign=middle;whiteSpace=wrap;rounded=0;fontSize=11;fontColor=#333333;fontStyle=1;',
                        'vertex': '1',
                        'parent': '1'
                    })

                    ET.SubElement(label_cell2, 'mxGeometry', {
                        'x': str(int((x1 + x2) / 2 - 15)),
                        'y': str(int(y1 - 5)),
                        'width': '30',
                        'height': '20',
                        'as': 'geometry'
                    })

        # 4. 基金到投资领域的连接线和"直接或间接股权投资"标签
        for nid, node in nodes.items():
            if node['category'] == 'target' and nid in positions:
                target_pos = positions[nid]
                target_center_x = target_pos['x'] + target_pos['width'] / 2
                target_top_y = target_pos['y']

                fund_bottom_y = fund_pos['y'] + fund_pos['height']

                # 创建连接线 - 绑定到节点
                conn_id = self.get_next_cell_id()
                cell = ET.SubElement(root, 'mxCell', {
                    'id': conn_id,
                    'value': '',
                    'style': self.styles['connection'] + 'exitX=0.5;exitY=1;entryX=0.5;entryY=0;',
                    'edge': '1',
                    'parent': '1',
                    'source': str(nodes[fund_id]['cell_id']),  # 绑定基金节点
                    'target': str(node['cell_id'])  # 绑定投资领域节点
                })

                # 添加几何信息
                geometry = ET.SubElement(cell, 'mxGeometry', {
                    'width': '50',
                    'height': '50',
                    'relative': '1',
                    'as': 'geometry'
                })

                # 添加"直接或间接股权投资"标签
                label_x = fund_center_x + 50
                label_y = (fund_bottom_y + target_top_y) / 2 - 20

                # 第一行：直接或间接
                label_id1 = self.get_next_cell_id()
                label_cell1 = ET.SubElement(root, 'mxCell', {
                    'id': label_id1,
                    'value': '直接或间接',
                    'style': 'text;html=1;strokeColor=none;fillColor=none;align=left;verticalAlign=middle;whiteSpace=wrap;rounded=0;fontSize=11;fontColor=#333333;fontStyle=1;',
                    'vertex': '1',
                    'parent': '1'
                })

                ET.SubElement(label_cell1, 'mxGeometry', {
                    'x': str(int(label_x)),
                    'y': str(int(label_y)),
                    'width': '70',
                    'height': '20',
                    'as': 'geometry'
                })

                # 第二行：股权投资
                label_id2 = self.get_next_cell_id()
                label_cell2 = ET.SubElement(root, 'mxCell', {
                    'id': label_id2,
                    'value': '股权投资',
                    'style': 'text;html=1;strokeColor=none;fillColor=none;align=left;verticalAlign=middle;whiteSpace=wrap;rounded=0;fontSize=11;fontColor=#333333;fontStyle=1;',
                    'vertex': '1',
                    'parent': '1'
                })

                ET.SubElement(label_cell2, 'mxGeometry', {
                    'x': str(int(label_x)),
                    'y': str(int(label_y + 20)),
                    'width': '70',
                    'height': '20',
                    'as': 'geometry'
                })

    def generate_from_excel(self, excel_file: str):
        """从Excel文件生成Draw.io数据（用于Web应用和CLI）"""
        fund_info, df = self.load_excel(excel_file)
        nodes, fund_id = self.build_graph(fund_info, df)
        positions = self.calculate_layout(nodes, fund_id)
        return fund_info, nodes, positions, fund_id

    def run(self, excel_file: str, output_file: str = None):
        """主流程（CLI用）"""
        if output_file is None:
            base_name = os.path.splitext(os.path.basename(excel_file))[0]
            output_file = f"{base_name}_drawio.drawio"

        print("\n" + "="*60)
        print("  Excel直接生成Draw.io格式转换器")
        print("="*60 + "\n")

        try:
            fund_info, nodes, positions, fund_id = self.generate_from_excel(excel_file)
            print(f"  基金: {fund_info['fund_name']}")
            print(f"  节点总数: {len(nodes)}\n")

            self.generate_drawio_xml(fund_info, nodes, positions, fund_id, output_file)

            print("\n" + "="*60)
            print(f"  [OK] 完成！请查看: {output_file}")
            print("  使用方法:")
            print("     1. 打开 https://app.diagrams.net/")
            print("     2. 文件 → 打开 → 选择生成的 .drawio 文件")
            print("     3. 开始编辑！")
            print("="*60 + "\n")
            return True

        except Exception as e:
            print(f"\n[ERROR] 错误: {e}")
            import traceback
            traceback.print_exc()
            return False


def cli_main():
    """命令行入口（整合自 converter/cli.py）"""
    import sys as _sys

    print("=== Excel到Draw.io转换器 ===")
    print("=" * 50)

    if len(_sys.argv) > 1:
        excel_file = _sys.argv[1]
    else:
        possible_files = [
            "产品结构图模板.xlsx",
            "static/产品结构图模板.xlsx",
        ]
        excel_file = None
        for fp in possible_files:
            if os.path.exists(fp):
                excel_file = fp
                break
        if not excel_file:
            print("[ERROR] 找不到Excel文件")
            print("\n使用方法:")
            print("  python generator.py [Excel文件路径]")
            _sys.exit(1)

    if not os.path.exists(excel_file):
        print(f"[ERROR] 文件不存在: {excel_file}")
        _sys.exit(1)

    generator = ExcelToDrawioGenerator()
    success = generator.run(excel_file)

    if success:
        print("\n提示:")
        print("- 生成的图表可以直接编辑")
        print("- 支持导出为PNG、PDF、SVG等格式")
        print("- 完全兼容Draw.io桌面版和在线版")
    else:
        _sys.exit(1)


if __name__ == "__main__":
    cli_main()
