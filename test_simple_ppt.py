#!/usr/bin/env python3
"""
测试简单版本 PPT
"""
import sys
sys.path.insert(0, '/Users/kaku/.workbuddy/skills/mck-ppt-design')
from mck_ppt.engine import MckEngine
from mck_ppt.constants import *
from mck_ppt.core import full_cleanup

eng = MckEngine()

# 只生成前 5 页测试
eng.cover(
    title='2026 全球大趋势：在碎片化世界中寻找新范式',
    subtitle='特朗普2.0时代的全球重构、AI的落地考量与经济气候的临界点',
    author='',
    date='2026'
)

eng.pyramid(
    title='核心洞见：2026年标志着旧有全球秩序的彻底终结',
    levels=[
        ('地缘政治', '美国走向孤立与交易型外交，中国加速布局全球南方，欧洲被迫军事觉醒', 1.0),
        ('全球经济', '关税引发供应链"敏捷重构"，高赤字导致债券市场面临灰犀牛风险', 2.5),
        ('技术演进', 'AI从"狂热投资"转向"企业ROI落地"，底层白领岗位面临系统性消失', 4.0),
        ('社会与环境', '1.5°C气候红线被正式突破，廉价GLP-1减肥药重塑全球健康版图', 5.5),
    ],
    source='Source: 全球趋势研究报告, 2026'
)

eng.big_number(
    title='美国的孤立主义与"交易型外交"重塑全球规则',
    number='250',
    unit='周年',
    description='美国建国250周年之际，内部分歧达到顶峰',
    detail_items=[
        '外交不再基于战略同盟，而是通过关税大棒和利益交换解决问题',
        '美国可能动摇美联储的独立性',
        '各国无法再依赖美国的"安全伞"',
    ],
    source='Source: 地缘政治分析, 2026'
)

output_path = 'output/test_simple.pptx'
eng.save(output_path)
print(f"✅ 测试 PPT 已生成: {output_path}")

full_cleanup(output_path)
print(f"✅ 文件已清理: {output_path}")
