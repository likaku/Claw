#!/usr/bin/env python3
"""
McKinsey 10 Tests — 战略诊断报告长图生成器 v2
白底黑字 + 执行摘要(From/To) + 雷达图 + 楷体中文/Arial英文
"""

import os
import math
import textwrap
from datetime import datetime

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
from PIL import Image, ImageDraw, ImageFont

# ── Fonts ──────────────────────────────────────────────────────
KAITI_PATH = '/System/Library/AssetsV2/com_apple_MobileAsset_Font8/88d6cc32a907955efa1d014207889413890573be.asset/AssetData/Kaiti.ttc'
ARIAL_PATH = '/Library/Fonts/Arial.ttf'

def kaiti(size, index=0):
    return ImageFont.truetype(KAITI_PATH, size, index=index)

def arial(size):
    return ImageFont.truetype(ARIAL_PATH, size)

# ── Color System ───────────────────────────────────────────────
BLACK   = (0, 0, 0)
DKGRAY  = (51, 51, 51)
MDGRAY  = (102, 102, 102)
LTGRAY  = (204, 204, 204)
BGGRAY  = (245, 245, 245)
WHITE   = (255, 255, 255)
NAVY    = (5, 28, 44)
RED     = (200, 50, 47)
ORANGE  = (210, 120, 30)
YELLOW  = (170, 150, 40)
GREEN   = (39, 130, 75)

SCORE_COLORS = {1: RED, 2: ORANGE, 3: YELLOW, 4: GREEN}
SCORE_DOTS   = {1: '●', 2: '▲', 3: '◆', 4: '★'}

# ── Layout Constants ───────────────────────────────────────────
CARD_W    = 750
MARGIN    = 48
CONTENT_W = CARD_W - 2 * MARGIN

# ── Text Utilities ─────────────────────────────────────────────
def _measure_text_height(draw, text, font, max_width, line_spacing=6):
    """Measure how tall a block of text will be without drawing."""
    if not text:
        return 0
    lines = []
    for paragraph in text.split('\n'):
        if not paragraph:
            lines.append('')
            continue
        # Estimate chars per line based on font size
        est_cpl = max(8, max_width // (font.size // 2 + 2))
        wrapped = textwrap.wrap(paragraph, width=est_cpl)
        lines.extend(wrapped if wrapped else [''])
    return len(lines) * (font.size + line_spacing)


def _draw_text(draw, x, y, text, font, fill=BLACK, max_width=None, line_spacing=6):
    """Draw text with auto-wrap. Returns new_y."""
    if not text:
        return y
    if max_width is None:
        max_width = CONTENT_W - x + MARGIN  # remaining width from x to right margin
    lines = []
    for paragraph in text.split('\n'):
        if not paragraph:
            lines.append('')
            continue
        est_cpl = max(8, max_width // (font.size // 2 + 2))
        wrapped = textwrap.wrap(paragraph, width=est_cpl)
        lines.extend(wrapped if wrapped else [''])
    line_h = font.size + line_spacing
    for i, line in enumerate(lines):
        draw.text((x, y + i * line_h), line, font=font, fill=fill)
    return y + len(lines) * line_h


def _section_header(draw, y, title):
    """Draw section header with navy underline."""
    font = kaiti(24)
    draw.text((MARGIN, y), title, font=font, fill=NAVY)
    y += 34
    draw.line([(MARGIN, y), (CARD_W - MARGIN, y)], fill=NAVY, width=2)
    y += 14
    return y


def _from_to_block(draw, y, from_text, to_text, rationale):
    """Draw a From→To block with background, return new y."""
    font_label = arial(12)
    font_body  = kaiti(16)
    font_rationale = kaiti(14)
    pad = 16
    
    # Measure height first
    h = pad  # top padding
    h += 16  # "FROM" label
    h += _measure_text_height(draw, from_text, font_body, CONTENT_W - pad * 2) + 6
    h += 24  # arrow
    h += 16  # "TO" label
    h += _measure_text_height(draw, to_text, font_body, CONTENT_W - pad * 2) + 6
    if rationale:
        h += 16  # "BECAUSE" label
        h += _measure_text_height(draw, rationale, font_rationale, CONTENT_W - pad * 2)
    h += pad  # bottom padding
    
    # Draw background first
    box_y = y
    draw.rounded_rectangle(
        [(MARGIN - 4, box_y), (CARD_W - MARGIN + 4, box_y + h)],
        radius=10, fill=BGGRAY
    )
    
    # Draw content
    cy = box_y + pad
    
    # FROM
    draw.text((MARGIN + pad, cy), 'FROM', font=font_label, fill=MDGRAY)
    cy += 18
    cy = _draw_text(draw, MARGIN + pad, cy, from_text, font_body, fill=DKGRAY,
                     max_width=CONTENT_W - pad * 2)
    cy += 8
    
    # Arrow
    draw.text((MARGIN + pad, cy), '→', font=arial(18), fill=NAVY)
    cy += 22
    
    # TO
    draw.text((MARGIN + pad, cy), 'TO', font=font_label, fill=MDGRAY)
    cy += 18
    cy = _draw_text(draw, MARGIN + pad, cy, to_text, font_body, fill=BLACK,
                     max_width=CONTENT_W - pad * 2)
    cy += 8
    
    # BECAUSE
    if rationale:
        draw.text((MARGIN + pad, cy), 'BECAUSE', font=font_label, fill=MDGRAY)
        cy += 18
        cy = _draw_text(draw, MARGIN + pad, cy, rationale, font_rationale, fill=MDGRAY,
                         max_width=CONTENT_W - pad * 2)
    
    return box_y + h + 12


def _score_bar(draw, y, label, score, max_score=4):
    """Draw a horizontal score bar."""
    bar_x = MARGIN + 120
    bar_w = 200
    bar_h = 10
    
    draw.text((MARGIN, y), label, font=kaiti(15), fill=DKGRAY)
    # Background bar
    draw.rounded_rectangle(
        [(bar_x, y + 4), (bar_x + bar_w, y + 4 + bar_h)],
        radius=5, fill=(230, 230, 230)
    )
    # Filled
    fill_w = int(bar_w * score / max_score)
    color = SCORE_COLORS.get(score, MDGRAY)
    if fill_w > 2:
        draw.rounded_rectangle(
            [(bar_x, y + 4), (bar_x + fill_w, y + 4 + bar_h)],
            radius=5, fill=color
        )
    # Score number
    draw.text((bar_x + bar_w + 10, y), f'{score}/{max_score}', font=arial(14), fill=DKGRAY)
    return y + 26


def generate_radar_chart(scores, labels, output_path):
    """Generate a radar/spider chart."""
    fm.fontManager.addfont(KAITI_PATH)
    plt.rcParams['font.family'] = ['Kaiti SC', 'Arial', 'sans-serif']
    
    N = len(labels)
    angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
    values = scores + [scores[0]]
    angles += [angles[0]]
    
    fig, ax = plt.subplots(figsize=(5.2, 5.2), subplot_kw=dict(polar=True))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')
    
    ax.fill(angles, values, color='#051C2C', alpha=0.07)
    ax.plot(angles, values, color='#051C2C', linewidth=2.5, linestyle='-')
    ax.scatter(angles[:-1], scores, color='#051C2C', s=50, zorder=5)
    
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, fontsize=12.5, fontfamily='Kaiti SC', color='#333333')
    
    ax.set_ylim(0, 4.5)
    ax.set_yticks([1, 2, 3, 4])
    ax.set_yticklabels(['1', '2', '3', '4'], fontsize=9, color='#AAAAAA', fontfamily='Arial')
    ax.yaxis.grid(True, color='#E0E0E0', linewidth=0.6)
    ax.xaxis.grid(True, color='#E0E0E0', linewidth=0.6)
    ax.spines['polar'].set_visible(False)
    
    # Score dots on vertices
    for angle, score in zip(angles[:-1], scores):
        color_hex = '#%02x%02x%02x' % SCORE_COLORS.get(score, MDGRAY)
        ax.scatter([angle], [score], color=color_hex, s=60, zorder=6, edgecolors='white', linewidths=1.5)
    
    plt.tight_layout(pad=1.2)
    plt.savefig(output_path, dpi=160, bbox_inches='tight', facecolor='white', pad_inches=0.2)
    plt.close()


def generate_report_card(data, output_path):
    """
    Generate a phone-width long image for the McKinsey 10 Tests report.
    
    data structure:
    {
        'title', 'subject', 'scenario', 'advisor', 'date',
        'total_score', 'max_score', 'grade', 'summary',
        'from_to_items': [{'from', 'to', 'rationale'}, ...],
        'dimensions': [{'name', 'name_en', 'score', 'finding', 'advice'}, ...],
        'top3_strengths': [{'name', 'score', 'reason'}],
        'top3_improvements': [{'name', 'score', 'reason', 'action'}],
        'actions': {'48h': [...], '1w': [...], '1m': [...]},
    }
    """
    # ── Radar chart ──
    radar_path = output_path.replace('.png', '_radar_tmp.png')
    dim_labels = [d['name'] for d in data['dimensions']]
    dim_scores = [d['score'] for d in data['dimensions']]
    generate_radar_chart(dim_scores, dim_labels, radar_path)
    
    # ── Canvas ──
    canvas_h = 8000
    img = Image.new('RGB', (CARD_W, canvas_h), WHITE)
    draw = ImageDraw.Draw(img)
    y = 0
    
    # ━━━ HEADER ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    y += 48
    draw.text((MARGIN, y), data.get('title', 'McKinsey 10 Tests'), font=kaiti(30), fill=NAVY)
    y += 42
    draw.text((MARGIN, y), '战略诊断报告', font=kaiti(17), fill=MDGRAY)
    y += 32
    
    # Meta
    font_meta = kaiti(13)
    for line in [
        f'诊断对象：{data["subject"]}',
        f'诊断场景：{data["scenario"]}',
        f'诊断顾问：{data["advisor"]}',
        f'日期：{data["date"]}',
    ]:
        draw.text((MARGIN, y), line, font=font_meta, fill=MDGRAY)
        y += 22
    y += 12
    draw.line([(MARGIN, y), (CARD_W - MARGIN, y)], fill=LTGRAY, width=1)
    y += 20
    
    # ━━━ EXECUTIVE SUMMARY ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    y = _section_header(draw, y, '执行摘要')
    
    # Disclaimer
    draw.text((MARGIN, y), '⚠ 本文为初始想法的梳理，而非最终结论', font=kaiti(14), fill=ORANGE)
    y += 28
    
    # Summary
    y = _draw_text(draw, MARGIN, y, data['summary'], kaiti(16), fill=DKGRAY)
    y += 18
    
    # From→To blocks
    for item in data.get('from_to_items', []):
        y = _from_to_block(draw, y, item['from'], item['to'], item.get('rationale', ''))
    
    y += 8
    
    # ━━━ RADAR CHART ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    y = _section_header(draw, y, '十维得分一览')
    
    radar_img = Image.open(radar_path)
    radar_w = CONTENT_W
    ratio = radar_w / radar_img.width
    radar_h = int(radar_img.height * ratio)
    radar_img = radar_img.resize((radar_w, radar_h), Image.LANCZOS)
    img.paste(radar_img, (MARGIN, y))
    y += radar_h + 12
    
    # Score bars
    for dim in data['dimensions']:
        y = _score_bar(draw, y, dim['name'], dim['score'])
    y += 8
    
    # ━━━ KEY FINDINGS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    y = _section_header(draw, y, '关键发现与建议')
    
    for dim in data['dimensions']:
        color = SCORE_COLORS.get(dim['score'], MDGRAY)
        dot = SCORE_DOTS.get(dim['score'], '●')
        
        # Dimension header: icon + name + score
        header = f'{dot}  {dim["name"]}（{dim["name_en"]}）  {dim["score"]}/4'
        draw.text((MARGIN, y), header, font=kaiti(16), fill=color)
        y += 24
        
        # Finding
        y = _draw_text(draw, MARGIN + 10, y, dim['finding'], kaiti(14), fill=DKGRAY,
                        max_width=CONTENT_W - 10)
        y += 4
        
        # Advice (with arrow prefix, navy color for emphasis)
        y = _draw_text(draw, MARGIN + 10, y, f'→ {dim["advice"]}', kaiti(14), fill=NAVY,
                        max_width=CONTENT_W - 10)
        y += 18
    
    # ━━━ TOP 3 STRENGTHS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    y = _section_header(draw, y, '最强维度 TOP 3')
    for i, s in enumerate(data['top3_strengths']):
        text = f'{i+1}. {s["name"]}（{s["score"]}/4）{s["reason"]}'
        y = _draw_text(draw, MARGIN, y, text, kaiti(15), fill=DKGRAY)
        y += 8
    
    # ━━━ TOP 3 IMPROVEMENTS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    y += 6
    y = _section_header(draw, y, '优先改进 TOP 3')
    for i, imp in enumerate(data['top3_improvements']):
        color = RED if imp['score'] == 1 else ORANGE
        text = f'{i+1}. {imp["name"]}（{imp["score"]}/4）{imp["reason"]}'
        y = _draw_text(draw, MARGIN, y, text, kaiti(15), fill=color)
        if imp.get('action'):
            y = _draw_text(draw, MARGIN + 12, y, f'→ {imp["action"]}', kaiti(14), fill=NAVY)
        y += 8
    
    # ━━━ ACTION ITEMS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    y = _section_header(draw, y, '行动清单')
    
    actions = data.get('actions', {})
    for label_cn, key in [('48小时内', '48h'), ('1周内', '1w'), ('1个月内', '1m')]:
        items = actions.get(key, [])
        if not items:
            continue
        draw.text((MARGIN, y), label_cn, font=arial(12), fill=MDGRAY)
        y += 18
        for item in items:
            y = _draw_text(draw, MARGIN + 10, y, f'• {item}', kaiti(14), fill=DKGRAY,
                            max_width=CONTENT_W - 10)
            y += 3
        y += 10
    
    # ━━━ FOOTER ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    y += 12
    draw.line([(MARGIN, y), (CARD_W - MARGIN, y)], fill=LTGRAY, width=1)
    y += 10
    draw.text((MARGIN, y), 'McKinsey Strategic 10 Tests', font=arial(10), fill=LTGRAY)
    y += 16
    draw.text((MARGIN, y), '基于 "Have You Tested Your Strategy Lately?" (McKinsey, 2011)',
              font=kaiti(11), fill=LTGRAY)
    y += 16
    draw.text((MARGIN, y), f'Generated: {data["date"]}', font=arial(10), fill=LTGRAY)
    y += 24
    
    # ── Crop & save ──
    img = img.crop((0, 0, CARD_W, y))
    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    img.save(output_path, 'PNG', quality=95)
    
    # Cleanup temp radar
    if os.path.exists(radar_path):
        os.remove(radar_path)
    
    print(f'✅ Report card: {output_path} ({img.width}×{img.height}px)')
    return output_path


# ── Demo ───────────────────────────────────────────────────────
if __name__ == '__main__':
    demo_data = {
        'title': 'McKinsey 10 Tests',
        'subject': '日世（中国）从B端冰淇淋原料供应商转型C端品牌 / 出海东南亚',
        'scenario': '🏢 企业战略',
        'advisor': 'Prof. Sterling · 麦肯锡资深董事合伙人',
        'date': '2026-04-08',
        'total_score': 25,
        'max_score': 40,
        'grade': 'C+',
        'summary': '你们拥有稀缺的品牌资产（日本软冰淇淋开创者、弟弟妹妹IP）和扎实的技术底盘，董事长的执行决心是最大亮点。但核心问题在于聚焦不足和独到洞见缺失——同时打太多仗，又在关键战场缺乏信息差。',
        'from_to_items': [
            {
                'from': '在中国B端"求量"——试图挽回所有流失份额，陷入价格战',
                'to': '在中国B端"求质"——聚焦3-5个高价值客户深度绑定，利润率优先',
                'rationale': '肯德基已在培养替代供应商，价格战只会加速利润流失，不如主动收缩战线保利润',
            },
            {
                'from': '中国五线并行（大客户/小B/C端/出海/收购），资源分散',
                'to': '东南亚B端大客户单一焦点，第一年只做一件事',
                'rationale': '资源有限时，同时做5件事=每件事都做到20%；先做1件事做到80%，再复制模式',
            },
            {
                'from': '"同一起跑线"——对东南亚市场无信息差',
                'to': '董事长3个月深度浸泡，建立至少2-3个独家认知（口味偏好/渠道结构/冷链痛点）',
                'rationale': '明治、格力高已在东南亚深耕多年，没有信息差就是在用弱点打别人的强点',
            },
        ],
        'dimensions': [
            {'name': '市场竞胜力', 'name_en': 'Beat the Market', 'score': 2,
             'finding': '在中国B端正失去竞争力——大客户培养替代供应商，价格战压缩利润。C端盒马已失败。东南亚尚未验证。',
             'advice': '立即从"求量"转"求质"，聚焦3-5个高价值客户，年利润率目标>15%'},
            {'name': '优势来源', 'name_en': 'Advantage Source', 'score': 3,
             'finding': '底层资产强——日本开创者品牌、弟弟妹妹IP、设备+原料闭环。但优势未被翻译成消费者可感知的差异化。',
             'advice': '在东南亚重新包装"日本品牌"故事，用创始人叙事+IP情感化做差异化核武器'},
            {'name': '精准聚焦', 'name_en': 'Granularity', 'score': 2,
             'finding': '同时在中国打五条线，东南亚又走三条路。资源严重分散，没有一条被充分验证。',
             'advice': '东南亚第一年只做B端大客户供应一件事；中国砍掉或暂停至少两条线'},
            {'name': '趋势前瞻', 'name_en': 'Ahead of Trends', 'score': 3,
             'finding': '对大趋势判断正确——中国B端红海化、东南亚冷饮爆发、收购扩品类增粘性。但看到机会≠做出取舍。',
             'advice': '把"看到好机会"的判断力转化为"不做什么"的勇气——每个机会必须回答"为什么只有我能抓住"'},
            {'name': '独到洞见', 'name_en': 'Privileged Insights', 'score': 1,
             'finding': '最大红色预警。在最关键的战场（东南亚）与竞争对手站在同一起跑线——不了解当地口味、供应链、客户决策逻辑。',
             'advice': '董事长在投第一分钱之前，花3个月深度浸泡——蹲便利店、跑经销商、理解冷链，把"同一起跑线"变成信息差'},
            {'name': '不确定性管理', 'name_en': 'Embrace Uncertainty', 'score': 3,
             'finding': '有明确止损线（2027年6月/5000万），但缺少中间里程碑，可能导致"温水煮青蛙"。',
             'advice': '设3个季度检查点：Q4签首客→Q1首批交付→Q2月销XX万，任何一个miss立即复盘'},
            {'name': '承诺-灵活', 'name_en': 'Commit vs Flex', 'score': 3,
             'finding': '1000万+董事长亲自下场的承诺度够高，先B后C的节奏合理。',
             'advice': '把B端定义为"No-Regrets Move"（无悔之举），C端/加盟定义为"Option"，等B端站稳再激活'},
            {'name': '去偏见', 'name_en': 'Free from Bias', 'score': 2,
             'finding': '存在"选择逃避偏见"——每次面对取舍时"都要"而非"只选一个"。中国五线并行，东南亚三路齐发。',
             'advice': '董事会引入"红队机制"——指定一人专门质疑每个新方向，强制做"如果只能做一件事"实验'},
            {'name': '执行决心', 'name_en': 'Conviction', 'score': 4,
             'finding': '最强维度。董事长亲自带队出海，愿意亏损1000万，有明确止损线——真正的战略承诺。',
             'advice': '保持决心但警惕沉没成本——如果季度检查点miss，要有壮士断腕的勇气'},
            {'name': '行动计划', 'name_en': 'Action Plan', 'score': 3,
             'finding': '有计划和时间表（2027年6月/5000万），但尚在早期，未签下第一个东南亚客户来验证模式。',
             'advice': '第一个90天唯一目标：签下东南亚第一个标杆客户，用这一个case跑通全流程'},
        ],
        'top3_strengths': [
            {'name': '执行决心', 'score': 4, 'reason': '董事长亲自下场+1000万承诺，最稀缺的战略资源'},
            {'name': '优势来源', 'score': 3, 'reason': '日本开创者品牌+弟弟妹妹IP+技术闭环，底牌够硬'},
            {'name': '趋势前瞻', 'score': 3, 'reason': '对市场大趋势判断准确，方向没选错'},
        ],
        'top3_improvements': [
            {'name': '独到洞见', 'score': 1, 'reason': '最关键战场无信息差——最紧急',
             'action': '董事长3个月深度浸泡东南亚，建2-3个独家认知'},
            {'name': '精准聚焦', 'score': 2, 'reason': '战线太长资源分散',
             'action': '东南亚第一年只做B端；中国砍掉≥2条线'},
            {'name': '去偏见', 'score': 2, 'reason': '"都要做"的逃避取舍偏见',
             'action': '董事会引入红队机制，强制"只做一件事"实验'},
        ],
        'actions': {
            '48h': [
                '列出中国五条线的资源投入占比，标注哪两条可立即暂停',
                '开始收集东南亚目标市场基础信息（冷饮规模、竞争格局、冷链基建）',
            ],
            '1w': [
                '确定东南亚出海第一站国家（建议泰国或越南二选一）',
                '联系当地行业协会/经销商/日系食品企业本地负责人',
                '董事会明确"东南亚B端优先"战略决议，暂缓加盟和小B',
            ],
            '1m': [
                '董事长完成第一次东南亚深度调研（至少两周驻扎）',
                '识别3-5个东南亚标杆客户并启动接触',
                '设置出海项目季度检查点和里程碑表',
            ],
        },
    }
    
    output = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output', '日世战略诊断长图.png')
    generate_report_card(demo_data, output)
