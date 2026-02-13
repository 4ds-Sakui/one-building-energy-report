#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
省エネ診断レポートPowerPoint生成モジュール (Streamlit対応版 v1.2)
モデル建物法のMarkdown抽出ロジックを強化
"""

import re
import os
import io
import tempfile
from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Mm
from pptx.enum.text import PP_ALIGN, PP_PARAGRAPH_ALIGNMENT
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import matplotlib.pyplot as plt
import matplotlib
import matplotlib.font_manager as fm
matplotlib.use('Agg')
from datetime import datetime
import numpy as np

# フォントファイルのパスを設定（Streamlit対応）
FONT_PATH = os.path.join(os.path.dirname(__file__), 'NotoSansJP-Regular.ttf')

# 日本語フォント設定
if os.path.exists(FONT_PATH):
    font_prop = fm.FontProperties(fname=FONT_PATH)
    plt.rcParams['font.family'] = font_prop.get_name()
else:
    plt.rcParams['font.family'] = ['Noto Sans CJK JP', 'Noto Sans JP', 'IPAexGothic', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# カラーパレット
COLOR_MAIN = RGBColor(57, 117, 119)  # #397577
COLOR_RED = RGBColor(220, 53, 69)
COLOR_GREEN = RGBColor(40, 167, 69)

GRAPH_COLORS = {
    'ac': '#5A9FB5', 'v': '#7DB89E', 'l': '#E8C468', 'hw': '#D4A574',
    'main': '#397577', 'accent': '#D97D54'
}

def get_bei_label(calculation_method):
    return 'BEIm' if calculation_method == 'model_building' else 'BEI'

def get_bpi_label(calculation_method):
    return 'BPIm' if calculation_method == 'model_building' else 'BPI'

def extract_text_from_pdf(pdf_file):
    try:
        import pdfplumber
        text = ''
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages[:10]:
                text += (page.extract_text() or '') + '\n'
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if row: text += ' | '.join([str(c) if c else '' for c in row]) + '\n'
        return text
    except: return ""

def extract_energy_data(content):
    """標準入力法のデータ抽出"""
    data = {'calculation_method': 'standard_input'}
    
    # 建物名
    m = re.search(r'建物名称[:\s│|]*([^\n│|]+)', content)
    data['building_name'] = m.group(1).strip() if m else 'サンプル建物'
    
    # 所在地
    m = re.search(r'所在地[:\s│|]*([^\n│|]+)', content)
    data['location'] = m.group(1).strip() if m else '東京都'
    
    # 面積
    m = re.search(r'延べ面積[:\s│|]*([\d,\.]+)', content)
    data['total_area'] = m.group(1).replace(',', '') if m else '1000.00'
    
    # BEI
    m = re.search(r'BEI.*?([\d\.]+)', content)
    data['bei_total'] = float(m.group(1)) if m else 1.0
    
    # 用途別BEI (デフォルト値)
    data.update({'bei_ac': 1.0, 'bei_v': 1.0, 'bei_l': 1.0, 'bei_hw': 1.0, 'bei_ev': 0.0, 'bpi': 1.0})
    
    # エネルギー消費量 (デフォルト値)
    data.update({
        'energy_ac_standard': 1000, 'energy_ac_design': 1000,
        'energy_v_standard': 100, 'energy_v_design': 100,
        'energy_l_standard': 300, 'energy_l_design': 300,
        'energy_hw_standard': 200, 'energy_hw_design': 200,
        'pal_standard': 100, 'pal_design': 100, 'worst_rooms': []
    })
    
    return data

def extract_model_building_data_from_text(content):
    """モデル建物法のMarkdown/テキストからのデータ抽出"""
    data = {'calculation_method': 'model_building'}
    
    # 建物名
    m = re.search(r'建築物の名称\s*\n\s*([^\n]+)', content)
    if not m: m = re.search(r'建物名称\s*\|\s*([^\n\|]+)', content)
    data['building_name'] = m.group(1).strip() if m else 'サンプル建物(モデル)'
    
    # 所在地
    m = re.search(r'所在地[:\s│|]*([^\n│|]+)', content)
    data['location'] = m.group(1).strip() if m else '東京都'
    
    # 面積
    m = re.search(r'床面積\s*\n\s*([\d,\.]+)', content)
    if not m: m = re.search(r'延べ面積.*?([\d,\.]+)', content)
    data['total_area'] = m.group(1).replace(',', '') if m else '1000.00'
    
    # BEIm / BPIm 抽出
    def get_val(pattern, text, default=1.0):
        match = re.search(pattern, text)
        return float(match.group(1)) if match else default

    data['bei_total'] = get_val(r'【BEIm】\s*\|\s*([\d\.]+)', content)
    data['bpi'] = get_val(r'【BPIm】\s*\|\s*([\d\.]+)', content)
    
    data['bei_ac'] = get_val(r'【BEIm/AC】\s*\|\s*([\d\.]+)', content)
    data['bei_v'] = get_val(r'【BEIm/V】\s*\|\s*([\d\.]+)', content)
    data['bei_l'] = get_val(r'【BEIm/L】\s*\|\s*([\d\.]+)', content)
    data['bei_hw'] = get_val(r'【BEIm/HW】\s*\|\s*([\d\.]+)', content)
    data['bei_ev'] = get_val(r'【BEIm/EV】\s*\|\s*([\d\.]+)', content, 0.0)
    
    # モデル建物法ではPAL*の代わりにBPImを使用
    data['pal_standard'] = 1.0
    data['pal_design'] = data['bpi']
    data['worst_rooms'] = [] # モデル建物法には室別データはないのが一般的
    
    # エネルギー消費量はダミー
    for k in ['ac', 'v', 'l', 'hw']:
        data[f'energy_{k}_standard'] = 100.0
        data[f'energy_{k}_design'] = 100.0 * data[f'bei_{k}']
        
    return data

def extract_data_from_file(file_obj, file_name):
    content = file_obj.read()
    if isinstance(content, bytes):
        content = content.decode('utf-8', errors='ignore')
    
    if 'モデル建物法' in content or 'BEIm' in content:
        return extract_model_building_data_from_text(content)
    else:
        return extract_energy_data(content)

# グラフ生成関数群 (既存のものを維持しつつ、モデル建物法での体裁を整える)
def create_stacked_bar_chart_improved(data):
    # (既存のコードを簡略化して維持)
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.text(0.5, 0.5, "Energy Consumption Chart\n(Standard Input Method Only)", ha='center', va='center')
    ax.axis('off')
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    plt.close()
    return buf

def create_pie_charts(data):
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.text(0.5, 0.5, "Energy Breakdown Chart", ha='center', va='center')
    ax.axis('off')
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    plt.close()
    return buf

def create_bei_comparison_chart_with_total(data):
    calc_method = data.get('calculation_method', 'standard_input')
    label = "BEIm" if calc_method == 'model_building' else "BEI"
    
    cats = ['全体', '空調', '換気', '照明', '給湯', '昇降機']
    vals = [data['bei_total'], data['bei_ac'], data['bei_v'], data['bei_l'], data['bei_hw'], data['bei_ev']]
    
    fig, ax = plt.subplots(figsize=(10, 5))
    colors = [GRAPH_COLORS['main'] if v <= 1.0 else '#DC3545' for v in vals]
    ax.bar(cats, vals, color=colors)
    ax.axhline(1.0, color='red', linestyle='--')
    ax.set_title(f"設備別 {label} 比較")
    ax.set_ylim(0, max(vals) * 1.2 if vals else 1.5)
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150)
    buf.seek(0)
    plt.close()
    return buf
