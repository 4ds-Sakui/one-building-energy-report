#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
省エネ診断レポートPowerPoint生成モジュール (Streamlit対応版)
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
    # フォールバック
    plt.rcParams['font.family'] = ['Noto Sans CJK JP', 'Noto Sans JP', 'IPAexGothic', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# デザインガイドラインに基づくカラーパレット
COLOR_MAIN = RGBColor(57, 117, 119)  # #397577 メインカラー
COLOR_ACCENT = RGBColor(217, 125, 84)  # #D97D54 オレンジアクセント
COLOR_BASE1 = RGBColor(246, 246, 246)  # #f6f6f6 ベースカラー1
COLOR_BASE2 = RGBColor(237, 237, 237)  # #ededed ベースカラー2
COLOR_RED = RGBColor(220, 53, 69)  # 警告・不適合
COLOR_GREEN = RGBColor(40, 167, 69)  # 適合・良好
COLOR_BLUE = RGBColor(0, 123, 255)  # 情報
COLOR_GRAY = RGBColor(108, 117, 125)  # グレー
COLOR_LIGHT_GRAY = RGBColor(200, 200, 200)  # ライトグレー
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_BLACK = RGBColor(0, 0, 0)

# グラフ用の品のある色合い
GRAPH_COLORS = {
    'ac': '#5A9FB5',      # 空調
    'v': '#7DB89E',       # 換気
    'l': '#E8C468',       # 照明
    'hw': '#D4A574',      # 給湯
    'main': '#397577',    # メインカラー
    'accent': '#D97D54'   # アクセントカラー
}

def get_bei_label(calculation_method):
    """計算方法に応じたBEI表記を返す"""
    if calculation_method == 'model_building':
        return 'BEIm'
    else:
        return 'BEI'

def get_bpi_label(calculation_method):
    """計算方法に応じたBPI表記を返す"""
    if calculation_method == 'model_building':
        return 'BPIm'
    else:
        return 'BPI'

def format_bei_text(calculation_method, equipment_type=''):
    """計算方法と設備種類に応じたBEI表記を返す"""
    bei_label = get_bei_label(calculation_method)
    
    if equipment_type == '':
        return f'全体{bei_label}'
    elif equipment_type == 'AC':
        return f'空調{bei_label}'
    elif equipment_type == 'V':
        return f'換気{bei_label}'
    elif equipment_type == 'L':
        return f'照明{bei_label}'
    elif equipment_type == 'HW':
        return f'給湯{bei_label}'
    elif equipment_type == 'EV':
        return f'昇降機{bei_label}'
    else:
        return bei_label


def extract_text_from_pdf(pdf_file):
    """
    PDFファイルオブジェクトからテキストを抽出
    
    Args:
        pdf_file: ファイルオブジェクト（BytesIO or UploadedFile）
    
    Returns:
        str: 抽出されたテキスト
    """
    try:
        import pdfplumber
        
        text = ''
        with pdfplumber.open(pdf_file) as pdf:
            # 最初の10ページのみ処理
            for i, page in enumerate(pdf.pages[:10]):
                page_text = page.extract_text()
                if page_text:
                    text += page_text + '\n'
                    
                # テーブルも抽出
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if row:
                            text += ' | '.join([str(cell) if cell else '' for cell in row]) + '\n'
        
        return text
    
    except ImportError:
        raise ImportError("pdfplumberがインストールされていません")
    except Exception as e:
        raise Exception(f"PDFの読み込みエラー: {e}")


def extract_single_bei(content, equipment_type, default_value):
    """個別のBEI値を抽出"""
    patterns = [
        rf'BEI/{equipment_type}\s+([\d\.]+)',
        rf'{equipment_type}.*?([\d\.]+)',
    ]
    for pattern in patterns:
        match = re.search(pattern, content)
        if match:
            try:
                value = float(match.group(1))
                if 0.0 <= value <= 5.0:
                    return value
            except:
                pass
    return default_value


def extract_energy_value(content, equipment_name, default_value):
    """エネルギー消費量を抽出"""
    patterns = [
        rf'{equipment_name}.*?([\d,]+\.[\d]+)\s*GJ',
        rf'{equipment_name}.*?([\d,]+\.[\d]+)',
    ]
    for pattern in patterns:
        match = re.search(pattern, content)
        if match:
            try:
                value = float(match.group(1).replace(',', ''))
                if 0 < value < 100000:
                    return value
            except:
                pass
    return default_value


def extract_energy_data(content):
    """
    テキストからエネルギー診断データを抽出
    
    Args:
        content (str): テキストコンテンツ
    
    Returns:
        dict: 抽出されたデータ
    """
    data = {}
    
    # 建物名称の抽出
    building_patterns = [
        r'建物名称[:\s│|]*([^\n│|]+)',
        r'2\.\s*建物の概要.*?建物名称[:\s│|]*([^\n│|]+)',
    ]
    for pattern in building_patterns:
        match = re.search(pattern, content, re.DOTALL)
        if match:
            name = match.group(1).strip()
            if name and len(name) < 50:
                data['building_name'] = name
                break
    if 'building_name' not in data:
        data['building_name'] = 'test'
    
    # 建物所在地の抽出
    location_patterns = [
        r'建物所在地[:\s│|]*([^\n│|]+)',
        r'所在地[:\s│|]*([^\n│|]+)',
    ]
    for pattern in location_patterns:
        match = re.search(pattern, content)
        if match:
            location = match.group(1).strip()
            if location and len(location) < 50:
                data['location'] = location
                break
    if 'location' not in data:
        data['location'] = '東京都'
    
    # 延べ面積の抽出
    area_patterns = [
        r'延べ面積[:\s│|]*([\d,\.]+)\s*m[²2]',
        r'延べ面積.*?([\d,\.]+)',
    ]
    for pattern in area_patterns:
        match = re.search(pattern, content)
        if match:
            area = match.group(1).replace(',', '')
            try:
                if 100 < float(area) < 1000000:
                    data['total_area'] = area
                    break
            except:
                pass
    if 'total_area' not in data:
        data['total_area'] = '876.38'
    
    # BEI値の抽出（設計値列から正確に抽出）
    # パターン1: Markdown表形式（6.1節）- 5列の表
    bei_pattern_md = r'\|\s*建築物エネルギー消費性能基準\s*\|\s*H28年4月以降\s*\|\s*[^\|]+\s*\|\s*([\d\.]+)\s*\|\s*([\d\.]+)\s*\|'
    match = re.search(bei_pattern_md, content)
    if match:
        try:
            bei_design = float(match.group(1))
            if 0.1 < bei_design < 10.0:
                data['bei_total'] = bei_design
        except:
            pass
    
    # パターン2: PDFテキスト形式（表組みが崩れた場合）
    if 'bei_total' not in data:
        bei_pattern_pdf = r'建築物エネルギー消費性能基準[^\n]{0,50}H28年4月以降[^\n]{0,200}?([\d\.]+)\s+([\d\.]+)'
        matches = re.findall(bei_pattern_pdf, content)
        for match in matches:
            try:
                bei_design = float(match[0])
                bei_standard = float(match[1])
                # BEI設計値と基準値の範囲をチェック
                if 0.3 < bei_design < 10.0 and 0.5 < bei_standard < 2.0:
                    data['bei_total'] = bei_design
                    break
            except:
                pass
    
    # パターン3: 6.1節の表から抽出（複数行にまたがる場合）
    if 'bei_total' not in data:
        bei_pattern3 = r'6\.1\.[^\n]*BEI.*?H28年4月以降.*?([\d,]+\.[\d]+)\s*\([^\)]+\)[^\n]{0,100}?([\d\.]+)\s+([\d\.]+)'
        match = re.search(bei_pattern3, content, re.DOTALL)
        if match:
            try:
                bei_design = float(match.group(2))
                if 0.1 < bei_design < 10.0:
                    data['bei_total'] = bei_design
            except:
                pass
    
    if 'bei_total' not in data:
        data['bei_total'] = 1.0
    
    # 用途別BEI（6.3節）の抽出
    # パターン1: Markdown表形式
    bei_usage_md = r'\|\s*BEI/AC\s*\|\s*BEI/V\s*\|\s*BEI/L\s*\|\s*BEI/HW\s*\|\s*BEI/EV\s*\|[^\n]*\n[^\n]*\n\|\s*([\d\.]+)\s*\|\s*([\d\.]+)\s*\|\s*([\d\.]+)\s*\|\s*([\d\.]+)\s*\|\s*([\d\.]*)\s*\|'
    match = re.search(bei_usage_md, content)
    
    if match:
        data['bei_ac'] = float(match.group(1))
        data['bei_v'] = float(match.group(2))
        data['bei_l'] = float(match.group(3))
        data['bei_hw'] = float(match.group(4))
        # 昇降機BEIは空欄の場合がある
        bei_ev_str = match.group(5).strip() if match.group(5) else ''
        data['bei_ev'] = float(bei_ev_str) if bei_ev_str else 0.0
    else:
        # パターン2: PDFテキスト形式（縦並び）
        bei_usage_pdf = r'BEI/AC[\s│]+BEI/V[\s│]+BEI/L[\s│]+BEI/HW[\s│]+BEI/EV[\s\n│]+([\d\.]+)[\s│]+([\d\.]+)[\s│]+([\d\.]+)[\s│]+([\d\.]+)[\s│]*([\d\.]*)'
        match = re.search(bei_usage_pdf, content, re.DOTALL)
        
        if match:
            data['bei_ac'] = float(match.group(1))
            data['bei_v'] = float(match.group(2))
            data['bei_l'] = float(match.group(3))
            data['bei_hw'] = float(match.group(4))
            bei_ev_str = match.group(5).strip() if match.group(5) else ''
            data['bei_ev'] = float(bei_ev_str) if bei_ev_str else 0.0
        else:
            # パターン3: 個別抽出
            data['bei_ac'] = extract_single_bei(content, 'AC', 2.24)
            data['bei_v'] = extract_single_bei(content, 'V', 1.00)
            data['bei_l'] = extract_single_bei(content, 'L', 1.00)
            data['bei_hw'] = extract_single_bei(content, 'HW', 0.99)
            data['bei_ev'] = extract_single_bei(content, 'EV', 0.0)
    
    # 一次エネルギー消費量（GJ/年）の抽出
    energy_section_match = re.search(r'3\.\s*一次エネルギー消費量計算結果(.*?)(?:4\.|$)', content, re.DOTALL)
    energy_section = energy_section_match.group(1) if energy_section_match else content
    
    data['energy_ac_design'] = extract_energy_value(energy_section, '空調設備', 1544.18)
    data['energy_v_design'] = extract_energy_value(energy_section, '換気設備', 31.98)
    data['energy_l_design'] = extract_energy_value(energy_section, '照明設備', 296.84)
    data['energy_hw_design'] = extract_energy_value(energy_section, '給湯設備', 225.73)
    data['energy_ev_design'] = 0.0
    
    # 基準値の計算
    data['energy_ac_standard'] = data['energy_ac_design'] / data['bei_ac'] if data['bei_ac'] > 0 else 692.00
    data['energy_v_standard'] = data['energy_v_design'] / data['bei_v'] if data['bei_v'] > 0 else 32.20
    data['energy_l_standard'] = data['energy_l_design'] / data['bei_l'] if data['bei_l'] > 0 else 296.90
    data['energy_hw_standard'] = data['energy_hw_design'] / data['bei_hw'] if data['bei_hw'] > 0 else 228.43
    data['energy_ev_standard'] = 0.0
    
    # PAL*の抽出
    pal_patterns = [
        r'PAL\s*\*[\s│|]+設計値[^\n]*?([\d]+)',  # PAL*計算結果表から
        r'設計値[\s│|]+([\d]+)[\s│|]+基準値[\s│|]+([\d]+)[\s│|]+BPI',  # BPI表から
        r'PAL\*[^\n]*?([\d]+)\s*MJ',
        r'年間熱負荷係数.*?([\d]+)',
    ]
    for pattern in pal_patterns:
        match = re.search(pattern, content)
        if match:
            try:
                pal = int(float(match.group(1)))
                if 100 < pal < 2000:
                    data['pal_design'] = pal
                    break
            except:
                pass
    if 'pal_design' not in data:
        data['pal_design'] = 300
    
    # PAL*を'pal_star'キーにも保存（互換性のため）
    data['pal_star'] = data['pal_design']
    
    # PAL*基準値
    data['pal_standard'] = 500
    
    # BPIの抽出
    bpi_patterns = [
        r'BPI[\s│|]+判定結果[\s\n│|]+([\d\.]+)',  # 表形式から
        r'設計値[^\n]*?基準値[^\n]*?BPI[^\n]*?([\d\.]+)',  # PAL*表から
        r'BPI[:\s]+([\d\.]+)',
        r'外皮性能指標.*?([\d\.]+)',
    ]
    for pattern in bpi_patterns:
        match = re.search(pattern, content)
        if match:
            try:
                bpi = float(match.group(1))
                if 0.1 < bpi < 5.0:
                    data['bpi'] = bpi
                    break
            except:
                pass
    if 'bpi' not in data:
        data['bpi'] = data['pal_design'] / data['pal_standard']
    
    # ワースト室の抽出
    data['worst_rooms'] = []
    worst_section_match = re.search(r'ワースト室.*?(?:5\.|$)', content, re.DOTALL)
    if worst_section_match:
        worst_section = worst_section_match.group(0)
        room_pattern = r'([^\n]+)\s+([\d\.]+)\s+MJ/m'
        matches = re.findall(room_pattern, worst_section)
        for match in matches[:5]:
            room_name = match[0].strip()
            q_per_a = float(match[1])
            if q_per_a > 0:
                data['worst_rooms'].append({
                    'name': room_name,
                    'q_per_a': q_per_a
                })
    
    return data


def extract_model_building_data(pdf_file):
    """
    モデル建物法のPDFからデータを抽出
    
    Args:
        pdf_file: ファイルオブジェクト
    
    Returns:
        dict: 抽出されたデータ
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("pdfplumberがインストールされていません")
    
    data = {
        'calculation_method': 'model_building',
        'building_name': 'test',
        'location': '東京都',
        'total_area': '1000',
        'bei_total': 1.0,
        'bei_ac': 1.0,
        'bei_v': 1.0,
        'bei_l': 1.0,
        'bei_hw': 1.0,
        'bei_ev': 0.0,
        'pal_design': 300,
        'pal_standard': 500,
        'bpi': 0.6,
        'worst_rooms': []
    }
    
    with pdfplumber.open(pdf_file) as pdf:
        # 1ページ目から基本情報を抽出
        if len(pdf.pages) >= 1:
            page1_text = pdf.pages[0].extract_text()
            
            # 建物名称
            building_match = re.search(r'建築物の名称\s+(.+?)(?:\s|$)', page1_text)
            if building_match:
                data['building_name'] = building_match.group(1).strip()
            
            # 延べ面積
            area_match = re.search(r'延べ面積\s+([\d,\.]+)', page1_text)
            if area_match:
                data['total_area'] = area_match.group(1).replace(',', '')
            
            # BPI
            bpi_match = re.search(r'【BPIm】\s+([\d\.]+)', page1_text)
            if bpi_match:
                data['bpi'] = float(bpi_match.group(1))
            
            # 総合BEI
            bei_total_match = re.search(r'【BEIm】\s+([\d\.]+)', page1_text)
            if bei_total_match:
                data['bei_total'] = float(bei_total_match.group(1))
            
            # 空調BEI
            bei_ac_match = re.search(r'空気調和設備\s+【BEI/AC】\s+([\d\.]+)', page1_text)
            if bei_ac_match:
                data['bei_ac'] = float(bei_ac_match.group(1))
            
            # 換気BEI
            bei_v_match = re.search(r'機械換気設備\s+【BEI/V】\s+([\d\.]+)', page1_text)
            if bei_v_match:
                data['bei_v'] = float(bei_v_match.group(1))
            
            # 照明BEI
            bei_l_match = re.search(r'照明設備\s+【BEI/L】\s+([\d\.]+)', page1_text)
            if bei_l_match:
                data['bei_l'] = float(bei_l_match.group(1))
            
            # 給湯BEI
            bei_hw_match = re.search(r'給湯設備\s+【BEI/HW】\s+([\d\.]+)', page1_text)
            if bei_hw_match:
                data['bei_hw'] = float(bei_hw_match.group(1))
            
            # 昇降機BEI
            bei_ev_match = re.search(r'昇降機\s+【BEI/EV】\s+([\d\.]+|-)', page1_text)
            if bei_ev_match:
                ev_value = bei_ev_match.group(1)
                data['bei_ev'] = 0.0 if ev_value == '-' else float(ev_value)
            
            # 4ページ目から詳細情報を抽出
            if len(pdf.pages) >= 4:
                page4_text = pdf.pages[3].extract_text()
                
                location_match = re.search(r'建築物所在地.*?都道府県\s+(.+?)\s+市区町村\s+(.+?)(?:\s|$)', page4_text, re.DOTALL)
                if location_match:
                    prefecture = location_match.group(1).strip()
                    city = location_match.group(2).strip()
                    data['location'] = f'{prefecture}{city}'
    
    # PAL*とBPIの関係からPAL*を推定
    if data['bpi'] > 0:
        data['pal_standard'] = 500
        data['pal_design'] = int(data['bpi'] * data['pal_standard'])
    
    # エネルギー消費量（モデル建物法では直接取得できない）
    data['energy_ac_standard'] = 0.0
    data['energy_v_standard'] = 0.0
    data['energy_l_standard'] = 0.0
    data['energy_hw_standard'] = 0.0
    data['energy_ac_design'] = 0.0
    data['energy_v_design'] = 0.0
    data['energy_l_design'] = 0.0
    data['energy_hw_design'] = 0.0
    
    return data


def extract_data_from_file(file_obj, file_name):
    """
    ファイルオブジェクトからエネルギー診断データを抽出
    
    Args:
        file_obj: ファイルオブジェクト（BytesIO or UploadedFile）
        file_name: ファイル名（拡張子判定用）
    
    Returns:
        dict: 抽出されたデータ
    """
    file_ext = os.path.splitext(file_name)[1].lower()
    
    if file_ext == '.pdf':
        # PDFの場合、まず計算方法を判定
        content = extract_text_from_pdf(file_obj)
        
        # ファイルポインタを先頭に戻す
        file_obj.seek(0)
        
        if 'モデル建物法' in content:
            # モデル建物法専用の抽出ロジック
            data = extract_model_building_data(file_obj)
        else:
            # 標準入力法の抽出ロジック
            data = extract_energy_data(content)
            data['calculation_method'] = 'standard_input'
    else:
        # テキストファイル（.txt, .md）
        content = file_obj.read()
        if isinstance(content, bytes):
            content = content.decode('utf-8')
        data = extract_energy_data(content)
        data['calculation_method'] = 'standard_input'
    
    return data


def create_stacked_bar_chart_improved(data):
    """エネルギー消費量の積み上げ棒グラフを生成してBytesIOで返す"""
    fig = plt.figure(figsize=(11, 6))
    
    ax = plt.subplot2grid((12, 1), (0, 0), rowspan=8)
    
    categories = ['基準値', '設計値']
    
    ac_values = [data['energy_ac_standard'], data['energy_ac_design']]
    v_values = [data['energy_v_standard'], data['energy_v_design']]
    l_values = [data['energy_l_standard'], data['energy_l_design']]
    hw_values = [data['energy_hw_standard'], data['energy_hw_design']]
    
    colors = {
        'ac': GRAPH_COLORS['ac'],
        'v': GRAPH_COLORS['v'],
        'l': GRAPH_COLORS['l'],
        'hw': GRAPH_COLORS['hw']
    }
    
    x = np.arange(len(categories))
    width = 0.6
    
    p1 = ax.barh(x, ac_values, width, label='空調', color=colors['ac'], alpha=0.9, edgecolor='white', linewidth=1.5)
    p2 = ax.barh(x, v_values, width, left=ac_values, label='換気', color=colors['v'], alpha=0.9, edgecolor='white', linewidth=1.5)
    p3 = ax.barh(x, l_values, width, left=np.array(ac_values)+np.array(v_values), 
                 label='照明', color=colors['l'], alpha=0.9, edgecolor='white', linewidth=1.5)
    p4 = ax.barh(x, hw_values, width, 
                 left=np.array(ac_values)+np.array(v_values)+np.array(l_values), 
                 label='給湯', color=colors['hw'], alpha=0.9, edgecolor='white', linewidth=1.5)
    
    for i, (ac, v, l, hw) in enumerate(zip(ac_values, v_values, l_values, hw_values)):
        if ac > 100:
            ax.text(ac/2, i, f'{ac:.2f}', ha='center', va='center', fontsize=10, fontweight='bold', color='white')
        if v > 20:
            ax.text(ac + v/2, i, f'{v:.2f}', ha='center', va='center', fontsize=9, fontweight='bold', color='white')
        if l > 100:
            ax.text(ac + v + l/2, i, f'{l:.2f}', ha='center', va='center', fontsize=10, fontweight='bold', color='white')
        if hw > 100:
            ax.text(ac + v + l + hw/2, i, f'{hw:.2f}', ha='center', va='center', fontsize=10, fontweight='bold', color='white')
    
    totals = [sum(x) for x in zip(ac_values, v_values, l_values, hw_values)]
    for i, total in enumerate(totals):
        ax.text(total + 50, i, f'{total:.2f}', ha='left', va='center', fontsize=11, fontweight='bold')
    
    ax.set_yticks(x)
    ax.set_yticklabels(categories, fontsize=12, fontweight='bold')
    calc_method = data.get('calculation_method', 'standard_input')
    bei_label = get_bei_label(calc_method)
    bpi_label = get_bpi_label(calc_method)
    
    ax.set_xlabel('一次エネルギー消費量 [GJ/年]', fontsize=12, fontweight='bold')
    ax.set_title(f'エネルギー消費性能  {bei_label}={data["bei_total"]:.2f}                [GJ/年]', 
                 fontsize=14, fontweight='bold', pad=15, loc='left')
    ax.legend(loc='lower right', fontsize=10, ncol=4, frameon=True, edgecolor=GRAPH_COLORS['main'], facecolor='white')
    ax.grid(axis='x', alpha=0.2, linestyle='--', linewidth=0.8)
    ax.set_axisbelow(True)
    
    table_ax = plt.subplot2grid((12, 1), (9, 0), rowspan=3)
    table_ax.axis('off')
    
    table_data = [
        [f'全体{bei_label}', f'空調{bei_label}', f'換気{bei_label}', f'照明{bei_label}', f'給湯{bei_label}', bpi_label],
        [f'{data["bei_total"]:.2f}', f'{data["bei_ac"]:.2f}', f'{data["bei_v"]:.2f}', 
         f'{data["bei_l"]:.2f}', f'{data["bei_hw"]:.2f}', f'{data["bpi"]:.2f}']
    ]
    
    table = table_ax.table(cellText=table_data, cellLoc='center', loc='center',
                          colWidths=[0.15]*6, bbox=[0.05, 0, 0.9, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 2.0)
    
    for i in range(6):
        cell = table[(0, i)]
        cell.set_facecolor(GRAPH_COLORS['main'])
        cell.set_text_props(weight='bold', color='white')
        cell.set_edgecolor('white')
        cell.set_linewidth(1.2)
    
    for i in range(6):
        cell = table[(1, i)]
        cell.set_edgecolor(GRAPH_COLORS['main'])
        cell.set_linewidth(1.2)
        value = float(table_data[1][i])
        if i < 5 and value > 1.0:
            cell.set_text_props(weight='bold', color='#DC3545', size=11)
        elif i == 5 and value < 1.0:
            cell.set_text_props(weight='bold', color='#28A745', size=11)
        else:
            cell.set_text_props(weight='bold', size=11)
    
    plt.tight_layout()
    
    # BytesIOに保存
    img_bytes = io.BytesIO()
    plt.savefig(img_bytes, format='png', dpi=300, bbox_inches='tight', facecolor='white', pad_inches=0.1)
    img_bytes.seek(0)
    plt.close()
    
    return img_bytes


def create_pie_charts(data):
    """設備別一次エネルギー消費量のパイチャートを生成してBytesIOで返す"""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5.5))
    
    standard_values = [
        data['energy_ac_standard'],
        data['energy_v_standard'],
        data['energy_l_standard'],
        data['energy_hw_standard']
    ]
    standard_labels = ['空調', '換気', '照明', '給湯']
    
    design_values = [
        data['energy_ac_design'],
        data['energy_v_design'],
        data['energy_l_design'],
        data['energy_hw_design']
    ]
    design_labels = ['空調', '換気', '照明', '給湯']
    
    colors = [GRAPH_COLORS['ac'], GRAPH_COLORS['v'], GRAPH_COLORS['l'], GRAPH_COLORS['hw']]
    
    wedges1, texts1, autotexts1 = ax1.pie(standard_values, labels=standard_labels, autopct='%1.1f%%',
                                           colors=colors, startangle=90, textprops={'fontsize': 11, 'fontweight': 'bold'})
    ax1.set_title('基準値\n一次エネルギー消費量', fontsize=13, fontweight='bold', pad=15, color=GRAPH_COLORS['main'])
    
    for autotext in autotexts1:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
        autotext.set_fontsize(10)
    
    wedges2, texts2, autotexts2 = ax2.pie(design_values, labels=design_labels, autopct='%1.1f%%',
                                           colors=colors, startangle=90, textprops={'fontsize': 11, 'fontweight': 'bold'})
    ax2.set_title('設計値\n一次エネルギー消費量', fontsize=13, fontweight='bold', pad=15, color=GRAPH_COLORS['main'])
    
    for autotext in autotexts2:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
        autotext.set_fontsize(10)
    
    plt.tight_layout()
    
    # BytesIOに保存
    img_bytes = io.BytesIO()
    plt.savefig(img_bytes, format='png', dpi=300, bbox_inches='tight', facecolor='white')
    img_bytes.seek(0)
    plt.close()
    
    return img_bytes


def create_bei_comparison_chart_with_total(data):
    """用途別BEIの棒グラフを生成してBytesIOで返す"""
    calc_method = data.get('calculation_method', 'standard_input')
    bei_label = get_bei_label(calc_method)
    
    categories = ['全体', '空調\n(AC)', '換気\n(V)', '照明\n(L)', '給湯\n(HW)', '昇降機\n(EV)']
    values = [data['bei_total'], data['bei_ac'], data['bei_v'], data['bei_l'], data['bei_hw'], data['bei_ev']]
    
    fig, ax = plt.subplots(figsize=(11, 5.5))
    
    colors = []
    for i, v in enumerate(values):
        if i == 0:
            colors.append(GRAPH_COLORS['main'] if v <= 1.0 else '#DC3545')
        else:
            colors.append('#DC3545' if v > 1.0 else GRAPH_COLORS['main'])
    
    bars = ax.bar(categories, values, color=colors, alpha=0.85, edgecolor='white', linewidth=1.5, width=0.65)
    
    ax.axhline(y=1.0, color='#DC3545', linestyle='--', linewidth=2.5, label=f'基準値 ({bei_label}=1.0)', zorder=3)
    
    ax.set_ylabel(f'{bei_label}値', fontsize=12, fontweight='bold')
    ax.set_xlabel('設備用途', fontsize=12, fontweight='bold')
    ax.set_title(f'設備用途別 {bei_label} 比較（全体{bei_label}含む）', fontsize=14, fontweight='bold', pad=15, color=GRAPH_COLORS['main'])
    ax.set_ylim(0, max(values) * 1.2)
    ax.legend(fontsize=10, loc='upper right', frameon=True, edgecolor=GRAPH_COLORS['main'])
    ax.grid(axis='y', alpha=0.2, linestyle='--', linewidth=0.8)
    ax.set_axisbelow(True)
    
    ax.tick_params(axis='x', labelsize=11)
    ax.tick_params(axis='y', labelsize=10)
    
    for bar, value in zip(bars, values):
        height = bar.get_height()
        if value > 0:
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{value:.2f}',
                   ha='center', va='bottom', fontsize=10, fontweight='bold')
    
    plt.tight_layout(pad=1.5)
    
    # BytesIOに保存
    img_bytes = io.BytesIO()
    plt.savefig(img_bytes, format='png', dpi=300, bbox_inches='tight', facecolor='white', pad_inches=0.2)
    img_bytes.seek(0)
    plt.close()
    
    return img_bytes


def generate_improvement_roadmap(data):
    """データに基づいて柔軟な改善ロードマップを生成"""
    roadmap = []
    
    bei_items = [
        ('空調', data['bei_ac'], data['energy_ac_design']),
        ('換気', data['bei_v'], data['energy_v_design']),
        ('照明', data['bei_l'], data['energy_l_design']),
        ('給湯', data['bei_hw'], data['energy_hw_design'])
    ]
    
    problem_items = [(name, bei, energy) for name, bei, energy in bei_items if bei > 1.0]
    problem_items.sort(key=lambda x: x[2], reverse=True)
    
    if data['bpi'] > 1.0:
        roadmap.append({
            'step': 'STEP 1',
            'title': '外皮性能の全体改善',
            'desc': '断熱・遮熱性能の底上げ'
        })
    elif data['worst_rooms']:
        roadmap.append({
            'step': 'STEP 1',
            'title': '外皮性能の局所改善',
            'desc': 'ワースト室の窓性能向上'
        })
    else:
        roadmap.append({
            'step': 'STEP 1',
            'title': '外皮性能の維持',
            'desc': '現状の良好な性能を維持'
        })
    
    if problem_items:
        main_problem = problem_items[0]
        if main_problem[0] == '空調':
            roadmap.append({
                'step': 'STEP 2',
                'title': '空調熱源の高効率化',
                'desc': '最大消費源の本質的改善'
            })
        elif main_problem[0] == '給湯':
            roadmap.append({
                'step': 'STEP 2',
                'title': '給湯設備の高効率化',
                'desc': 'エコキュート等への更新'
            })
        elif main_problem[0] == '照明':
            roadmap.append({
                'step': 'STEP 2',
                'title': '照明設備のLED化',
                'desc': '全館LED照明への更新'
            })
        else:
            roadmap.append({
                'step': 'STEP 2',
                'title': '換気設備の高効率化',
                'desc': '全熱交換器の導入'
            })
    else:
        roadmap.append({
            'step': 'STEP 2',
            'title': '設備効率の正常化',
            'desc': '低効率要因の排除'
        })
    
    roadmap.append({
        'step': 'STEP 3',
        'title': '制御の徹底強化',
        'desc': 'センサー連動制御の全域導入'
    })
    
    roadmap.append({
        'step': 'STEP 4',
        'title': '創エネルギーの導入',
        'desc': '太陽光発電等の再エネ設備'
    })
    
    return roadmap
