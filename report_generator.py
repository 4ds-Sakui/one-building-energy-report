#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
データ抽出・評価モジュール (v1.4.11)
モデル建物法の詳細項目抽出、ZEB比較ロジック、標準入力法サンプル統合対応
"""

import re
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use("Agg")

# 日本語フォントの設定
matplotlib.rcParams["font.family"] = "Noto Sans CJK JP"
matplotlib.rcParams["font.sans-serif"] = ["Noto Sans CJK JP"]

# カラー定義
COLOR_MAIN = "#397577"
COLOR_RED = "#e76f51"
COLOR_GREEN = "#2a9d8f"
COLOR_ACCENT = "#F4A261"
COLOR_GRAY = "#999999"

def extract_data_from_markdown(content):
    """
    Markdownからデータを抽出する
    """
    data = {
        'building_name': '不明',
        'total_area': 0.0,
        'location': '不明',
        'region': '不明',
        'solar_region': '不明',
        'building_model': '不明',
        'calculation_method': 'standard_input',
        'bei_total': 1.0,
        'bpi': 1.0,
        'bei_ac': 1.0,
        'bei_v': 1.0,
        'bei_l': 1.0,
        'bei_hw': 1.0,
        'bei_ev': 1.0,
        'solar_pv': 'なし',
        'cgs': 'なし',
        'judgment': {},
        'envelope_details': {},
        'equipment_details': {},
        'energy_consumption': {}
    }

    # モデル建物法かどうかの判定
    if "モデル建物法" in content:
        data['calculation_method'] = 'model_building'

    # 基本情報の抽出
    m = re.search(r'建築物の名称\s*\n\s*(.*)', content)
    if m: data['building_name'] = m.group(1).strip()
    
    m = re.search(r'床面積\s*\n\s*([\d,.]+)', content)
    if m: data['total_area'] = float(m.group(1).replace(',', ''))
    
    m = re.search(r'所在地[:：]\s*(.*)', content)
    if m: data['location'] = m.group(1).strip()

    m = re.search(r'地域区分/年間日射地域区分\s*\n\s*(.*)/(.*)', content)
    if m:
        data['region'] = m.group(1).strip()
        data['solar_region'] = m.group(2).strip()

    m = re.search(r'モデル建物\s*\n\s*(.*)', content)
    if m: data['building_model'] = m.group(1).strip()

    # BEI/BPIの抽出
    m = re.search(r'年間熱負荷係数\s*【BPIm?】\s*\|\s*([\d.]+)', content)
    if m: data['bpi'] = float(m.group(1))
    
    m = re.search(r'一次エネルギー消費量\s*【BEIm?】\s*\|\s*([\d.]+)', content)
    if m: data['bei_total'] = float(m.group(1))

    m = re.search(r'【誘導BEIm】\s*\|\s*([\d.]+)', content)
    if m: data['bei_target'] = float(m.group(1))

    # 設備別BEI
    m = re.search(r'空気調和設備\s*【BEIm?/AC】\s*\|\s*([\d.]+)', content)
    if m: data['bei_ac'] = float(m.group(1))
    m = re.search(r'機械換気設備\s*【BEIm?/V】\s*\|\s*([\d.]+)', content)
    if m: data['bei_v'] = float(m.group(1))
    m = re.search(r'照明設備\s*【BEIm?/L】\s*\|\s*([\d.]+)', content)
    if m: data['bei_l'] = float(m.group(1))
    m = re.search(r'給湯設備\s*【BEIm?/HW】\s*\|\s*([\d.]+)', content)
    if m: data['bei_hw'] = float(m.group(1))
    m = re.search(r'昇降機\s*【BEIm?/EV】\s*\|\s*([\d.]+)', content)
    if m: data['bei_ev'] = float(m.group(1))

    m = re.search(r'太陽光発電\s*\|\s*(.*)', content)
    if m: data['solar_pv'] = m.group(1).strip()
    m = re.search(r'コージェネレーション設備\s*\|\s*(.*)', content)
    if m: data['cgs'] = m.group(1).strip()

    # 判定結果
    data['judgment'] = {
        'base': '達成' if data['bei_total'] <= 1.0 else '非達成',
        'large': '達成' if data['bei_total'] <= 0.8 else '非達成',
        'target': '達成' if data.get('bei_target', 1.0) <= 0.6 else '非達成'
    }

    # 数値変換を試みるヘルパー関数
    def to_float_or_str(value):
        try:
            return float(value)
        except (ValueError, TypeError):
            return value

    # モデル建物法詳細項目の抽出 (PAL6-23)
    for code in range(6, 24):
        pattern = rf'PAL{code}\s*\|\s*[^|]*\|\s*([\d.]+)'
        m = re.search(pattern, content)
        if m: data['envelope_details'][f'PAL{code}'] = to_float_or_str(m.group(1))

    # 空調詳細 (AC1, AC4, AC6, AC7, AC10, AC12, AC13)
    for code in [1, 4, 6, 7, 10, 12, 13]:
        pattern = rf'AC{code}\s*\|\s*[^|]*\|\s*([^|\n]*)'
        m = re.search(pattern, content)
        if m: data['equipment_details'][f'AC{code}'] = to_float_or_str(m.group(1).strip())

    # 換気 (V5-7)
    v_sections = [("機械室", "V_機械室"), ("便所", "V_便所"), ("駐車場", "V_駐車場"), ("厨房", "V_厨房")]
    for section_name, key_name in v_sections:
        section_pattern = rf'\*\*{section_name}\*\*\n.*?\| V5 \|.*?\| ([^|\n]*) \|.*?\| V6 \|.*?\| ([^|\n]*) \|.*?\| V7 \|.*?\| ([^|\n]*) \|'
        m = re.search(section_pattern, content, re.DOTALL)
        if m:
            data['equipment_details'][key_name] = {
                'V5': m.group(1).strip(),
                'V6': m.group(2).strip(),
                'V7': m.group(3).strip()
            }

    # 照明 (L4-7)
    l_pattern = r'\| L4 \|.*?\| ([^|\n]*) \|.*?\| L5 \|.*?\| ([^|\n]*) \|.*?\| L6 \|.*?\| ([^|\n]*) \|.*?\| L7 \|.*?\| ([^|\n]*) \|'
    m = re.search(l_pattern, content, re.DOTALL)
    if m:
        data['equipment_details']['L'] = {
            'L4': m.group(1).strip(),
            'L5': m.group(2).strip(),
            'L6': m.group(3).strip(),
            'L7': m.group(4).strip()
        }

    # 給湯 (HW4-5)
    hw_sections = [("洗面手洗い", "HW_洗面手洗い"), ("浴室", "HW_浴室"), ("厨房", "HW_厨房")]
    for section_name, key_name in hw_sections:
        section_pattern = rf'\*\*{section_name}\*\*\n.*?\| HW4 \|.*?\| ([^|\n]*) \|.*?\| HW5 \|.*?\| ([^|\n]*) \|'
        m = re.search(section_pattern, content, re.DOTALL)
        if m:
            data['equipment_details'][key_name] = {
                'HW4': m.group(1).strip(),
                'HW5': m.group(2).strip()
            }

    return data

def get_zeb_comparison(data):
    """
    ZEB化相当との比較データを生成
    """
    comparison = []
    
    # 外皮性能
    u_wall = data['envelope_details'].get('PAL12', None)
    comparison.append({
        'category': '外壁U値',
        'current': f"{u_wall:.2f}" if isinstance(u_wall, (int, float)) else str(u_wall),
        'zeb_target': "0.60以下",
        'status': '良好' if (isinstance(u_wall, (int, float)) and u_wall <= 0.6) else '要改善',
        'action': '断熱材の厚肉化'
    })
    
    u_window = data['envelope_details'].get('PAL20', None)
    comparison.append({
        'category': '窓U値',
        'current': f"{u_window:.2f}" if isinstance(u_window, (int, float)) else str(u_window),
        'zeb_target': "2.33以下",
        'status': '良好' if (isinstance(u_window, (int, float)) and u_window <= 2.33) else '要改善',
        'action': 'Low-E複層ガラス採用'
    })

    # 開口率の計算とZEB目標
    total_wall_area_net = sum(data['envelope_details'].get(f'PAL{i}', 0) for i in range(6, 10)) # PAL6-9: 外壁面積 (窓を除く)
    total_window_area = sum(data['envelope_details'].get(f'PAL{i}', 0) for i in range(15, 19)) # PAL15-18: 窓面積
    
    # 開口率 = 窓面積 / (外壁面積 + 窓面積)
    opening_ratio = (total_window_area / (total_wall_area_net + total_window_area) * 100) if (total_wall_area_net + total_window_area) > 0 else 0
    comparison.append({
        'category': '開口率',
        'current': f"{opening_ratio:.1f}%",
        'zeb_target': "30%以下",
        'status': '良好' if opening_ratio <= 30 else '要改善',
        'action': '窓面積の削減、高断熱化'
    })

    # 空調
    ac_type = data['equipment_details'].get('AC1', '不明')
    comparison.append({
        'category': '主たる熱源',
        'current': ac_type,
        'zeb_target': "高効率ヒートポンプ等",
        'status': '良好' if "ヒートポンプ" in str(ac_type) or "エアコン" in str(ac_type) else '要検討',
        'action': '電気式高効率ヒートポンプへの転換'
    })
    
    ac_efficiency = data['equipment_details'].get('AC6', None)
    comparison.append({
        'category': '熱源効率 (AC6)',
        'current': f"{ac_efficiency:.2f}" if isinstance(ac_efficiency, (int, float)) else str(ac_efficiency),
        'zeb_target': "1.2以上",
        'status': '良好' if (isinstance(ac_efficiency, (int, float)) and ac_efficiency >= 1.2) else '要改善',
        'action': '高効率熱源機の導入'
    })
    
    ac_total_heat_exchanger = data['equipment_details'].get('AC13', '無')
    comparison.append({
        'category': '全熱交換器',
        'current': ac_total_heat_exchanger,
        'zeb_target': "有",
        'status': '良好' if str(ac_total_heat_exchanger) == '有' else '要検討',
        'action': '全熱交換器の導入'
    })

    # 換気 (V5-7)
    v_control_machine_room = data['equipment_details'].get('V_機械室', {}).get('V7', '無')
    comparison.append({
        'category': '換気制御 (機械室)',
        'current': v_control_machine_room,
        'zeb_target': "有",
        'status': '良好' if v_control_machine_room == '有' else '要検討',
        'action': '送風量制御の導入'
    })
    # V5-7の他の制御も追加
    v_control_toilet = data['equipment_details'].get('V_便所', {}).get('V7', '無')
    comparison.append({
        'category': '換気制御 (便所)',
        'current': v_control_toilet,
        'zeb_target': "有",
        'status': '良好' if v_control_toilet == '有' else '要検討',
        'action': '送風量制御の導入'
    })
    v_control_parking = data['equipment_details'].get('V_駐車場', {}).get('V7', '無')
    comparison.append({
        'category': '換気制御 (駐車場)',
        'current': v_control_parking,
        'zeb_target': "有",
        'status': '良好' if v_control_parking == '有' else '要検討',
        'action': '送風量制御の導入'
    })
    v_control_kitchen = data['equipment_details'].get('V_厨房', {}).get('V7', '無')
    comparison.append({
        'category': '換気制御 (厨房)',
        'current': v_control_kitchen,
        'zeb_target': "有",
        'status': '良好' if v_control_kitchen == '有' else '要検討',
        'action': '送風量制御の導入'
    })

    # 照明 (L4-7)
    l_occupancy_sensor = data['equipment_details'].get('L', {}).get('L4', '無')
    l_brightness_sensor = data['equipment_details'].get('L', {}).get('L5', '無')
    comparison.append({
        'category': '照明制御 (在室検知)',
        'current': l_occupancy_sensor,
        'zeb_target': "有",
        'status': '良好' if l_occupancy_sensor == '有' else '要検討',
        'action': '人感センサーの導入'
    })
    comparison.append({
        'category': '照明制御 (明るさ)',
        'current': l_brightness_sensor,
        'zeb_target': "有",
        'status': '良好' if l_brightness_sensor == '有' else '要検討',
        'action': '昼光利用制御の導入'
    })
    # L6, L7も追加
    l_time_control = data['equipment_details'].get('L', {}).get('L6', '無')
    comparison.append({
        'category': '照明制御 (時間)',
        'current': l_time_control,
        'zeb_target': "有",
        'status': '良好' if l_time_control == '有' else '要検討',
        'action': '時間制御の導入'
    })
    l_partial_lighting = data['equipment_details'].get('L', {}).get('L7', '無')
    comparison.append({
        'category': '照明制御 (部分照明)',
        'current': l_partial_lighting,
        'zeb_target': "有",
        'status': '良好' if l_partial_lighting == '有' else '要検討',
        'action': '部分照明の導入'
    })

    # 給湯 (HW4-5)
    hw_washroom_saving = data['equipment_details'].get('HW_洗面手洗い', {}).get('HW5', '無')
    comparison.append({
        'category': '給湯設備 (洗面節湯)',
        'current': hw_washroom_saving,
        'zeb_target': "有",
        'status': '良好' if hw_washroom_saving == '有' else '要検討',
        'action': '節湯器具の導入'
    })
    hw_bathroom_saving = data['equipment_details'].get('HW_浴室', {}).get('HW5', '無')
    comparison.append({
        'category': '給湯設備 (浴室節湯)',
        'current': hw_bathroom_saving,
        'zeb_target': "有",
        'status': '良好' if hw_bathroom_saving == '有' else '要検討',
        'action': '節湯器具の導入'
    })
    hw_kitchen_saving = data['equipment_details'].get('HW_厨房', {}).get('HW5', '無')
    comparison.append({
        'category': '給湯設備 (厨房節湯)',
        'current': hw_kitchen_saving,
        'zeb_target': "有",
        'status': '良好' if hw_kitchen_saving == '有' else '要検討',
        'action': '節湯器具の導入'
    })

    return comparison

def create_radar_chart(data):
    """
    設備別BEImのレーダーチャートを作成
    """
    categories = ["空調", "換気", "照明", "給湯", "昇降機"]
    values = [
        data.get('bei_ac', 1.0),
        data.get('bei_v', 1.0),
        data.get('bei_l', 1.0),
        data.get('bei_hw', 1.0),
        data.get('bei_ev', 0.0)
    ]
    
    base_values = [1.0] * 5
    
    N = len(categories)
    angles = [n / float(N) * 2 * np.pi for n in range(N)]
    values += values[:1]
    base_values += base_values[:1]
    angles += angles[:1]
    
    fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
    
    ax.plot(angles, base_values, color=COLOR_GRAY, linewidth=1, linestyle='dashed', label='基準値(1.0)')
    ax.fill(angles, base_values, color=COLOR_GRAY, alpha=0.1)
    
    ax.plot(angles, values, color=COLOR_MAIN, linewidth=2, label='設計値')
    ax.fill(angles, values, color=COLOR_MAIN, alpha=0.25)
    
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontproperties='Noto Sans CJK JP')
    ax.set_ylim(0, max(max(values), 1.5))
    ax.legend(loc='upper right', bbox_to_anchor=(0.1, 0.1), prop={'family': 'Noto Sans CJK JP'})
    plt.title('設備別BEIm分析', fontproperties='Noto Sans CJK JP', fontsize=14, pad=20)
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=100)
    buf.seek(0)
    plt.close()
    return buf

def extract_standard_sample_data(content):
    """
    チラ見せ用に標準入力法サンプルからデータを抽出
    """
    # 今回は画像埋め込みのため、この関数は実質不要だが、インターフェースとして残す
    return {}
