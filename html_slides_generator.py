#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HTMLスライド生成モジュール (v1.4.10)
モデル建物法詳細分析、ZEB比較、標準入力法チラ見せ
"""

import base64
from report_generator import COLOR_MAIN, COLOR_RED, COLOR_GREEN, COLOR_ACCENT, get_zeb_comparison, create_radar_chart

def generate_html_slides(data, standard_sample_data=None):
    """
    Reveal.jsベースのHTMLスライドを生成する
    """
    radar_buf = create_radar_chart(data)
    radar_base64 = base64.b64encode(radar_buf.read()).decode("utf-8")
    
    zeb_comp = get_zeb_comparison(data)
    
    def get_badge(status):
        color = COLOR_GREEN if status == "達成" else COLOR_RED
        return '<span style="background-color: ' + color + '; color: white; padding: 2px 10px; border-radius: 5px; font-weight: bold;">' + status + '</span>'

    # 画像ファイルをbase64エンコード
    def get_image_base64(filepath):
        try:
            with open(filepath, "rb") as f:
                return base64.b64encode(f.read()).decode("utf-8")
        except FileNotFoundError:
            return ""

    individual_bpi_base64 = get_image_base64("/home/ubuntu/streamlit_app/individual_bpi.png")
    energy_breakdown_base64 = get_image_base64("/home/ubuntu/streamlit_app/energy_breakdown.png")
    energy_comparison_base64 = get_image_base64("/home/ubuntu/streamlit_app/energy_comparison.png")

    # 数値または文字列として安全に表示するためのヘルパー関数
    def format_value(value, fmt=None):
        if isinstance(value, (int, float)):
            if fmt:
                return f"{value:{fmt}}"
            return str(value)
        return str(value)

    # 事前計算: 方位別開口率
    opening_ratio_n = ((data["envelope_details"].get("PAL15", 0) / (data["envelope_details"].get("PAL6", 0) + data["envelope_details"].get("PAL15", 0))) * 100) if (data["envelope_details"].get("PAL6", 0) + data["envelope_details"].get("PAL15", 0)) > 0 else 0
    opening_ratio_e = ((data["envelope_details"].get("PAL16", 0) / (data["envelope_details"].get("PAL7", 0) + data["envelope_details"].get("PAL16", 0))) * 100) if (data["envelope_details"].get("PAL7", 0) + data["envelope_details"].get("PAL16", 0)) > 0 else 0
    opening_ratio_s = ((data["envelope_details"].get("PAL17", 0) / (data["envelope_details"].get("PAL8", 0) + data["envelope_details"].get("PAL17", 0))) * 100) if (data["envelope_details"].get("PAL8", 0) + data["envelope_details"].get("PAL17", 0)) > 0 else 0
    opening_ratio_w = ((data["envelope_details"].get("PAL18", 0) / (data["envelope_details"].get("PAL9", 0) + data["envelope_details"].get("PAL18", 0))) * 100) if (data["envelope_details"].get("PAL9", 0) + data["envelope_details"].get("PAL18", 0)) > 0 else 0

    # 事前計算: 外皮性能判定
    pal12 = data["envelope_details"].get("PAL12", 1.0)
    pal12_ok = isinstance(pal12, (int, float)) and pal12 <= 0.6
    pal12_badge = '✅' if pal12_ok else '⚠️'

    pal20 = data["envelope_details"].get("PAL20", 3.0)
    pal20_ok = isinstance(pal20, (int, float)) and pal20 <= 2.33
    pal20_badge = '✅' if pal20_ok else '⚠️'

    pal21 = data["envelope_details"].get("PAL21", 0.5)
    pal21_ok = isinstance(pal21, (int, float)) and pal21 <= 0.4
    pal21_badge = '✅' if pal21_ok else '⚠️'

    # 事前計算: 設備性能判定
    ac1 = data["equipment_details"].get("AC1", "-")
    ac6 = data["equipment_details"].get("AC6", "-")
    ac13 = data["equipment_details"].get("AC13", "無")
    l4 = data["equipment_details"].get("L", {}).get("L4", "無")
    l5 = data["equipment_details"].get("L", {}).get("L5", "無")
    v_machine = data["equipment_details"].get("V_機械室", {}).get("V7", "無")
    hw_bath = data["equipment_details"].get("HW_浴室", {}).get("HW5", "無")

    # 建物情報
    building_name = data["building_name"]
    total_area = data["total_area"]
    region = data["region"]
    solar_region = data["solar_region"]
    building_model = data["building_model"]
    bei_total = data["bei_total"]
    judgment_base = get_badge(data["judgment"]["base"])
    judgment_large = get_badge(data["judgment"]["large"])
    judgment_target = get_badge(data["judgment"]["target"])

    # HTML生成
    html_content = """<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="utf-8">
    <title>技術レポート - """ + building_name + """</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/4.3.1/reveal.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/4.3.1/theme/white.min.css">
    <style>
        @font-face {
            font-family: 'Noto Sans CJK JP';
            src: url('https://fonts.gstatic.com/ea/notosansjp/v5/NotoSansJP-Regular.woff2') format('woff2');
            font-weight: normal;
            font-style: normal;
        }
        body, .reveal { font-family: 'Noto Sans CJK JP', sans-serif; }
        :root { --r-main-color: """ + COLOR_MAIN + """; --r-heading-color: """ + COLOR_MAIN + """; }
        .reveal h1, .reveal h2, .reveal h3 { color: var(--r-heading-color); font-weight: bold; }
        .reveal section { font-size: 28px; text-align: left; }
        .title-slide { text-align: center !important; background-color: """ + COLOR_MAIN + """; color: white !important; }
        .title-slide h1, .title-slide h3 { color: white !important; }
        .card { background: #f8f9fa; padding: 20px; border-radius: 10px; border-left: 5px solid """ + COLOR_MAIN + """; margin-bottom: 20px; }
        .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        table { width: 100%; border-collapse: collapse; font-size: 0.8em; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: center; }
        th { background-color: """ + COLOR_MAIN + """; color: white; }
        .accent-text { color: """ + COLOR_ACCENT + """; font-weight: bold; }
        .benefit-card { background: #e7f3ff; padding: 15px; border-radius: 8px; border: 1px solid #b3d7ff; font-size: 0.9em; }
    </style>
</head>
<body>
    <div class="reveal">
        <div class="slides">
            
            <section class="title-slide">
                <h3 style="font-size: 0.8em; text-transform: lowercase;">one building</h3>
                <h1>技術レポート</h1>
                <p>""" + building_name + """</p>
                <p style="font-size: 0.6em;">作成日: 2026.02.13</p>
                <p style="font-size: 0.5em; position: absolute; bottom: 50px; right: 20px; color: rgba(255,255,255,0.8);">v1.4.10</p>
                <p style="font-size: 0.4em; position: absolute; bottom: 20px; width: 100%;">© 2026 one building</p>
            </section>

            <section>
                <h2>1. 総合評価サマリー</h2>
                <div class="grid">
                    <div>
                        <p><b>建物概要</b></p>
                        <ul style="font-size: 0.7em;">
                            <li>延床面積: """ + format_value(total_area, ",.0f") + """ m²</li>
                            <li>地域区分: """ + region + """ / """ + solar_region + """</li>
                            <li>モデル建物: """ + building_model + """</li>
                        </ul>
                        <div class="card">
                            <p><b>判定結果</b></p>
                            <table style="font-size: 0.7em;">
                                <tr><td>基準適合 (BEIm≦1.00)</td><td>""" + judgment_base + """</td></tr>
                                <tr><td>大規模基準 (BEIm≦0.80)</td><td>""" + judgment_large + """</td></tr>
                                <tr><td>誘導基準 (BEIm≦0.60)</td><td>""" + judgment_target + """</td></tr>
                            </table>
                        </div>
                    </div>
                    <div style="text-align: center;">
                        <p><b>設備別BEIm分析</b></p>
                        <img src="data:image/png;base64,""" + radar_base64 + """" style="width: 80%;">
                    </div>
                </div>
                <p class="accent-text" style="text-align: center; margin-top: 20px;">💡 建物全体のBEImは""" + format_value(bei_total, ".2f") + """です。</p>
            </section>

            <section>
                <h2>2. 外皮性能の詳細分析</h2>
                <div class="grid">
                    <div style="font-size: 0.7em;">
                        <p><b>方位別面積・開口率</b></p>
                        <table>
                            <tr><th>方位</th><th>外壁面積</th><th>窓面積</th><th>開口率</th></tr>
                            <tr><td>北</td><td>""" + format_value(data["envelope_details"].get("PAL6", 0), ".1f") + """</td><td>""" + format_value(data["envelope_details"].get("PAL15", 0), ".1f") + """</td><td>""" + format_value(opening_ratio_n, ".1f") + """%</td></tr>
                            <tr><td>東</td><td>""" + format_value(data["envelope_details"].get("PAL7", 0), ".1f") + """</td><td>""" + format_value(data["envelope_details"].get("PAL16", 0), ".1f") + """</td><td>""" + format_value(opening_ratio_e, ".1f") + """%</td></tr>
                            <tr><td>南</td><td>""" + format_value(data["envelope_details"].get("PAL8", 0), ".1f") + """</td><td>""" + format_value(data["envelope_details"].get("PAL17", 0), ".1f") + """</td><td>""" + format_value(opening_ratio_s, ".1f") + """%</td></tr>
                            <tr><td>西</td><td>""" + format_value(data["envelope_details"].get("PAL9", 0), ".1f") + """</td><td>""" + format_value(data["envelope_details"].get("PAL18", 0), ".1f") + """</td><td>""" + format_value(opening_ratio_w, ".1f") + """%</td></tr>
                        </table>
                        <p style="margin-top: 10px;">※開口率は「外壁全体の面積に対する窓の割合」です。ZEBを目指す場合は30%以下を目標とします。</p>
                    </div>
                    <div>
                        <div class="card" style="font-size: 0.7em;">
                            <p><b>ZEB化相当との比較 (外皮)</b></p>
                            <table>
                                <tr><th>項目</th><th>現状値</th><th>ZEB目標</th><th>判定</th></tr>
                                <tr><td>外壁U値</td><td>""" + format_value(pal12, ".2f") + """</td><td>0.60以下</td><td>""" + pal12_badge + """</td></tr>
                                <tr><td>窓U値</td><td>""" + format_value(pal20, ".2f") + """</td><td>2.33以下</td><td>""" + pal20_badge + """</td></tr>
                                <tr><td>窓η値</td><td>""" + format_value(pal21, ".2f") + """</td><td>0.40以下</td><td>""" + pal21_badge + """</td></tr>
                            </table>
                            <p style="margin-top: 10px;"><b>推奨策:</b> Low-E複層ガラスへの変更、断熱材の厚肉化を検討してください。</p>
                        </div>
                    </div>
                </div>
            </section>

            <section>
                <h2>3. 設備性能の詳細分析</h2>
                <div style="font-size: 0.7em;">
                    <div class="grid">
                        <div class="card">
                            <p><b>空調設備 (AC)</b></p>
                            <ul>
                                <li>主熱源(冷): """ + format_value(ac1, "") + """ (ZEB目標: 高効率HP)</li>
                                <li>熱源効率(冷): """ + format_value(ac6, ".2f") + """ (ZEB目標: 1.2以上)</li>
                                <li>全熱交換器: """ + format_value(ac13, "") + """ (ZEB目標: 有)</li>
                            </ul>
                        </div>
                        <div class="card">
                            <p><b>照明・換気・給湯</b></p>
                            <ul style="font-size: 0.9em;">
                                <li>照明制御: 在室検知:""" + format_value(l4, "") + """, 明るさ:""" + format_value(l5, "") + """ (ZEB目標: 両方有)</li>
                                <li>換気制御: 送風量制御:""" + format_value(v_machine, "") + """ (ZEB目標: 有)</li>
                                <li>給湯仕様: 浴室節湯器具:""" + format_value(hw_bath, "") + """ (ZEB目標: 有)</li>
                            </ul>
                        </div>
                    </div>
                    <p class="accent-text">💡 ZEB Ready(0.50以下)達成には、高効率ヒートポンプへの転換と全熱交換器の導入が必須です。</p>
                </div>
            </section>

            <section style="background-color: #f0f4f8;">
                <h2 style="text-align: center;">4. さらなる価値へ：標準入力法のご案内</h2>
                <p class="accent-text" style="text-align: center; margin-bottom: 20px;">モデル建物法では見えない「真の課題」を、標準入力法で可視化</p>
                <div style="display: grid; grid-template-columns: 0.9fr 1fr 1.1fr; gap: 10px; grid-template-rows: auto auto;">
                    <!-- 左列: 経営的メリット -->
                    <div style="grid-column: 1; grid-row: 1 / 3; background: white; padding: 12px; border-radius: 8px; border-left: 4px solid """ + COLOR_ACCENT + """; font-size: 0.65em;">
                        <p style="margin: 0 0 8px 0; font-weight: bold; color: """ + COLOR_MAIN + """; font-size: 0.85em;">💰 経営的メリット</p>
                        <ul style="margin: 0; padding-left: 16px; text-align: left; line-height: 1.3;">
                            <li>光熱費削減額の正確な算出</li>
                            <li>投資回収期間の明確化</li>
                            <li>資産価値向上の定量評価</li>
                            <li>ZEB認定による企業価値向上</li>
                        </ul>
                    </div>
                    <!-- 中央上: 基準値と設計値の比較 -->
                    <div style="grid-column: 2; grid-row: 1; text-align: center;">
                        <p style="font-size: 0.6em; color: #666; margin: 0 0 6px 0;"><b>基準値と設計値の比較</b></p>
                        """ + (f'<img src="data:image/png;base64,{energy_comparison_base64}" style="width: 100%; height: auto;">' if energy_comparison_base64 else '<p style="color: red;">画像が見つかりません</p>') + """
                    </div>
                    <!-- 中央下: 設備別エネルギー消費内訳 -->
                    <div style="grid-column: 2; grid-row: 2; text-align: center;">
                        <p style="font-size: 0.6em; color: #666; margin: 0 0 6px 0;"><b>設備別エネルギー消費内訳</b></p>
                        """ + (f'<img src="data:image/png;base64,{energy_breakdown_base64}" style="width: 100%; height: auto;">' if energy_breakdown_base64 else '<p style="color: red;">画像が見つかりません</p>') + """
                    </div>
                    <!-- 右列: 室別の外皮性能評価 -->
                    <div style="grid-column: 3; grid-row: 1 / 3; text-align: center;">
                        <p style="font-size: 0.6em; color: #666; margin: 0 0 6px 0;"><b>室別の外皮性能評価</b></p>
                        """ + (f'<img src="data:image/png;base64,{individual_bpi_base64}" style="width: 100%; height: auto;">' if individual_bpi_base64 else '<p style="color: red;">画像が見つかりません</p>') + """
                    </div>
                </div>
            </section>

        </div>
    </div>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/4.3.1/reveal.min.js"></script>
    <script>
        Reveal.initialize({
            hash: true,
            center: true,
            transition: 'slide',
            width: 1280,
            height: 720,
            margin: 0.1
        });
    </script>
</body>
</html>
"""
    return html_content
