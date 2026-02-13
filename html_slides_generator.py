#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HTMLスライド生成モジュール (v1.4.3)
モデル建物法詳細分析、ZEB比較、標準入力法チラ見せ、組織自立診断
"""

import base64
from report_generator import COLOR_MAIN, COLOR_RED, COLOR_GREEN, COLOR_ACCENT, get_zeb_comparison, create_radar_chart

def generate_html_slides(data, standard_sample_data=None):
    """
    Reveal.jsベースのHTMLスライドを生成する
    """
    radar_buf = create_radar_chart(data)
    radar_base64 = base64.b64encode(radar_buf.read()).decode('utf-8')
    
    zeb_comp = get_zeb_comparison(data)
    
    def get_badge(status):
        color = COLOR_GREEN if status == "達成" else COLOR_RED
        return f'<span style="background-color: {color}; color: white; padding: 2px 10px; border-radius: 5px; font-weight: bold;">{status}</span>'

    html_content = f"""
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="utf-8">
    <title>技術レポート - {data['building_name']}</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/4.3.1/reveal.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/4.3.1/theme/white.min.css">
    <style>
        :root {{ --r-main-color: {COLOR_MAIN}; --r-heading-color: {COLOR_MAIN}; }}
        .reveal h1, .reveal h2, .reveal h3 {{ color: var(--r-heading-color); font-weight: bold; }}
        .reveal section {{ font-size: 28px; text-align: left; }}
        .title-slide {{ text-align: center !important; background-color: {COLOR_MAIN}; color: white !important; }}
        .title-slide h1, .title-slide h3 {{ color: white !important; }}
        .card {{ background: #f8f9fa; padding: 20px; border-radius: 10px; border-left: 5px solid {COLOR_MAIN}; margin-bottom: 20px; }}
        .grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 0.8em; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: center; }}
        th {{ background-color: {COLOR_MAIN}; color: white; }}
        .accent-text {{ color: {COLOR_ACCENT}; font-weight: bold; }}
        .benefit-card {{ background: #e7f3ff; padding: 15px; border-radius: 8px; border: 1px solid #b3d7ff; font-size: 0.9em; }}
        .treemap-box {{ display: flex; flex-wrap: wrap; gap: 10px; margin-top: 10px; }}
        .treemap-item {{ color: white; padding: 10px; border-radius: 5px; text-align: center; flex-grow: 1; }}
    </style>
</head>
<body>
    <div class="reveal">
        <div class="slides">
            
            <section class="title-slide">
                <h3>one building</h3>
                <h1>技術レポート</h1>
                <p>{data['building_name']}</p>
                <p style="font-size: 0.6em;">作成日: 2026.02.13</p>
                <p style="font-size: 0.4em; position: absolute; bottom: 20px; width: 100%;">© 2026 one building</p>
            </section>

            <section>
                <h2>1. 総合評価サマリー</h2>
                <div class="grid">
                    <div>
                        <p><b>建物概要</b></p>
                        <ul style="font-size: 0.7em;">
                            <li>延床面積: {data['total_area']:,} m²</li>
                            <li>地域区分: {data['region']} / {data['solar_region']}</li>
                            <li>モデル建物: {data['building_model']}</li>
                        </ul>
                        <div class="card">
                            <p><b>判定結果</b></p>
                            <table style="font-size: 0.7em;">
                                <tr><td>基準適合 (BEIm≦1.00)</td><td>{get_badge(data['judgment']['base'])}</td></tr>
                                <tr><td>大規模基準 (BEIm≦0.80)</td><td>{get_badge(data['judgment']['large'])}</td></tr>
                                <tr><td>誘導基準 (BEIm≦0.60)</td><td>{get_badge(data['judgment']['target'])}</td></tr>
                            </table>
                        </div>
                    </div>
                    <div style="text-align: center;">
                        <p><b>設備別BEIm分析</b></p>
                        <img src="data:image/png;base64,{radar_base64}" style="width: 80%;">
                    </div>
                </div>
            </section>

            <section>
                <h2>2. 外皮性能の詳細分析</h2>
                <div class="grid">
                    <div style="font-size: 0.7em;">
                        <p><b>方位別面積・開口率</b></p>
                        <table>
                            <tr><th>方位</th><th>外壁面積</th><th>窓面積</th><th>開口率</th></tr>
                            <tr><td>北</td><td>{data['envelope_details'].get('PAL6', 0)}</td><td>{data['envelope_details'].get('PAL15', 0)}</td><td>{data['envelope_details'].get('PAL15', 0)/data['envelope_details'].get('PAL6', 1)*100:.1f}%</td></tr>
                            <tr><td>東</td><td>{data['envelope_details'].get('PAL7', 0)}</td><td>{data['envelope_details'].get('PAL16', 0)}</td><td>{data['envelope_details'].get('PAL16', 0)/data['envelope_details'].get('PAL7', 1)*100:.1f}%</td></tr>
                            <tr><td>南</td><td>{data['envelope_details'].get('PAL8', 0)}</td><td>{data['envelope_details'].get('PAL17', 0)}</td><td>{data['envelope_details'].get('PAL17', 0)/data['envelope_details'].get('PAL8', 1)*100:.1f}%</td></tr>
                            <tr><td>西</td><td>{data['envelope_details'].get('PAL9', 0)}</td><td>{data['envelope_details'].get('PAL18', 0)}</td><td>{data['envelope_details'].get('PAL18', 0)/data['envelope_details'].get('PAL9', 1)*100:.1f}%</td></tr>
                        </table>
                    </div>
                    <div>
                        <div class="card" style="font-size: 0.7em;">
                            <p><b>ZEB化相当との比較 (外皮)</b></p>
                            <table>
                                <tr><th>項目</th><th>現状値</th><th>ZEB目標</th><th>判定</th></tr>
                                <tr><td>外壁U値</td><td>{data['envelope_details'].get('PAL12', '-')}</td><td>0.60以下</td><td>{'✅' if data['envelope_details'].get('PAL12', 1.0) <= 0.6 else '⚠️'}</td></tr>
                                <tr><td>窓U値</td><td>{data['envelope_details'].get('PAL20', '-')}</td><td>2.33以下</td><td>{'✅' if data['envelope_details'].get('PAL20', 3.0) <= 2.33 else '⚠️'}</td></tr>
                                <tr><td>窓η値</td><td>{data['envelope_details'].get('PAL21', '-')}</td><td>0.40以下</td><td>{'✅' if data['envelope_details'].get('PAL21', 0.5) <= 0.4 else '⚠️'}</td></tr>
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
                                <li>主熱源(冷): {data['equipment_details'].get('AC1', '-')}</li>
                                <li>熱源効率: {data['equipment_details'].get('AC6', '-')} (ZEB目標: 1.2以上)</li>
                                <li>全熱交換器: {data['equipment_details'].get('AC13', '無')} (ZEB目標: 有)</li>
                            </ul>
                        </div>
                        <div class="card">
                            <p><b>照明・換気・給湯</b></p>
                            <ul style="font-size: 0.9em;">
                                <li>照明制御: 在室検知:{data['equipment_details'].get('L', {}).get('L4', '無')}, 明るさ:{data['equipment_details'].get('L', {}).get('L5', '無')}</li>
                                <li>換気制御: 送風量制御:{data['equipment_details'].get('V_機械室', {}).get('V7', '無')}</li>
                                <li>給湯仕様: 浴室節湯器具:{data['equipment_details'].get('HW_浴室', {}).get('HW5', '無')}</li>
                            </ul>
                        </div>
                    </div>
                    <p class="accent-text">💡 設備全体のBEImは{data['bei_total']}です。ZEB Ready(0.50以下)達成には、高効率ヒートポンプへの転換と全熱交換器の導入が必須です。</p>
                </div>
            </section>

            <section style="background-color: #f0f4f8;">
                <h2 style="text-align: center;">4. さらなる価値へ：標準入力法のご案内</h2>
                <div class="grid">
                    <div>
                        <p class="accent-text">モデル建物法では見えない「真の課題」</p>
                        <p style="font-size: 0.8em;">現在の「モデル建物法」はあくまで簡易計算です。標準入力法へアップグレードすることで、以下の詳細分析が可能になります。</p>
                        <div class="benefit-card">
                            <b>✅ メリット1: 室別の外皮性能評価</b><br>
                            どの部屋が熱損失の「犯人」か特定し、ピンポイントで改修。
                        </div>
                        <div class="benefit-card" style="margin-top: 10px;">
                            <b>✅ メリット2: 正確なLCC・投資回収計算</b><br>
                            GJ単位のエネルギー消費量から、光熱費削減額を算出。
                        </div>
                    </div>
                    <div style="text-align: center;">
                        <p style="font-size: 0.6em; color: #666;">標準入力法による詳細分析イメージ</p>
                        <div style="background: white; padding: 15px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                            <p style="font-size: 0.6em; font-weight: bold;">【室別BPIツリーマップ】</p>
                            <div class="treemap-box">
                                <div class="treemap-item" style="background: {COLOR_RED}; width: 60%;">事務室<br>1.43</div>
                                <div class="treemap-item" style="background: {COLOR_ACCENT}; width: 30%;">ロビー<br>1.10</div>
                                <div class="treemap-item" style="background: {COLOR_GREEN}; width: 20%;">客室<br>0.45</div>
                            </div>
                            <p style="font-size: 0.5em; color: #888; margin-top: 10px;">※標準入力法でのみ可能な分析例</p>
                        </div>
                    </div>
                </div>
            </section>

            <section class="title-slide" style="background-color: {COLOR_ACCENT};">
                <h2>組織自立診断へのご招待</h2>
                <p>技術的な「建物」の改善と同時に、推進する「組織」の力を診断しませんか？</p>
                <div style="background: white; color: black; padding: 20px; border-radius: 10px; margin-top: 30px; text-align: left; font-size: 0.8em;">
                    <p><b>組織エンゲージメント診断 (所要時間: 5分)</b></p>
                    <ul>
                        <li>経営層のZEBに対するコミットメント度</li>
                        <li>技術部門と営業部門の連携体制</li>
                        <li>社内推進チームの自立性</li>
                    </ul>
                    <p style="text-align: center; margin-top: 20px;">
                        <button style="padding: 10px 30px; background: {COLOR_MAIN}; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: bold;">診断をスタートする</button>
                    </p>
                </div>
            </section>

        </div>
    </div>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/4.3.1/reveal.min.js"></script>
    <script>
        Reveal.initialize({{
            hash: true,
            center: true,
            transition: 'slide',
            width: 1280,
            height: 720,
            margin: 0.1
        }});
    </script>
</body>
</html>
"""
    return html_content
