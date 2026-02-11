'''
HTMLスライド生成モジュール（reveal.js対応）
PowerPoint版と完全に同期した動的評価ロジック搭載版
'''

import base64
from io import BytesIO
import matplotlib.pyplot as plt
from datetime import datetime

# カラー定数の定義（report_generator.pyのRGBColor型との競合を避けるため文字列で定義）
COLOR_MAIN_HEX = "#397577"
COLOR_RED_HEX = "#E63946"
COLOR_GREEN_HEX = "#2A9D8F"
COLOR_GRAY_HEX = "#6D6D6D"
COLOR_BLACK_HEX = "#333333"

def get_bei_label_local(calc_method):
    return "BEIm" if calc_method == "モデル建物法" else "BEI"

def get_bpi_label_local(calc_method):
    return "BPIm" if calc_method == "モデル建物法" else "BPI"

def generate_html_slides(data, charts):
    """
    データとグラフからHTMLスライドを生成する
    """
    
    # グラフをBase64エンコード
    chart_images = {}
    for chart_name, chart_obj in charts.items():
        if chart_obj:
            if isinstance(chart_obj, BytesIO):
                chart_obj.seek(0)
                img_base64 = base64.b64encode(chart_obj.read()).decode('utf-8')
                chart_images[chart_name] = f"data:image/png;base64,{img_base64}"
            else:
                buf = BytesIO()
                chart_obj.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
                buf.seek(0)
                img_base64 = base64.b64encode(buf.read()).decode('utf-8')
                chart_images[chart_name] = f"data:image/png;base64,{img_base64}"
                plt.close(chart_obj)
    
    # 基本情報
    building_name = data.get('building_name', '不明')
    location = data.get('location', '不明')
    total_area = data.get('total_area', 'N/A')
    calc_method = data.get('calculation_method', 'standard_input')
    bei_total = data.get('bei_total', 0.0)
    bei_label = get_bei_label_local(calc_method)
    bpi_label = get_bpi_label_local(calc_method)
    
    # 1. 診断結果とZEBレベルの判定
    if bei_total > 1.0:
        status_text = "診断結果: 2024年基準非適合"
        status_color = COLOR_RED_HEX
        zeb_level = "ZEB基準未達成"
        is_compliant = False
    elif bei_total > 0.80:
        status_text = "診断結果: H28年基準適合 / 2024年基準非適合"
        status_color = COLOR_RED_HEX
        zeb_level = "ZEB基準未達成"
        is_compliant = False
    elif bei_total > 0.70:
        status_text = "診断結果: 2024年基準適合"
        status_color = COLOR_GREEN_HEX
        zeb_level = "ZEB Oriented相当（未認定）"
        is_compliant = True
    elif bei_total > 0.50:
        status_text = "診断結果: 2024年基準適合"
        status_color = COLOR_GREEN_HEX
        zeb_level = "ZEB Ready相当（未認定）"
        is_compliant = True
    else:
        status_text = "診断結果: 2024年基準適合"
        status_color = COLOR_GREEN_HEX
        zeb_level = "Nearly ZEB相当（未認定）"
        is_compliant = True

    # 2. 経営リスク/優位性の判定
    if not is_compliant:
        risk_title = "▲重要: 経営リスクの特定"
        risk_color = COLOR_RED_HEX
        risk_bg = "#FFF3F3"
        risk_main_desc = "2024年4月施行の改正省エネ法（BEI 1.0以下義務化）に対応していません。"
        risk_legal = "適合判定（省エネ適判）をパスできず、建築確認申請が受理されません。"
        risk_business = "著しく高い光熱費が継続し、運用コストが逼迫します。"
        risk_asset = "ESG投資基準を満たせず、企業価値が低下する可能性があります。"
    else:
        risk_title = "▲優位性: 省エネ性能の強み"
        risk_color = COLOR_GREEN_HEX
        risk_bg = "#F3FFF3"
        risk_main_desc = f"2024年基準に適合し、{zeb_level}の省エネ性能を有しています。"
        risk_legal = "省エネ適判をクリアし、建築確認申請がスムーズに進みます。"
        risk_business = "光熱費が抑えられ、長期的な運用コスト削減が期待できます。"
        risk_asset = "ESG投資基準を満たし、企業価値向上に貢献します。"

    # 3. 外皮性能評価
    pal_design = data.get('pal_design', 0.0)
    pal_standard = data.get('pal_standard', 1.0)
    bpi_val = data.get('bpi', 0.0)
    bpi_status = "良好" if bpi_val < 1.0 else "要改善"
    bpi_color = COLOR_GREEN_HEX if bpi_val < 1.0 else COLOR_RED_HEX
    
    # 4. ワースト室分析
    worst_rooms = data.get('worst_rooms', [])
    if worst_rooms:
        wr = worst_rooms[0]
        worst_room_name = wr.get('name', '不明')
        worst_room_q = wr.get('q_per_a', 0.0)
        over_rate = ((worst_room_q - pal_standard) / pal_standard * 100) if pal_standard > 0 else 0
        
        if 'ミーティング' in worst_room_name or '会議' in worst_room_name:
            load_factors = "南側または西側の大きな窓面積による日射負荷、高い人員密度と内部発発熱、窓の遮熱性能不足"
        elif '事務' in worst_room_name:
            load_factors = "OA機器による内部発熱、窓面積比率が高い可能性、断熱性能の不足"
        else:
            load_factors = "窓の性能（遮熱・断熱）の不足、外壁断熱材の不足、日射負荷の増大"
        
        worst_analysis_html = f'''
            <div class="worst-factor-item">
                <div class="worst-factor-label">最大ワースト室: {worst_room_name}</div>
                <div class="worst-factor-value" style="color: {COLOR_RED_HEX};">超過率: +{over_rate:.0f}%</div>
                <p class="worst-factor-description">
                    <strong>負荷要因:</strong> {load_factors}<br>
                    <strong>改善策:</strong> 窓の遮熱性能向上（Low-Eガラス、外付けブラインド等）により、建物全体のPAL*を大幅に改善可能な優先箇所です。
                </p>
            </div>
        '''
    else:
        worst_analysis_html = f'''
            <div class="worst-factor-item">
                <p class="worst-factor-description">
                    外皮性能は全体的に良好です。設備側の改善に注力することで、{bei_label}目標達成が可能です。
                </p>
            </div>
        '''

    # 5. ロードマップ効果
    target_bei = 0.70
    reduction = bei_total - target_bei
    roadmap_effect = f"現在の{bei_label} {bei_total:.2f} から ZEB Oriented 要件 {target_bei:.2f} まで、約 {reduction:.2f} の削減が必要です。"

    # 日付
    today = datetime.now().strftime('%Y年%m月%d日')
    
    html = f'''<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{building_name} - 省エネ診断レポート</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.0/dist/reset.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.0/dist/reveal.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.0/dist/theme/white.css">
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            --main-color: {COLOR_MAIN_HEX};
            --red-color: {COLOR_RED_HEX};
            --green-color: {COLOR_GREEN_HEX};
            --text-dark: {COLOR_BLACK_HEX};
            --text-light: {COLOR_GRAY_HEX};
            --base-bg: #f6f6f6;
        }}
        
        * {{ font-family: 'Noto Sans JP', sans-serif; }}
        
        .reveal {{ font-size: 24px; }}
        .reveal h1, .reveal h2, .reveal h3 {{ color: var(--main-color); font-weight: 700; text-transform: none; }}
        .reveal h1 {{ font-size: 1.8em; margin-bottom: 0.1em; }}
        .reveal h2 {{ font-size: 1.2em; border-bottom: 3px solid var(--main-color); padding-bottom: 0.2em; margin-bottom: 0.5em; }}
        
        .reveal section {{ text-align: left; padding: 1em; }}
        
        .title-slide {{ text-align: center !important; display: flex !important; flex-direction: column; justify-content: center; height: 100%; }}
        .title-box {{ background: var(--main-color); color: white; padding: 2em; margin: 0 auto; width: 60%; border-radius: 8px; }}
        .title-box h1 {{ color: white; margin: 0; font-size: 1.8em; }}
        .title-box .building-name {{ font-size: 1.2em; margin-top: 0.5em; opacity: 0.9; }}
        
        .footer {{ position: absolute; bottom: 1em; left: 0; right: 0; font-size: 0.6em; color: var(--text-light); text-align: center; }}
        
        .summary-grid {{ display: grid; grid-template-columns: 1.2fr 0.8fr; gap: 1em; }}
        .card {{ background: white; padding: 0.8em; border-radius: 6px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border: 1px solid #eee; }}
        .card h3 {{ font-size: 0.9em; margin-bottom: 0.4em; }}
        
        .risk-box {{ background: {risk_bg}; border: 2px solid {risk_color}; border-radius: 8px; padding: 1em; margin-top: 1em; }}
        .risk-title {{ color: {risk_color}; font-weight: 700; font-size: 1.1em; margin-bottom: 0.5em; }}
        .risk-desc {{ font-size: 0.85em; line-height: 1.4; }}
        .risk-list {{ font-size: 0.8em; margin-top: 0.5em; padding-left: 1.2em; }}
        .risk-list li {{ margin-bottom: 0.3em; }}
        
        .bei-display {{ text-align: center; }}
        .bei-val {{ font-size: 3em; font-weight: 700; color: {status_color}; line-height: 1; }}
        .bei-label {{ font-size: 0.8em; color: var(--text-light); }}
        
        .eq-list {{ font-size: 0.75em; list-style: none; padding: 0; }}
        .eq-list li {{ display: flex; justify-content: space-between; padding: 0.2em 0; border-bottom: 1px solid #eee; }}
        
        .pal-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 1em; }}
        .worst-factor-container {{ background: var(--base-bg); padding: 1em; border-radius: 6px; }}
        .worst-factor-label {{ font-weight: 700; color: var(--main-color); font-size: 0.9em; }}
        .worst-factor-value {{ font-size: 1.1em; font-weight: 700; margin: 0.2em 0; }}
        .worst-factor-description {{ font-size: 0.75em; color: var(--text-dark); line-height: 1.4; }}
        
        .chart-container {{ text-align: center; margin: 0.5em 0; }}
        .chart-container img {{ max-width: 95%; max-height: 400px; object-fit: contain; }}
        
        .roadmap-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 0.5em; margin-top: 1em; }}
        .roadmap-step {{ background: var(--base-bg); padding: 0.6em; border-radius: 6px; text-align: center; border: 1px solid var(--main-color); }}
        .step-num {{ color: var(--main-color); font-weight: 700; font-size: 0.8em; }}
        .step-title {{ font-weight: 700; font-size: 0.85em; margin: 0.3em 0; min-height: 2.4em; display: flex; align-items: center; justify-content: center; }}
        .step-desc {{ font-size: 0.65em; color: var(--text-light); line-height: 1.3; }}
        
        .effect-box {{ background: white; border: 1px dashed var(--main-color); padding: 0.8em; margin-top: 1em; text-align: center; font-size: 0.8em; }}
    </style>
</head>
<body>
    <div class="reveal">
        <div class="slides">
            <!-- 1. タイトル -->
            <section class="title-slide">
                <div class="title-box">
                    <div style="font-size: 0.8em; font-weight: 700; margin-bottom: 1em;">one building</div>
                    <div class="building-name">{building_name}</div>
                    <h1>技術レポート</h1>
                    <div style="margin-top: 1em; font-size: 0.7em;">{today}</div>
                </div>
                <div class="footer">
                    BIM sustaina for Energy | copyright © one building
                </div>
            </section>
            
            <!-- 2. 総合評価サマリー -->
            <section>
                <h2>1. 総合評価サマリー: 現状と経営リスク</h2>
                <div class="summary-grid">
                    <div>
                        <div style="font-size: 0.75em; color: var(--text-light); margin-bottom: 0.5em;">
                            所在地: {location} | 延べ面積: {total_area} m²
                        </div>
                        <div style="font-size: 1.1em; font-weight: 700; color: {status_color}; margin-bottom: 0.5em;">
                            {status_text}
                        </div>
                        <div class="risk-box">
                            <div class="risk-title">{risk_title}</div>
                            <div class="risk-desc"><strong>{risk_main_desc}</strong></div>
                            <ul class="risk-list">
                                <li><strong>法的リスク:</strong> {risk_legal}</li>
                                <li><strong>事業リスク:</strong> {risk_business}</li>
                                <li><strong>資産価値リスク:</strong> {risk_asset}</li>
                            </ul>
                        </div>
                    </div>
                    <div style="display: flex; flex-direction: column; gap: 0.8em;">
                        <div class="card bei-display">
                            <div class="bei-label">現在の {bei_label}</div>
                            <div class="bei-val">{bei_total:.2f}</div>
                        </div>
                        <div class="card">
                            <h3>設備別{bei_label}サマリー</h3>
                            <ul class="eq-list">
                                <li><span>空調</span> <span style="color: {'#2A9D8F' if data.get('bei_ac', 0) <= 1.0 else '#E63946'}">{data.get('bei_ac', 0):.2f}</span></li>
                                <li><span>換気</span> <span style="color: {'#2A9D8F' if data.get('bei_v', 0) <= 1.0 else '#E63946'}">{data.get('bei_v', 0):.2f}</span></li>
                                <li><span>照明</span> <span style="color: {'#2A9D8F' if data.get('bei_l', 0) <= 1.0 else '#E63946'}">{data.get('bei_l', 0):.2f}</span></li>
                                <li><span>給湯</span> <span style="color: {'#2A9D8F' if data.get('bei_hw', 0) <= 1.0 else '#E63946'}">{data.get('bei_hw', 0):.2f}</span></li>
                                <li><span>昇降機</span> <span style="color: {'#2A9D8F' if data.get('bei_ev', 0) <= 1.0 else '#E63946'}">{data.get('bei_ev', 0):.2f}</span></li>
                            </ul>
                        </div>
                    </div>
                </div>
            </section>

            <!-- 3. PAL*分析 -->
            <section>
                <h2>2. 外皮性能評価とPAL*ワースト要因分析</h2>
                <div class="pal-grid">
                    <div class="card" style="text-align: center;">
                        <div class="worst-factor-label">外皮性能指標 (PAL*)</div>
                        <div style="font-size: 2.5em; font-weight: 700; color: var(--main-color);">{pal_design} <span style="font-size: 0.4em;">MJ/m²・年</span></div>
                        <div style="font-size: 0.9em; font-weight: 700; color: {bpi_color}; margin-top: 0.5em;">
                            {bpi_label} {bpi_val:.2f} ({bpi_status})
                        </div>
                    </div>
                    <div class="worst-factor-container">
                        <div class="worst-factor-label" style="margin-bottom: 0.5em;">負荷要因の分解分析: ワースト室の特定</div>
                        {worst_analysis_html}
                    </div>
                </div>
            </section>

            <!-- 4. BEI比較 -->
            <section>
                <h2>3. 用途別エネルギー消費傾向: {bei_label}分析</h2>
                <div class="chart-container">
                    <img src="{chart_images.get('bei_chart', '')}" alt="BEI比較グラフ">
                </div>
            </section>

            <!-- 5. ロードマップ -->
            <section>
                <h2>4. ZEB化への改善ロードマップ</h2>
                <div class="roadmap-grid">
                    <div class="roadmap-step">
                        <div class="step-num">STEP 1</div>
                        <div class="step-title">外皮性能の維持</div>
                        <div class="step-desc">現在の断熱性能を維持しつつ、日射遮蔽を強化します。</div>
                    </div>
                    <div class="roadmap-step">
                        <div class="step-num">STEP 2</div>
                        <div class="step-title">高効率熱源の導入</div>
                        <div class="step-desc">空調熱源を最新の高効率機へ更新し、一次エネを削減。</div>
                    </div>
                    <div class="roadmap-step">
                        <div class="step-num">STEP 3</div>
                        <div class="step-title">照明・換気制御</div>
                        <div class="step-desc">LED化および人感・CO2センサーによる無駄な消費を抑制。</div>
                    </div>
                    <div class="roadmap-step">
                        <div class="step-num">STEP 4</div>
                        <div class="step-title">創エネ（PV）</div>
                        <div class="step-desc">太陽光発電の導入により、ZEB Ready以上の達成を目指す。</div>
                    </div>
                </div>
                <div class="effect-box">
                    <strong>期待される総合効果:</strong><br>
                    {roadmap_effect}<br>
                    上記ステップの実施により、段階的に{bei_label}目標達成を目指します。
                </div>
            </section>
        </div>
    </div>
    
    <script src="https://cdn.jsdelivr.net/npm/reveal.js@4.6.0/dist/reveal.js"></script>
    <script>
        Reveal.initialize({{
            hash: true,
            width: 1280,
            height: 720,
            margin: 0.05,
            minScale: 0.2,
            maxScale: 2.0,
            transition: 'slide',
            center: false,
            slideNumber: 'c/t',
            plugins: []
        }});
    </script>
</body>
</html>'''
    
    return html
