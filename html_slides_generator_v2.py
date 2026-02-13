"""
HTMLスライド生成モジュール v2
reveal.jsを使用してone buildingデザインガイドラインに準拠した
インタラクティブなスライド形式のレポートを生成
"""

import base64
from io import BytesIO
from datetime import datetime
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt


def generate_html_slides_v2(data, charts):
    """
    reveal.jsを使用したHTMLスライドを生成（デザインガイドライン準拠版）
    
    Args:
        data: 抽出されたデータ辞書
        charts: グラフ画像の辞書 {'bei_chart': fig1, 'energy_chart': fig2, 'pie_chart': fig3}
    
    Returns:
        str: HTML文字列
    """
    
    # グラフをBase64エンコード
    chart_images = {}
    for chart_name, fig in charts.items():
        if fig:
            buf = BytesIO()
            fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
            buf.seek(0)
            img_base64 = base64.b64encode(buf.read()).decode('utf-8')
            chart_images[chart_name] = f"data:image/png;base64,{img_base64}"
            plt.close(fig)
    
    # 建物情報
    building_name = data.get('building_name', '不明')
    building_type = data.get('building_type', '不明')
    location = data.get('location', '不明')
    total_area = data.get('total_area', 'N/A')
    calculation_method = data.get('calculation_method', '不明')
    
    # BEI/BEIm、BPI/BPIm表記の切り替え
    bei_label = "BEIm" if calculation_method == "モデル建物法" else "BEI"
    bpi_label = "BPIm" if calculation_method == "モデル建物法" else "BPI"
    
    # BEI値
    bei_total = data.get('bei_total', 'N/A')
    bei_ac = data.get('bei_ac', 'N/A')
    bei_v = data.get('bei_v', 'N/A')
    bei_l = data.get('bei_l', 'N/A')
    bei_hw = data.get('bei_hw', 'N/A')
    bei_ev = data.get('bei_ev', 'N/A')
    
    # その他の値
    bpi = data.get('bpi', 'N/A')
    pal = data.get('pal_star', data.get('pal_design', 'N/A'))
    
    # エネルギー消費量
    energy_data = data.get('energy_consumption', {})
    
    # 現在日時
    current_date = datetime.now().strftime('%Y.%m.%d')
    
    # 診断結果の判定
    try:
        bei_value = float(bei_total) if bei_total != 'N/A' else 0
        if bei_value <= 1.0:
            diagnosis_result = "2024年基準適合"
            diagnosis_class = "compliant"
        else:
            diagnosis_result = "2024年基準非適合"
            diagnosis_class = "non-compliant"
    except:
        diagnosis_result = "判定不可"
        diagnosis_class = "unknown"
    
    # HTMLテンプレート
    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{building_name} - 省エネ診断レポート</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.0/dist/reset.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.0/dist/reveal.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.0/dist/theme/white.css">
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&family=Prompt:wght@400;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            --main-color: #397577;
            --accent-color: #013E34;
            --base-color-1: #f6f6f6;
            --base-color-2: #ededed;
            --text-dark: #333333;
            --text-light: #666666;
        }}
        
        * {{
            font-family: 'Noto Sans JP', 'Prompt', sans-serif;
        }}
        
        .reveal {{
            font-family: 'Noto Sans JP', 'Prompt', sans-serif;
        }}
        
        .reveal h1, .reveal h2, .reveal h3 {{
            color: var(--main-color);
            font-weight: 700;
            text-transform: none;
        }}
        
        .reveal h1 {{
            font-size: 2.8em;
            margin-bottom: 0.3em;
        }}
        
        .reveal h2 {{
            font-size: 2.0em;
            margin-bottom: 0.8em;
            border-bottom: 3px solid var(--main-color);
            padding-bottom: 0.3em;
        }}
        
        .reveal h3 {{
            font-size: 1.5em;
        }}
        
        .reveal section {{
            text-align: left;
        }}
        
        .reveal .title-slide {{
            text-align: center;
            display: flex;
            flex-direction: column;
            justify-content: center;
            height: 100%;
            background: linear-gradient(135deg, var(--base-color-1) 0%, white 100%);
        }}
        
        .reveal .title-slide h1 {{
            font-size: 3.5em;
            margin-bottom: 0.2em;
            color: var(--text-dark);
        }}
        
        .reveal .title-slide .building-name {{
            font-size: 2.5em;
            color: var(--main-color);
            margin-bottom: 0.5em;
            font-weight: 700;
        }}
        
        .reveal .title-slide .subtitle {{
            font-size: 1.3em;
            color: var(--text-light);
            margin-bottom: 2em;
        }}
        
        .reveal .title-slide .date {{
            font-size: 1.0em;
            color: var(--text-light);
            margin-top: 2em;
        }}
        
        .reveal .title-slide .footer {{
            position: absolute;
            bottom: 2em;
            left: 0;
            right: 0;
            font-size: 0.8em;
            color: var(--text-light);
        }}
        
        /* 総合評価サマリースライド */
        .reveal .summary-slide {{
            background: linear-gradient(to right, var(--base-color-1) 0%, white 50%);
        }}
        
        .reveal .summary-content {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 2em;
            margin-top: 1.5em;
        }}
        
        .reveal .summary-left {{
            padding-right: 1em;
            border-right: 2px solid var(--main-color);
        }}
        
        .reveal .summary-left h3 {{
            color: var(--accent-color);
            margin-bottom: 1em;
        }}
        
        .reveal .diagnosis-box {{
            background: white;
            border-left: 5px solid var(--main-color);
            padding: 1.2em;
            margin: 1em 0;
            border-radius: 4px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        
        .reveal .diagnosis-box.compliant {{
            border-left-color: #28a745;
        }}
        
        .reveal .diagnosis-box.non-compliant {{
            border-left-color: #dc3545;
        }}
        
        .reveal .diagnosis-result {{
            font-size: 1.3em;
            font-weight: 700;
            margin-bottom: 0.5em;
        }}
        
        .reveal .diagnosis-result.compliant {{
            color: #28a745;
        }}
        
        .reveal .diagnosis-result.non-compliant {{
            color: #dc3545;
        }}
        
        .reveal .risk-item {{
            background: var(--base-color-2);
            padding: 0.8em;
            margin: 0.5em 0;
            border-radius: 4px;
            font-size: 0.95em;
            line-height: 1.6;
        }}
        
        .reveal .risk-item strong {{
            color: var(--main-color);
        }}
        
        .reveal .summary-right {{
            display: flex;
            flex-direction: column;
            justify-content: center;
        }}
        
        .reveal .bei-card {{
            background: linear-gradient(135deg, var(--main-color) 0%, var(--accent-color) 100%);
            color: white;
            padding: 2em;
            border-radius: 8px;
            text-align: center;
            margin-bottom: 1.5em;
            box-shadow: 0 4px 12px rgba(57, 117, 119, 0.3);
        }}
        
        .reveal .bei-card-label {{
            font-size: 1.0em;
            opacity: 0.9;
            margin-bottom: 0.5em;
        }}
        
        .reveal .bei-card-value {{
            font-size: 3.5em;
            font-weight: 700;
        }}
        
        .reveal .equipment-summary {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 0.8em;
        }}
        
        .reveal .equipment-item {{
            background: var(--base-color-1);
            padding: 1em;
            border-radius: 4px;
            border-left: 3px solid var(--main-color);
        }}
        
        .reveal .equipment-item-label {{
            font-size: 0.9em;
            color: var(--text-light);
        }}
        
        .reveal .equipment-item-value {{
            font-size: 1.8em;
            font-weight: 700;
            color: var(--main-color);
        }}
        
        /* PAL*スライド */
        .reveal .pal-slide {{
            background: linear-gradient(135deg, white 0%, var(--base-color-1) 100%);
        }}
        
        .reveal .pal-container {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 2em;
            margin-top: 2em;
        }}
        
        .reveal .pal-metric {{
            background: white;
            padding: 2em;
            border-radius: 8px;
            text-align: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        
        .reveal .pal-metric-label {{
            font-size: 1.1em;
            color: var(--text-light);
            margin-bottom: 0.5em;
        }}
        
        .reveal .pal-metric-value {{
            font-size: 3.0em;
            font-weight: 700;
            color: var(--main-color);
            margin-bottom: 0.3em;
        }}
        
        .reveal .pal-metric-unit {{
            font-size: 0.9em;
            color: var(--text-light);
        }}
        
        .reveal .pal-description {{
            grid-column: 1 / -1;
            background: var(--base-color-1);
            padding: 1.5em;
            border-radius: 4px;
            border-left: 4px solid var(--main-color);
        }}
        
        /* ZEBロードマップスライド */
        .reveal .roadmap-slide {{
            background: linear-gradient(to bottom, var(--base-color-1) 0%, white 100%);
        }}
        
        .reveal .roadmap-container {{
            margin-top: 2em;
        }}
        
        .reveal .roadmap-steps {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 1.5em;
            margin-bottom: 2em;
        }}
        
        .reveal .roadmap-step {{
            position: relative;
            background: white;
            padding: 1.5em;
            border-radius: 8px;
            border-top: 4px solid var(--main-color);
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            text-align: center;
        }}
        
        .reveal .roadmap-step::after {{
            content: '';
            position: absolute;
            right: -1.2em;
            top: 50%;
            transform: translateY(-50%);
            width: 1.5em;
            height: 2px;
            background: var(--main-color);
        }}
        
        .reveal .roadmap-step:last-child::after {{
            display: none;
        }}
        
        .reveal .roadmap-step-number {{
            display: inline-block;
            width: 2.5em;
            height: 2.5em;
            line-height: 2.5em;
            background: var(--main-color);
            color: white;
            border-radius: 50%;
            font-weight: 700;
            font-size: 1.2em;
            margin-bottom: 0.8em;
        }}
        
        .reveal .roadmap-step-title {{
            font-size: 1.1em;
            font-weight: 700;
            color: var(--text-dark);
            margin-bottom: 0.5em;
        }}
        
        .reveal .roadmap-step-desc {{
            font-size: 0.9em;
            color: var(--text-light);
            line-height: 1.5;
        }}
        
        .reveal .roadmap-summary {{
            background: linear-gradient(135deg, var(--main-color) 0%, var(--accent-color) 100%);
            color: white;
            padding: 1.5em;
            border-radius: 8px;
            margin-top: 1.5em;
        }}
        
        .reveal .roadmap-summary-title {{
            font-size: 1.2em;
            font-weight: 700;
            margin-bottom: 0.8em;
        }}
        
        .reveal .roadmap-summary-content {{
            font-size: 0.95em;
            line-height: 1.8;
        }}
        
        .reveal .roadmap-summary-highlight {{
            background: rgba(255,255,255,0.2);
            padding: 0.5em 1em;
            border-radius: 4px;
            margin: 0.5em 0;
        }}
        
        /* 共通スタイル */
        .reveal .chart-container {{
            text-align: center;
            margin: 1.5em 0;
        }}
        
        .reveal .chart-container img {{
            max-width: 100%;
            max-height: 500px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        
        .reveal table {{
            font-size: 0.85em;
            margin: 1.5em auto;
            border-collapse: collapse;
        }}
        
        .reveal table th {{
            background-color: var(--main-color);
            color: white;
            padding: 1em;
            font-weight: 700;
        }}
        
        .reveal table td {{
            padding: 0.8em;
            border-bottom: 1px solid var(--base-color-2);
        }}
        
        .reveal table tr:hover {{
            background-color: var(--base-color-1);
        }}
        
        .reveal .info-box {{
            background: var(--base-color-1);
            border-left: 4px solid var(--main-color);
            padding: 1.2em;
            margin: 1.5em 0;
            border-radius: 4px;
        }}
        
        .reveal .highlight {{
            color: var(--main-color);
            font-weight: 700;
        }}
        
        .reveal .footer-text {{
            position: absolute;
            bottom: 1em;
            right: 1em;
            font-size: 0.7em;
            color: var(--text-light);
        }}
        
        @media print {{
            .reveal .slide-background {{
                background-color: white !important;
            }}
        }}
    </style>
</head>
<body>
    <div class="reveal">
        <div class="slides">
            
            <!-- スライド1: タイトル -->
            <section class="title-slide">
                <div class="title-slide">
                    <h1>one building</h1>
                    <div class="building-name">{building_name}</div>
                    <div class="subtitle">技術レポート</div>
                    <div class="date">{current_date}</div>
                    <div class="footer">
                        <div style="font-size: 0.9em; margin-bottom: 0.5em;">BIM sustaina for Energy</div>
                        <div>copyright © one building</div>
                    </div>
                </div>
            </section>
            
            <!-- スライド2: 総合評価サマリー -->
            <section class="summary-slide">
                <h2>1. 総合評価サマリー: 現状と経営リスク</h2>
                <div class="summary-content">
                    <div class="summary-left">
                        <h3>建物概要</h3>
                        <div style="font-size: 0.95em; line-height: 1.8; color: var(--text-light);">
                            <div><strong>建物名称:</strong> {building_name}</div>
                            <div><strong>所在地:</strong> {location}</div>
                            <div><strong>延べ面積:</strong> {total_area} m²</div>
                        </div>
                        
                        <div class="diagnosis-box {diagnosis_class}">
                            <div class="diagnosis-result {diagnosis_class}">診断結果: {diagnosis_result}</div>
                            <div style="font-size: 0.9em; color: var(--text-light);">
                                2024年4月施行の改正省エネ法に基づく評価
                            </div>
                        </div>
                        
                        <h3 style="margin-top: 1.5em;">▲重要: 経営リスクの特定</h3>
                        <div class="risk-item">
                            <strong>●法的リスク:</strong> BEI 1.0以下に対応していない場合、適合判定（省エネ適判）をパスできず、建築確認申請が受理されません。
                        </div>
                        <div class="risk-item">
                            <strong>●事業リスク:</strong> 新築時の適合義務化により、改修・建替え時の追加コストが発生します。
                        </div>
                    </div>
                    
                    <div class="summary-right">
                        <div class="bei-card">
                            <div class="bei-card-label">現在の {bei_label}</div>
                            <div class="bei-card-value">{bei_total}</div>
                        </div>
                        
                        <h3 style="margin-top: 1em; margin-bottom: 0.8em;">設備別{bei_label}サマリー</h3>
                        <div class="equipment-summary">
                            <div class="equipment-item">
                                <div class="equipment-item-label">空調</div>
                                <div class="equipment-item-value">{bei_ac}</div>
                            </div>
                            <div class="equipment-item">
                                <div class="equipment-item-label">換気</div>
                                <div class="equipment-item-value">{bei_v}</div>
                            </div>
                            <div class="equipment-item">
                                <div class="equipment-item-label">照明</div>
                                <div class="equipment-item-value">{bei_l}</div>
                            </div>
                            <div class="equipment-item">
                                <div class="equipment-item-label">給湯</div>
                                <div class="equipment-item-value">{bei_hw}</div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="footer-text">© 2026 one building</div>
            </section>
            
            <!-- スライド3: エネルギー消費性能の詳細分析 -->
            <section>
                <h2>エネルギー消費性能の詳細分析</h2>
                <div class="chart-container">
                    <img src="{chart_images.get('bei_chart', '')}" alt="BEI比較グラフ">
                </div>
                <div class="footer-text">© 2026 one building</div>
            </section>
            
            <!-- スライド4: 設備別一次エネルギー消費量 -->
            <section>
                <h2>設備別一次エネルギー消費量の比較</h2>
                <div class="chart-container">
                    <img src="{chart_images.get('energy_chart', '')}" alt="エネルギー消費量グラフ">
                </div>
                <div class="footer-text">© 2026 one building</div>
            </section>
            
            <!-- スライド5: PAL*評価 -->
            <section class="pal-slide">
                <h2>2. 外皮性能評価とPAL*ワースト要因の詳細分析</h2>
                <div class="pal-container">
                    <div class="pal-metric">
                        <div class="pal-metric-label">外皮性能指標</div>
                        <div class="pal-metric-value">{pal}</div>
                        <div class="pal-metric-unit">MJ/m²・年</div>
                    </div>
                    <div class="pal-metric">
                        <div class="pal-metric-label">{bpi_label}</div>
                        <div class="pal-metric-value">{bpi}</div>
                        <div class="pal-metric-unit">（基準 0.80 に対して）</div>
                    </div>
                    <div class="pal-description">
                        <strong>評価:</strong> 外皮性能は全体的に良好です。設備側の改善に注力することで、BEI目標達成が可能です。
                    </div>
                </div>
                <div class="footer-text">© 2026 one building</div>
            </section>
            
            <!-- スライド6: 用途別エネルギー消費傾向 -->
            <section>
                <h2>用途別エネルギー消費傾向: BEI分析</h2>
                <div class="chart-container">
                    <img src="{chart_images.get('pie_chart', '')}" alt="エネルギー構成比">
                </div>
                <div class="footer-text">© 2026 one building</div>
            </section>
            
            <!-- スライド7: ZEB化への改善ロードマップ -->
            <section class="roadmap-slide">
                <h2>3. ZEB化への改善ロードマップ</h2>
                <div class="roadmap-container">
                    <div class="roadmap-steps">
                        <div class="roadmap-step">
                            <div class="roadmap-step-number">1</div>
                            <div class="roadmap-step-title">外皮性能の維持</div>
                            <div class="roadmap-step-desc">現状の良好な性能を維持</div>
                        </div>
                        <div class="roadmap-step">
                            <div class="roadmap-step-number">2</div>
                            <div class="roadmap-step-title">空調熱源の高効率化</div>
                            <div class="roadmap-step-desc">最大消費源の本質的改善</div>
                        </div>
                        <div class="roadmap-step">
                            <div class="roadmap-step-number">3</div>
                            <div class="roadmap-step-title">制御の徹底強化</div>
                            <div class="roadmap-step-desc">センサー連動制御の全域導入</div>
                        </div>
                        <div class="roadmap-step">
                            <div class="roadmap-step-number">4</div>
                            <div class="roadmap-step-title">創エネルギーの導入</div>
                            <div class="roadmap-step-desc">太陽光発電等の再エネ設備</div>
                        </div>
                    </div>
                    
                    <div class="roadmap-summary">
                        <div class="roadmap-summary-title">期待される総合効果</div>
                        <div class="roadmap-summary-content">
                            現在のBEI <span class="roadmap-summary-highlight" style="font-weight: 700;">{bei_total}</span> から ZEB Oriented 要件 <span class="roadmap-summary-highlight" style="font-weight: 700;">0.70</span> まで、約 <span class="roadmap-summary-highlight" style="font-weight: 700;">0.98</span> の削減が必要です。上記4ステップの実施により、段階的にBEI目標達成を目指します。
                        </div>
                    </div>
                </div>
                <div class="footer-text">© 2026 one building</div>
            </section>
            
            <!-- 最終スライド -->
            <section class="title-slide">
                <h1>ありがとうございました</h1>
                <div class="subtitle" style="margin-top: 2em;">
                    本レポートは one building 技術レポート生成ツールで自動生成されました
                </div>
                <div class="footer">
                    <div style="font-size: 0.9em; margin-bottom: 0.5em;">BIM sustaina for Energy</div>
                    <div>copyright © one building</div>
                </div>
            </section>
            
        </div>
    </div>
    
    <script src="https://cdn.jsdelivr.net/npm/reveal.js@4.6.0/dist/reveal.js"></script>
    <script>
        Reveal.initialize({{
            hash: true,
            center: false,
            transition: 'slide',
            width: 1280,
            height: 720,
            margin: 0.1,
            controls: true,
            progress: true,
            slideNumber: 'c/t',
            keyboard: true,
            overview: true,
            touch: true,
            loop: false,
            rtl: false,
            navigationMode: 'default',
            shuffle: false,
            fragments: true,
            fragmentInURL: true,
            embedded: false,
            help: true,
            pause: true,
            showNotes: false,
            autoPlayMedia: null,
            preloadIframes: null,
            autoAnimate: true,
            autoAnimateDuration: 1.0,
            autoAnimateUnmatched: true,
            autoSlide: 0,
            autoSlideStoppable: true,
            autoSlideMethod: null,
            defaultTiming: null,
            mouseWheel: false,
            previewLinks: false,
            postMessage: true,
            postMessageEvents: false,
            focusBodyOnPageVisibilityChange: true,
            viewDistance: 3,
            mobileViewDistance: 2,
            display: 'block',
            hideInactiveCursor: true,
            hideCursorTime: 5000
        }});
    </script>
</body>
</html>
"""
    
    return html
