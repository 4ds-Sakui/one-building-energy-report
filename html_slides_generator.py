"""
HTMLスライド生成モジュール
reveal.jsを使用してインタラクティブなスライド形式のレポートを生成
"""

import base64
from io import BytesIO
from datetime import datetime
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt


def generate_html_slides(data, charts):
    """
    reveal.jsを使用したHTMLスライドを生成
    
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
    pal = data.get('pal', 'N/A')
    
    # エネルギー消費量
    energy_data = data.get('energy_consumption', {})
    
    # 現在日時
    current_date = datetime.now().strftime('%Y年%m月%d日')
    
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
    <style>
        :root {{
            --one-building-color: #397577;
            --one-building-light: #5a9799;
        }}
        
        .reveal {{
            font-family: 'Noto Sans JP', 'Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'メイリオ', Meiryo, sans-serif;
        }}
        
        .reveal h1, .reveal h2, .reveal h3 {{
            color: var(--one-building-color);
            font-weight: 700;
        }}
        
        .reveal h1 {{
            font-size: 2.5em;
            margin-bottom: 0.5em;
        }}
        
        .reveal h2 {{
            font-size: 1.8em;
            margin-bottom: 0.8em;
            border-bottom: 3px solid var(--one-building-color);
            padding-bottom: 0.3em;
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
        }}
        
        .reveal .subtitle {{
            font-size: 1.2em;
            color: #666;
            margin-top: 1em;
        }}
        
        .reveal .date {{
            font-size: 0.9em;
            color: #999;
            margin-top: 2em;
        }}
        
        .reveal .info-grid {{
            display: grid;
            grid-template-columns: 150px 1fr;
            gap: 0.8em;
            margin: 1.5em 0;
            font-size: 0.9em;
        }}
        
        .reveal .info-label {{
            font-weight: bold;
            color: var(--one-building-color);
        }}
        
        .reveal .info-value {{
            color: #333;
        }}
        
        .reveal .bei-grid {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 1em;
            margin: 1.5em 0;
        }}
        
        .reveal .bei-card {{
            background: linear-gradient(135deg, var(--one-building-light) 0%, var(--one-building-color) 100%);
            color: white;
            padding: 1.5em;
            border-radius: 10px;
            text-align: center;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        
        .reveal .bei-card-label {{
            font-size: 0.9em;
            opacity: 0.9;
            margin-bottom: 0.5em;
        }}
        
        .reveal .bei-card-value {{
            font-size: 2.5em;
            font-weight: bold;
        }}
        
        .reveal .chart-container {{
            text-align: center;
            margin: 1em 0;
        }}
        
        .reveal .chart-container img {{
            max-width: 100%;
            max-height: 500px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        
        .reveal .footer {{
            position: absolute;
            bottom: 1em;
            left: 1em;
            font-size: 0.6em;
            color: #999;
        }}
        
        .reveal .highlight {{
            color: var(--one-building-color);
            font-weight: bold;
        }}
        
        .reveal table {{
            font-size: 0.8em;
            margin: 1em auto;
        }}
        
        .reveal table th {{
            background-color: var(--one-building-color);
            color: white;
            padding: 0.8em;
        }}
        
        .reveal table td {{
            padding: 0.6em;
            border-bottom: 1px solid #ddd;
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
                <h1>📊 省エネ診断レポート</h1>
                <div class="subtitle">{building_name}</div>
                <div class="date">{current_date}</div>
                <div class="footer">© 2026 one building | BIM sustaina for Energy</div>
            </section>
            
            <!-- スライド2: 建物情報 -->
            <section>
                <h2>🏢 建物情報</h2>
                <div class="info-grid">
                    <div class="info-label">建物名称</div>
                    <div class="info-value">{building_name}</div>
                    
                    <div class="info-label">建物用途</div>
                    <div class="info-value">{building_type}</div>
                    
                    <div class="info-label">計算方法</div>
                    <div class="info-value highlight">{calculation_method}</div>
                </div>
            </section>
            
            <!-- スライド3: BEI総合評価 -->
            <section>
                <h2>📈 {bei_label} 総合評価</h2>
                <div class="bei-grid">
                    <div class="bei-card">
                        <div class="bei-card-label">全体{bei_label}</div>
                        <div class="bei-card-value">{bei_total}</div>
                    </div>
                    <div class="bei-card">
                        <div class="bei-card-label">{bpi_label}</div>
                        <div class="bei-card-value">{bpi}</div>
                    </div>
                    <div class="bei-card">
                        <div class="bei-card-label">PAL*</div>
                        <div class="bei-card-value">{pal}</div>
                    </div>
                </div>
            </section>
            
            <!-- スライド4: BEI用途別 -->
            <section>
                <h2>🔍 {bei_label} 用途別詳細</h2>
                <table>
                    <thead>
                        <tr>
                            <th>用途</th>
                            <th>{bei_label}値</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>空調（{bei_label}/AC）</td>
                            <td>{bei_ac}</td>
                        </tr>
                        <tr>
                            <td>換気（{bei_label}/V）</td>
                            <td>{bei_v}</td>
                        </tr>
                        <tr>
                            <td>照明（{bei_label}/L）</td>
                            <td>{bei_l}</td>
                        </tr>
                        <tr>
                            <td>給湯（{bei_label}/HW）</td>
                            <td>{bei_hw}</td>
                        </tr>
                        <tr>
                            <td>昇降機（{bei_label}/EV）</td>
                            <td>{bei_ev}</td>
                        </tr>
                    </tbody>
                </table>
            </section>
"""
    
    # スライド5: BEI比較グラフ
    if 'bei_chart' in chart_images:
        html += f"""
            <!-- スライド5: BEI比較グラフ -->
            <section>
                <h2>📊 {bei_label} 比較グラフ</h2>
                <div class="chart-container">
                    <img src="{chart_images['bei_chart']}" alt="BEI比較グラフ">
                </div>
            </section>
"""
    
    # スライド6: エネルギー消費量グラフ
    if 'energy_chart' in chart_images:
        html += f"""
            <!-- スライド6: エネルギー消費量 -->
            <section>
                <h2>⚡ エネルギー消費量</h2>
                <div class="chart-container">
                    <img src="{chart_images['energy_chart']}" alt="エネルギー消費量グラフ">
                </div>
            </section>
"""
    
    # スライド7: エネルギー構成比
    if 'pie_chart' in chart_images:
        html += f"""
            <!-- スライド7: エネルギー構成比 -->
            <section>
                <h2>🥧 エネルギー構成比</h2>
                <div class="chart-container">
                    <img src="{chart_images['pie_chart']}" alt="エネルギー構成比">
                </div>
            </section>
"""
    
    # 最終スライド
    html += """
            <!-- 最終スライド -->
            <section class="title-slide">
                <h2>ありがとうございました</h2>
                <p style="text-align: center; margin-top: 2em; color: #666;">
                    本レポートは one building 技術レポート生成ツールで自動生成されました
                </p>
                <div class="footer">© 2026 one building | BIM sustaina for Energy</div>
            </section>
            
        </div>
    </div>
    
    <script src="https://cdn.jsdelivr.net/npm/reveal.js@4.6.0/dist/reveal.js"></script>
    <script>
        Reveal.initialize({
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
            autoAnimateMatcher: null,
            autoAnimateEasing: 'ease',
            autoAnimateDuration: 1.0,
            autoAnimateUnmatched: true,
            autoAnimateStyles: [
                'opacity',
                'color',
                'background-color',
                'padding',
                'font-size',
                'line-height',
                'letter-spacing',
                'border-width',
                'border-color',
                'border-radius',
                'outline',
                'outline-offset'
            ],
            autoSlide: 0,
            autoSlideStoppable: true,
            autoSlideMethod: null,
            defaultTiming: null,
            mouseWheel: false,
            previewLinks: false,
            postMessage: true,
            postMessageEvents: false,
            focusBodyOnPageVisibilityChange: true,
            transition: 'slide',
            transitionSpeed: 'default',
            backgroundTransition: 'fade',
            viewDistance: 3,
            mobileViewDistance: 2,
            display: 'block',
            hideInactiveCursor: true,
            hideCursorTime: 5000
        });
    </script>
</body>
</html>
"""
    
    return html
