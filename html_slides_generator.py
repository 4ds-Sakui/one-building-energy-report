'''
HTMLスライド生成モジュール (v1.2)
モデル建物法と標準入力法の自動切り替え対応
'''

import base64
from io import BytesIO
import matplotlib.pyplot as plt
from datetime import datetime

COLOR_MAIN_HEX = "#397577"
COLOR_RED_HEX = "#E63946"
COLOR_GREEN_HEX = "#2A9D8F"

def generate_html_slides(data, charts):
    chart_images = {}
    for name, obj in charts.items():
        if obj:
            if isinstance(obj, BytesIO):
                obj.seek(0)
                img_base64 = base64.b64encode(obj.read()).decode('utf-8')
                chart_images[name] = f"data:image/png;base64,{img_base64}"
            else:
                buf = BytesIO()
                obj.savefig(buf, format='png', dpi=120, bbox_inches='tight')
                buf.seek(0)
                img_base64 = base64.b64encode(buf.read()).decode('utf-8')
                chart_images[name] = f"data:image/png;base64,{img_base64}"

    calc_method = data.get('calculation_method', 'standard_input')
    is_model = (calc_method == 'model_building')
    bei_label = "BEIm" if is_model else "BEI"
    bpi_label = "BPIm" if is_model else "BPI"
    
    bei_total = data.get('bei_total', 1.0)
    bpi_val = data.get('bpi', 1.0)
    
    # 判定ロジック
    is_compliant = (bei_total <= 1.0)
    status_color = COLOR_GREEN_HEX if is_compliant else COLOR_RED_HEX
    status_text = f"診断結果: {('基準適合' if is_compliant else '基準非適合')}"
    
    # モデル建物法専用の解説
    method_desc = "モデル建物法による簡易評価" if is_model else "標準入力法による詳細評価"
    
    html = f'''<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>省エネ診断レポート</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.0/dist/reset.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.0/dist/reveal.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.0/dist/theme/white.css">
    <style>
        :root {{ --main: {COLOR_MAIN_HEX}; --red: {COLOR_RED_HEX}; --green: {COLOR_GREEN_HEX}; }}
        .reveal {{ font-family: "Noto Sans JP", sans-serif; }}
        section {{ text-align: left; }}
        .title-slide {{ text-align: center !important; }}
        .header {{ border-bottom: 3px solid var(--main); margin-bottom: 20px; font-weight: bold; font-size: 1.2em; }}
        .card {{ background: #f8f9fa; padding: 20px; border-radius: 10px; margin-bottom: 20px; }}
        .val-big {{ font-size: 3em; font-weight: bold; color: var(--main); }}
        .grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
        .footer {{ position: absolute; bottom: 20px; width: 100%; text-align: center; font-size: 0.5em; color: #999; }}
    </style>
</head>
<body>
    <div class="reveal">
        <div class="slides">
            <section class="title-slide">
                <h1 style="color: var(--main)">技術レポート</h1>
                <h3>{data.get('building_name')}</h3>
                <p>{method_desc}</p>
                <div class="footer">© one building | {datetime.now().strftime('%Y/%m/%d')}</div>
            </section>

            <section>
                <div class="header">1. 総合判定サマリー</div>
                <div class="grid">
                    <div class="card">
                        <div style="font-size: 0.8em">全体 {bei_label}</div>
                        <div class="val-big" style="color: {status_color}">{bei_total:.2f}</div>
                        <div style="font-weight: bold; color: {status_color}">{status_text}</div>
                    </div>
                    <div class="card">
                        <div style="font-size: 0.8em">外皮性能 {bpi_label}</div>
                        <div class="val-big">{bpi_val:.2f}</div>
                        <div style="font-weight: bold; color: {COLOR_GREEN_HEX if bpi_val <= 1.0 else COLOR_RED_HEX}">
                            {('適合' if bpi_val <= 1.0 else '非適合')}
                        </div>
                    </div>
                </div>
                <div class="card" style="border-left: 5px solid {status_color}">
                    <strong>経営リスク評価:</strong><br>
                    {('法的リスクなし。建築確認申請をスムーズに進められます。' if is_compliant else '法的リスクあり。改正省エネ法に基づき建築確認が受理されない恐れがあります。')}
                </div>
            </section>

            <section>
                <div class="header">2. 設備別 {bei_label} 分析</div>
                <div style="text-align: center">
                    <img src="{chart_images.get('bei_chart', '')}" style="max-height: 450px;">
                </div>
            </section>

            <section>
                <div class="header">3. 改善ロードマップ</div>
                <div class="grid">
                    <div class="card">
                        <strong>STEP 1: 外皮強化</strong><br>
                        窓のLow-E化や断熱材の見直しにより{bpi_label}を改善。
                    </div>
                    <div class="card">
                        <strong>STEP 2: 設備高効率化</strong><br>
                        空調や照明(LED)の更新により{bei_label}を削減。
                    </div>
                </div>
                <div class="card">
                    <strong>目標:</strong> {bei_label} 0.60以下（ZEB Ready相当）を目指すための改修案を推奨します。
                </div>
            </section>
        </div>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/reveal.js@4.6.0/dist/reveal.js"></script>
    <script>
        Reveal.initialize({{ width: 1280, height: 720, margin: 0.1, center: false, hash: true }});
    </script>
</body>
</html>'''
    return html
