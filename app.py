import streamlit as st
import os
from report_generator import extract_data_from_markdown
from html_slides_generator import generate_html_slides

st.set_page_config(page_title="one building - 技術レポート生成", layout="wide")

st.title("one building 技術レポート生成 (v1.4)")
st.markdown("""
Markdown形式の省エネ診断結果をアップロードしてください。
モデル建物法の詳細分析と、標準入力法へのアップグレード提案を含むHTMLレポートを生成します。
""")

uploaded_file = st.file_uploader("Markdownファイルをアップロード (.md, .txt)", type=["md", "txt"])

if uploaded_file:
    content = uploaded_file.read().decode("utf-8")
    
    with st.spinner("データを解析中..."):
        data = extract_data_from_markdown(content)
        
        st.success(f"解析完了: {data['building_name']}")
        
        col1, col2, col3 = st.columns(3)
        col1.metric("BEIm / BEI", f"{data['bei_total']:.2f}")
        col2.metric("BPIm / BPI", f"{data['bpi']:.2f}")
        col3.metric("床面積", f"{data['total_area']:,} m²")

        # HTMLレポート生成
        html_report = generate_html_slides(data)
        
        st.subheader("レポート出力")
        st.download_button(
            label="HTMLレポートをダウンロード",
            data=html_report,
            file_name=f"Technical_Report_{data['building_name']}.html",
            mime="text/html"
        )
        
        st.info("ダウンロードしたHTMLファイルをブラウザで開くと、プレゼンテーション形式で閲覧できます。")

st.divider()
st.caption("© 2026 one building. 全ての権利を保有。")
