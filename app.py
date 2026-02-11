#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
one building 技術レポート生成ツール
省エネ診断レポートPowerPoint自動生成Webアプリ
"""

import streamlit as st
import io
from datetime import datetime
from report_generator import extract_data_from_file, create_stacked_bar_chart_improved, create_pie_charts, create_bei_comparison_chart_with_total
from slides import create_presentation

# ページ設定
st.set_page_config(
    page_title="one building 技術レポート生成ツール",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# カスタムCSS（one buildingブランドカラー）
st.markdown("""
<style>
    .main-title {
        color: #397577;
        font-size: 2.5rem;
        font-weight: bold;
        margin-bottom: 0.5rem;
    }
    .sub-title {
        color: #6c757d;
        font-size: 1.2rem;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .stButton>button {
        background-color: #397577;
        color: white;
        font-weight: bold;
        border-radius: 5px;
        padding: 0.5rem 2rem;
        border: none;
    }
    .stButton>button:hover {
        background-color: #013E34;
    }
</style>
""", unsafe_allow_html=True)

# タイトル
st.markdown('<div class="main-title">📊 one building 技術レポート生成ツール</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">省エネ診断レポートを自動生成します</div>', unsafe_allow_html=True)

# サイドバー
with st.sidebar:
    st.image("https://via.placeholder.com/200x80/397577/FFFFFF?text=one+building", use_container_width=True)
    st.markdown("---")
    st.markdown("### 📋 使い方")
    st.markdown("""
    **STEP 1**: ファイルをアップロード  
    PDF または テキストファイル（.txt, .md）を選択してください。
    
    **STEP 2**: レポート生成  
    「レポート生成」ボタンをクリックします。
    
    **STEP 3**: ダウンロード  
    生成されたPowerPointファイルをダウンロードします。
    """)
    
    st.markdown("---")
    st.markdown("### ⚠️ 重要な注意事項")
    st.warning("""
    **PDFファイルについて**:  
    PDFからのデータ抽出は、ファイルの形式や構造により正確でない場合があります。
    
    **推奨**: Markdown形式（.txt, .md）のテキストファイルを使用すると、より高精度なデータ抽出が可能です。
    
    省エネ診断プログラムの出力結果をテキスト形式で保存してご利用ください。
    """)
    
    st.markdown("---")
    st.markdown("### ℹ️ 対応形式")
    st.markdown("""
    - **標準入力法**: 詳細なエネルギー消費量分析
    - **モデル建物法**: BEI/BPIm表記での簡易分析
    
    計算方法は自動判定されます。
    """)
    
    st.markdown("---")
    st.markdown("### 🎨 デザイン")
    st.markdown("""
    - **スライド形式**: 16:9 ワイドスクリーン
    - **フォント**: Noto Sans JP
    - **カラー**: one building ブランドカラー
    """)

# メインコンテンツ
col1, col2 = st.columns([2, 1])

with col1:
    st.markdown("### 📁 ファイルアップロード")
    
    # 注意書きを表示
    st.info("""
    💡 **推奨**: より正確なデータ抽出のため、**Markdown/テキスト形式（.txt, .md）** のファイルをご利用ください。  
    PDFファイルは表の構造が崩れる場合があり、データ抽出精度が低下する可能性があります。
    """)
    
    uploaded_file = st.file_uploader(
        "省エネ計算結果のPDFまたはテキストファイルを選択してください",
        type=['pdf', 'txt', 'md'],
        help="対応形式: PDF, テキスト (.txt), Markdown (.md) | 推奨: Markdown/テキスト形式"
    )
    
    if uploaded_file is not None:
        st.markdown('<div class="info-box">', unsafe_allow_html=True)
        st.write(f"**ファイル名**: {uploaded_file.name}")
        st.write(f"**ファイルサイズ**: {uploaded_file.size / 1024:.2f} KB")
        st.write(f"**ファイル形式**: {uploaded_file.type}")
        st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown("### ⚙️ 生成オプション")
    st.info("現在、すべてのオプションは自動設定されます")

# レポート生成ボタン
if uploaded_file is not None:
    st.markdown("---")
    
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        generate_button = st.button("🚀 レポート生成", use_container_width=True)
    
    if generate_button:
        try:
            with st.spinner('📊 データを抽出しています...'):
                # ファイルポインタを先頭に戻す
                uploaded_file.seek(0)
                
                # データ抽出
                data = extract_data_from_file(uploaded_file, uploaded_file.name)
                
                st.success(f"✅ データ抽出完了: {data['building_name']}")
            
            # 抽出データの表示
            with st.expander("📋 抽出されたデータを確認"):
                col_data1, col_data2, col_data3 = st.columns(3)
                
                with col_data1:
                    st.metric("建物名称", data['building_name'])
                    st.metric("所在地", data['location'])
                    st.metric("延べ面積", f"{data['total_area']} m²")
                
                with col_data2:
                    bei_label = "BEIm" if data.get('calculation_method') == 'model_building' else "BEI"
                    st.metric(f"全体{bei_label}", f"{data['bei_total']:.2f}")
                    st.metric(f"空調{bei_label}", f"{data['bei_ac']:.2f}")
                    st.metric(f"換気{bei_label}", f"{data['bei_v']:.2f}")
                
                with col_data3:
                    st.metric(f"照明{bei_label}", f"{data['bei_l']:.2f}")
                    st.metric(f"給湯{bei_label}", f"{data['bei_hw']:.2f}")
                    bpi_label = "BPIm" if data.get('calculation_method') == 'model_building' else "BPI"
                    st.metric(bpi_label, f"{data['bpi']:.2f}")
            
            # グラフ生成
            with st.spinner('📈 グラフを生成しています...'):
                chart_bei_bytes = create_bei_comparison_chart_with_total(data)
                
                calc_method = data.get('calculation_method', 'standard_input')
                if calc_method == 'standard_input':
                    chart_stacked_bytes = create_stacked_bar_chart_improved(data)
                    chart_pie_bytes = create_pie_charts(data)
                else:
                    # モデル建物法の場合はダミーのBytesIO
                    chart_stacked_bytes = io.BytesIO()
                    chart_pie_bytes = io.BytesIO()
                
                st.success("✅ グラフ生成完了")
            
            # PowerPoint生成
            with st.spinner('📄 PowerPointレポートを生成しています...'):
                pptx_bytes = create_presentation(
                    data,
                    chart_stacked_bytes,
                    chart_pie_bytes,
                    chart_bei_bytes
                )
                
                st.success("✅ PowerPointレポート生成完了")
            
            # ダウンロードボタン
            st.markdown("---")
            st.markdown("### 📥 ダウンロード")
            
            # ファイル名を生成
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Energy_Diagnosis_Report_{data['building_name']}_{timestamp}.pptx"
            
            col_dl1, col_dl2, col_dl3 = st.columns([1, 2, 1])
            with col_dl2:
                st.download_button(
                    label="💾 PowerPointをダウンロード",
                    data=pptx_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
            
            st.markdown('<div class="success-box">', unsafe_allow_html=True)
            st.markdown(f"""
            **✨ レポート生成が完了しました！**
            
            - **建物名**: {data['building_name']}
            - **計算方法**: {'モデル建物法' if data.get('calculation_method') == 'model_building' else '標準入力法'}
            - **スライド数**: {'5枚' if data.get('calculation_method') == 'model_building' else '7枚'}
            - **ファイル名**: {filename}
            
            上のボタンからダウンロードしてください。
            """)
            st.markdown('</div>', unsafe_allow_html=True)
            
        except Exception as e:
            st.error(f"❌ エラーが発生しました: {str(e)}")
            st.exception(e)

else:
    # ファイル未選択時の案内
    st.info("👆 まず、省エネ計算結果のファイルをアップロードしてください。")
    
    st.markdown("---")
    st.markdown("### 📖 サンプルデータ")
    st.markdown("""
    以下のような省エネ計算結果ファイルに対応しています:
    
    - **標準入力法**: 詳細なエネルギー消費量データを含むPDF/テキスト
    - **モデル建物法**: BEI/BPIm値を含むPDF
    
    ファイルをアップロードすると、自動的に計算方法を判定し、適切なレポートを生成します。
    """)

# フッター
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #6c757d; font-size: 0.9rem;">
    <p>© 2026 one building | BIM sustaina for Energy</p>
    <p>技術レポート自動生成ツール v1.0</p>
</div>
""", unsafe_allow_html=True)
