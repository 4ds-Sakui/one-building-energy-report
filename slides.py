#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
スライド作成モジュール (Streamlit対応版 v1.2)
モデル建物法と標準入力法の自動切り替え対応
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from datetime import datetime
import io
import tempfile
import os

# report_generator.pyからカラー定義をインポート
from report_generator import (
    COLOR_MAIN, COLOR_RED, COLOR_GREEN, COLOR_GRAY,
    COLOR_WHITE, COLOR_BLACK,
    get_bei_label, get_bpi_label,
    generate_improvement_roadmap
)

def create_presentation(data, chart_stacked_bytes, chart_pie_bytes, chart_bei_bytes):
    """
    PowerPointプレゼンテーションを作成してBytesIOで返す
    """
    prs = Presentation()
    
    # 16:9ワイドスクリーン
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    
    add_title_slide_tech_report_style(prs, data)
    add_summary_slide(prs, data)
    
    calc_method = data.get('calculation_method', 'standard_input')
    is_model = (calc_method == 'model_building')
    
    # 標準入力法の場合のみ、詳細グラフを追加
    if not is_model:
        # 一時ファイルにグラフを保存してスライドに追加
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_stacked:
            tmp_stacked.write(chart_stacked_bytes.read())
            tmp_stacked_path = tmp_stacked.name
        chart_stacked_bytes.seek(0)
        
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_pie:
            tmp_pie.write(chart_pie_bytes.read())
            tmp_pie_path = tmp_pie.name
        chart_pie_bytes.seek(0)
        
        slide3 = prs.slides.add_slide(prs.slide_layouts[6])
        add_slide_title(slide3, "エネルギー消費性能の詳細分析")
        slide3.shapes.add_picture(tmp_stacked_path, Inches(0.3), Inches(0.95), width=Inches(9.4))
        
        slide4 = prs.slides.add_slide(prs.slide_layouts[6])
        add_slide_title(slide4, "設備別一次エネルギー消費量の比較")
        slide4.shapes.add_picture(tmp_pie_path, Inches(0.3), Inches(1.1), width=Inches(9.4))
        
        os.unlink(tmp_stacked_path)
        os.unlink(tmp_pie_path)
    
    add_envelope_worst_analysis_slide(prs, data)
    
    # BEI比較グラフを一時ファイルに保存
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_bei:
        tmp_bei.write(chart_bei_bytes.read())
        tmp_bei_path = tmp_bei.name
    chart_bei_bytes.seek(0)
    
    slide6 = prs.slides.add_slide(prs.slide_layouts[6])
    bei_label = get_bei_label(calc_method)
    add_slide_title(slide6, f"用途別エネルギー消費傾向: {bei_label}分析")
    slide6.shapes.add_picture(tmp_bei_path, Inches(0.3), Inches(1.05), width=Inches(9.4))
    
    os.unlink(tmp_bei_path)
    
    add_improvement_roadmap_slide(prs, data)
    
    # BytesIOに保存
    pptx_bytes = io.BytesIO()
    prs.save(pptx_bytes)
    pptx_bytes.seek(0)
    
    return pptx_bytes

def add_slide_title(slide, title_text):
    title_box = slide.shapes.add_textbox(Inches(0.3), Inches(0.3), Inches(9.4), Inches(0.5))
    title_frame = title_box.text_frame
    title_frame.text = title_text
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(22)
    title_para.font.bold = True
    title_para.font.color.rgb = COLOR_MAIN
    title_para.font.name = 'Noto Sans JP'

def add_title_slide_tech_report_style(prs, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    center_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.5), Inches(0.8), Inches(5), Inches(4))
    center_box.fill.solid()
    center_box.fill.fore_color.rgb = COLOR_MAIN
    center_box.line.fill.background()
    
    logo_box = slide.shapes.add_textbox(Inches(3.5), Inches(1.3), Inches(3), Inches(0.4))
    logo_box.text_frame.text = "one building"
    logo_box.text_frame.paragraphs[0].font.size = Pt(24)
    logo_box.text_frame.paragraphs[0].font.color.rgb = COLOR_WHITE
    logo_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    building_box = slide.shapes.add_textbox(Inches(3), Inches(2.0), Inches(4), Inches(0.4))
    building_box.text_frame.text = data['building_name']
    building_box.text_frame.paragraphs[0].font.size = Pt(22)
    building_box.text_frame.paragraphs[0].font.color.rgb = COLOR_WHITE
    building_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    title_box = slide.shapes.add_textbox(Inches(3), Inches(2.6), Inches(4), Inches(0.5))
    title_box.text_frame.text = "技術レポート"
    title_box.text_frame.paragraphs[0].font.size = Pt(28)
    title_box.text_frame.paragraphs[0].font.color.rgb = COLOR_WHITE
    title_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    date_box = slide.shapes.add_textbox(Inches(3.5), Inches(3.5), Inches(3), Inches(0.35))
    date_box.text_frame.text = datetime.now().strftime('%Y.%m.%d')
    date_box.text_frame.paragraphs[0].font.size = Pt(16)
    date_box.text_frame.paragraphs[0].font.color.rgb = COLOR_WHITE
    date_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

def add_summary_slide(prs, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_title(slide, "1. 総合評価サマリー: 現状と経営リスク")
    
    info_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(9), Inches(0.5))
    info_box.text_frame.text = f"建物名称: {data['building_name']}  |  所在地: {data['location']}  |  延べ面積: {data['total_area']} m²"
    info_box.text_frame.paragraphs[0].font.size = Pt(11)
    info_box.text_frame.paragraphs[0].font.color.rgb = COLOR_GRAY
    
    is_compliant = (data['bei_total'] <= 1.0)
    status_color = COLOR_GREEN if is_compliant else COLOR_RED
    status_text = f"診断結果: {('基準適合' if is_compliant else '基準非適合')}"
    
    res_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.6), Inches(4.5), Inches(0.5))
    res_box.text_frame.text = status_text
    res_box.text_frame.paragraphs[0].font.size = Pt(18)
    res_box.text_frame.paragraphs[0].font.bold = True
    res_box.text_frame.paragraphs[0].font.color.rgb = status_color
    
    risk_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(2.2), Inches(9), Inches(2.0))
    risk_shape.fill.solid()
    risk_shape.fill.fore_color.rgb = RGBColor(243, 255, 243) if is_compliant else RGBColor(255, 243, 243)
    risk_shape.line.color.rgb = status_color
    
    tf = risk_shape.text_frame
    p1 = tf.paragraphs[0]
    p1.text = f"▲{('優位性' if is_compliant else '重要')}: 経営影響の特定"
    p1.font.bold = True
    p1.font.color.rgb = status_color
    
    p2 = tf.add_paragraph()
    p2.text = f"● 法的リスク: {('基準適合。建築確認申請が受理されます。' if is_compliant else '基準非適合。建築確認申請が受理されない恐れがあります。')}"
    p2.font.size = Pt(12)
    
    p3 = tf.add_paragraph()
    p3.text = f"● 経済的影響: {('光熱費削減。' if is_compliant else '光熱費高騰。')}"
    p3.font.size = Pt(12)

def add_envelope_worst_analysis_slide(prs, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    calc_method = data.get('calculation_method', 'standard_input')
    bpi_label = get_bpi_label(calc_method)
    add_slide_title(slide, f"2. 外皮性能評価 ({bpi_label})")
    
    box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
    box.text_frame.text = f"{bpi_label} 値: {data['bpi']:.2f}"
    box.text_frame.paragraphs[0].font.size = Pt(44)
    box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    box.text_frame.paragraphs[0].font.color.rgb = COLOR_MAIN

def add_improvement_roadmap_slide(prs, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_title(slide, "4. ZEB化への改善ロードマップ")
    
    roadmap = generate_improvement_roadmap(data)
    for i, step in enumerate(roadmap[:4]):
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5 + i*2.3), Inches(1.5), Inches(2.1), Inches(3))
        box.fill.solid()
        box.fill.fore_color.rgb = COLOR_WHITE
        box.line.color.rgb = COLOR_MAIN
        
        tf = box.text_frame
        tf.text = step['step']
        tf.paragraphs[0].font.bold = True
        p2 = tf.add_paragraph()
        p2.text = step['title']
        p2.font.size = Pt(12)
        p3 = tf.add_paragraph()
        p3.text = step['desc']
        p3.font.size = Pt(10)
