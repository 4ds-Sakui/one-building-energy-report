#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
スライド作成モジュール (Streamlit対応版)
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from datetime import datetime
import io
import tempfile

# report_generator.pyからカラー定義をインポート
from report_generator import (
    COLOR_MAIN, COLOR_ACCENT, COLOR_BASE1, COLOR_BASE2,
    COLOR_RED, COLOR_GREEN, COLOR_BLUE, COLOR_GRAY,
    COLOR_LIGHT_GRAY, COLOR_WHITE, COLOR_BLACK,
    GRAPH_COLORS, get_bei_label, get_bpi_label,
    generate_improvement_roadmap
)


def create_presentation(data, chart_stacked_bytes, chart_pie_bytes, chart_bei_bytes):
    """
    PowerPointプレゼンテーションを作成してBytesIOで返す（Streamlit対応版）
    
    Args:
        data: 抽出されたエネルギーデータ
        chart_stacked_bytes: 積み上げ棒グラフのBytesIO
        chart_pie_bytes: パイチャートのBytesIO
        chart_bei_bytes: BEI比較グラフのBytesIO
    
    Returns:
        BytesIO: PPTXファイルのバイトストリーム
    """
    prs = Presentation()
    
    # 16:9ワイドスクリーン
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    
    add_title_slide_tech_report_style(prs, data)
    add_summary_slide(prs, data)
    
    # モデル建物法の場合、エネルギー消費量詳細は取得不可のためスキップ
    calc_method = data.get('calculation_method', 'standard_input')
    if calc_method == 'standard_input':
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
        
        # 一時ファイルを削除
        import os
        os.unlink(tmp_stacked_path)
        os.unlink(tmp_pie_path)
    
    add_envelope_worst_analysis_slide(prs, data)
    
    # BEI比較グラフを一時ファイルに保存
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_bei:
        tmp_bei.write(chart_bei_bytes.read())
        tmp_bei_path = tmp_bei.name
    chart_bei_bytes.seek(0)
    
    slide6 = prs.slides.add_slide(prs.slide_layouts[6])
    bei_label = get_bei_label(data.get('calculation_method', 'standard_input'))
    add_slide_title(slide6, f"用途別エネルギー消費傾向: {bei_label}分析")
    slide6.shapes.add_picture(tmp_bei_path, Inches(0.3), Inches(1.05), width=Inches(9.4))
    
    import os
    os.unlink(tmp_bei_path)
    
    add_improvement_roadmap_slide(prs, data)
    
    # BytesIOに保存
    pptx_bytes = io.BytesIO()
    prs.save(pptx_bytes)
    pptx_bytes.seek(0)
    
    return pptx_bytes

def add_slide_title(slide, title_text):
    """スライドにタイトルを追加（Noto Sans JP）"""
    title_box = slide.shapes.add_textbox(Inches(0.3), Inches(0.3), Inches(9.4), Inches(0.5))
    title_frame = title_box.text_frame
    title_frame.text = title_text
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(22)
    title_para.font.bold = True
    title_para.font.color.rgb = COLOR_MAIN
    title_para.font.name = 'Noto Sans JP'

def add_title_slide_tech_report_style(prs, data):
    """技術レポート風タイトルスライドを作成（フォントサイズ調整）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_WHITE
    
    # 中央の大きなティールグリーンの矩形
    center_box = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(2.5), Inches(0.8), Inches(5), Inches(4)
    )
    center_box.fill.solid()
    center_box.fill.fore_color.rgb = COLOR_MAIN
    center_box.line.fill.background()
    
    # one buildingロゴ風テキスト（サイズ縮小）
    logo_box = slide.shapes.add_textbox(Inches(3.5), Inches(1.3), Inches(3), Inches(0.4))
    logo_frame = logo_box.text_frame
    logo_frame.text = "one building"
    logo_para = logo_frame.paragraphs[0]
    logo_para.font.size = Pt(24)
    logo_para.font.bold = True
    logo_para.font.color.rgb = COLOR_WHITE
    logo_para.font.name = 'Noto Sans JP'
    logo_para.alignment = PP_ALIGN.CENTER
    
    # 建物名称（サイズ縮小）
    building_box = slide.shapes.add_textbox(Inches(3), Inches(2.0), Inches(4), Inches(0.4))
    building_frame = building_box.text_frame
    building_frame.text = data['building_name']
    building_para = building_frame.paragraphs[0]
    building_para.font.size = Pt(22)
    building_para.font.bold = True
    building_para.font.color.rgb = COLOR_WHITE
    building_para.font.name = 'Noto Sans JP'
    building_para.alignment = PP_ALIGN.CENTER
    
    # タイトル「技術レポート」（サイズ縮小）
    title_box = slide.shapes.add_textbox(Inches(3), Inches(2.6), Inches(4), Inches(0.5))
    title_frame = title_box.text_frame
    title_frame.text = "技術レポート"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(28)
    title_para.font.bold = True
    title_para.font.color.rgb = COLOR_WHITE
    title_para.font.name = 'Noto Sans JP'
    title_para.alignment = PP_ALIGN.CENTER
    
    # 日付（サイズ縮小）
    date_box = slide.shapes.add_textbox(Inches(3.5), Inches(3.5), Inches(3), Inches(0.35))
    date_frame = date_box.text_frame
    date_frame.text = datetime.now().strftime('%Y.%m.%d')
    date_para = date_frame.paragraphs[0]
    date_para.font.size = Pt(16)
    date_para.font.color.rgb = COLOR_WHITE
    date_para.font.name = 'Noto Sans JP'
    date_para.alignment = PP_ALIGN.CENTER
    
    # 右上「BIM sustaina for Energy」
    sustaina_box = slide.shapes.add_textbox(Inches(7.5), Inches(0.2), Inches(2.3), Inches(0.5))
    sustaina_frame = sustaina_box.text_frame
    sustaina_frame.text = "BIM sustaina\nfor Energy"
    for para in sustaina_frame.paragraphs:
        para.font.size = Pt(10)
        para.font.color.rgb = COLOR_GRAY
        para.font.name = 'Noto Sans JP'
        para.alignment = PP_ALIGN.RIGHT
    
    # 右下「copyright © one building」
    copyright_box = slide.shapes.add_textbox(Inches(7.5), Inches(5.1), Inches(2.3), Inches(0.3))
    copyright_frame = copyright_box.text_frame
    copyright_frame.text = "copyright © one building"
    copyright_para = copyright_frame.paragraphs[0]
    copyright_para.font.size = Pt(8)
    copyright_para.font.color.rgb = COLOR_GRAY
    copyright_para.font.name = 'Noto Sans JP'
    copyright_para.alignment = PP_ALIGN.RIGHT

def add_summary_slide(prs, data):
    """総合評価サマリースライドを作成（充実版）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_title(slide, "1. 総合評価サマリー: 現状と経営リスク")
    
    # 建物基本情報ボックス
    info_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(4.5), Inches(0.8))
    info_frame = info_box.text_frame
    info_frame.word_wrap = True
    
    p_info = info_frame.paragraphs[0]
    p_info.text = f"建物名称: {data['building_name']}  |  所在地: {data['location']}  |  延べ面積: {data['total_area']} m²"
    p_info.font.size = Pt(11)
    p_info.font.color.rgb = COLOR_GRAY
    p_info.font.name = 'Noto Sans JP'
    
    # 診断結果ボックス
    result_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.9), Inches(4.5), Inches(0.5))
    result_frame = result_box.text_frame
    
    if data['bei_total'] > 1.0:
        status_text = "診断結果: 2024年基準非適合"
        status_color = COLOR_RED
        zeb_level = "ZEB基準未達成"
    elif data['bei_total'] > 0.80:
        status_text = "診断結果: H28年基準適合 / 2024年基準非適合"
        status_color = COLOR_RED
        zeb_level = "ZEB基準未達成"
    elif data['bei_total'] > 0.70:
        status_text = "診断結果: 2024年基準適合"
        status_color = COLOR_GREEN
        zeb_level = "ZEB Oriented相当（未認定）"
    elif data['bei_total'] > 0.50:
        status_text = "診断結果: 2024年基準適合"
        status_color = COLOR_GREEN
        zeb_level = "ZEB Ready相当（未認定）"
    else:
        status_text = "診断結果: 2024年基準適合"
        status_color = COLOR_GREEN
        zeb_level = "Nearly ZEB相当（未認定）"
    
    result_frame.text = status_text
    result_para = result_frame.paragraphs[0]
    result_para.font.size = Pt(18)
    result_para.font.bold = True
    result_para.font.color.rgb = status_color
    result_para.font.name = 'Noto Sans JP'
    
    # 経営リスクまたは優位性ボックス
    if data['bei_total'] > 1.0:
        # 経営リスク警告
        warning_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.5), Inches(2.6), Inches(9), Inches(1.8)
        )
        warning_shape.fill.solid()
        warning_shape.fill.fore_color.rgb = RGBColor(255, 243, 243)
        warning_shape.line.color.rgb = COLOR_RED
        warning_shape.line.width = Pt(2)
        
        warning_text = warning_shape.text_frame
        warning_text.word_wrap = True
        warning_text.margin_left = Inches(0.2)
        warning_text.margin_right = Inches(0.2)
        warning_text.margin_top = Inches(0.15)
        warning_text.margin_bottom = Inches(0.15)
        
        p1 = warning_text.paragraphs[0]
        p1.text = "▲重要: 経営リスクの特定"
        p1.font.size = Pt(16)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_RED
        p1.font.name = 'Noto Sans JP'
        p1.space_after = Pt(10)
        
        p2 = warning_text.add_paragraph()
        p2.text = f"、2024年4月施行の改正省エネ法（BEI 1.0以下義務化）に対応していません。"
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_BLACK
        p2.font.name = 'Noto Sans JP'
        p2.space_after = Pt(8)
        
        p3 = warning_text.add_paragraph()
        p3.text = "● 法的リスク: 適合判定（省エネ適判）をパスできず、建築確認申請が受理されません。"
        p3.font.size = Pt(12)
        p3.font.color.rgb = COLOR_BLACK
        p3.font.name = 'Noto Sans JP'
        p3.space_after = Pt(4)
        
        p4 = warning_text.add_paragraph()
        p4.text = "● 経済的リスク: 著しく高い光熱費が継続し、運用コストが垧迫します。"
        p4.font.size = Pt(12)
        p4.font.color.rgb = COLOR_BLACK
        p4.font.name = 'Noto Sans JP'
        p4.space_after = Pt(4)
        
        p5 = warning_text.add_paragraph()
        p5.text = "● 社会的リスク: ESG投資基準を満たせず、企業価値が低下する可能性があります。"
        p5.font.size = Pt(12)
        p5.font.color.rgb = COLOR_BLACK
        p5.font.name = 'Noto Sans JP'
        
        bei_y_pos = 4.6
    else:
        # 優位性・強みボックス
        advantage_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.5), Inches(2.6), Inches(9), Inches(1.8)
        )
        advantage_shape.fill.solid()
        advantage_shape.fill.fore_color.rgb = RGBColor(243, 255, 243)
        advantage_shape.line.color.rgb = COLOR_GREEN
        advantage_shape.line.width = Pt(2)
        
        advantage_text = advantage_shape.text_frame
        advantage_text.word_wrap = True
        advantage_text.margin_left = Inches(0.2)
        advantage_text.margin_right = Inches(0.2)
        advantage_text.margin_top = Inches(0.15)
        advantage_text.margin_bottom = Inches(0.15)
        
        p1 = advantage_text.paragraphs[0]
        p1.text = "▲優位性: 省エネ性能の強み"
        p1.font.size = Pt(16)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_GREEN
        p1.font.name = 'Noto Sans JP'
        p1.space_after = Pt(10)
        
        p2 = advantage_text.add_paragraph()
        p2.text = f"2024年基準に適合し、{zeb_level}の省エネ性能を有しています。"
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_BLACK
        p2.font.name = 'Noto Sans JP'
        p2.space_after = Pt(8)
        
        p3 = advantage_text.add_paragraph()
        p3.text = "● 法的適合: 省エネ適判をクリアし、建築確認申請がスムーズに進みます。"
        p3.font.size = Pt(12)
        p3.font.color.rgb = COLOR_BLACK
        p3.font.name = 'Noto Sans JP'
        p3.space_after = Pt(4)
        
        p4 = advantage_text.add_paragraph()
        p4.text = "● 経済的メリット: 光熱費が抑えられ、長期的な運用コスト削減が期待できます。"
        p4.font.size = Pt(12)
        p4.font.color.rgb = COLOR_BLACK
        p4.font.name = 'Noto Sans JP'
        p4.space_after = Pt(4)
        
        p5 = advantage_text.add_paragraph()
        p5.text = "● 社会的価値: ESG投資基準を満たし、企業価値向上に貢献します。"
        p5.font.size = Pt(12)
        p5.font.color.rgb = COLOR_BLACK
        p5.font.name = 'Noto Sans JP'
        
        bei_y_pos = 4.6
    
    # BEI表示ボックス（左側）
    bei_box = slide.shapes.add_textbox(Inches(5.5), Inches(1.0), Inches(2.0), Inches(1.3))
    bei_frame = bei_box.text_frame
    bei_label = get_bei_label(data.get('calculation_method', 'standard_input'))
    bei_frame.text = f"現在の {bei_label}\n{data['bei_total']:.2f}"
    
    p1 = bei_frame.paragraphs[0]
    p1.font.size = Pt(14)
    p1.font.color.rgb = COLOR_GRAY
    p1.font.name = 'Noto Sans JP'
    p1.alignment = PP_ALIGN.CENTER
    
    p2 = bei_frame.paragraphs[1]
    p2.font.size = Pt(48)
    p2.font.bold = True
    bei_color = COLOR_RED if data['bei_total'] > 1.0 else COLOR_GREEN
    p2.font.color.rgb = bei_color
    p2.font.name = 'Noto Sans JP'
    p2.alignment = PP_ALIGN.CENTER
    
    # 設備別BEIサマリー（右側）
    equipment_box = slide.shapes.add_textbox(Inches(7.7), Inches(1.0), Inches(2.0), Inches(3.4))
    equipment_frame = equipment_box.text_frame
    equipment_frame.word_wrap = True
    
    pe = equipment_frame.paragraphs[0]
    bei_label = get_bei_label(data.get('calculation_method', 'standard_input'))
    pe.text = f"設備別{bei_label}サマリー"
    pe.font.size = Pt(12)
    pe.font.bold = True
    pe.font.color.rgb = COLOR_MAIN
    pe.font.name = 'Noto Sans JP'
    pe.space_after = Pt(8)
    
    equipment_data = [
        ("空調", data['bei_ac']),
        ("換気", data['bei_v']),
        ("照明", data['bei_l']),
        ("給湯", data['bei_hw']),
        ("昇降機", data['bei_ev'])
    ]
    
    for name, value in equipment_data:
        if value > 0:
            p_eq = equipment_frame.add_paragraph()
            status_mark = "●" if value <= 1.0 else "▲"
            status_color = COLOR_GREEN if value <= 1.0 else COLOR_RED
            p_eq.text = f"{status_mark} {name}: {value:.2f}"
            p_eq.font.size = Pt(11)
            p_eq.font.color.rgb = status_color
            p_eq.font.name = 'Noto Sans JP'
            p_eq.space_after = Pt(4)

def add_envelope_worst_analysis_slide(prs, data):
    """外皮性能評価とPAL*ワースト要因分析スライドを作成"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_title(slide, "2. 外皮性能評価とPAL*ワースト要因の詳細分析")
    
    pal_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(2.8), Inches(1.8))
    pal_frame = pal_box.text_frame
    pal_frame.word_wrap = True
    
    p1 = pal_frame.paragraphs[0]
    p1.text = "外皮性能指標 (PAL*)"
    p1.font.size = Pt(15)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_MAIN
    p1.font.name = 'Noto Sans JP'
    p1.space_after = Pt(8)
    
    p2 = pal_frame.add_paragraph()
    p2.text = f"{data['pal_design']} MJ/m²・年"
    p2.font.size = Pt(38)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_MAIN
    p2.font.name = 'Noto Sans JP'
    p2.alignment = PP_ALIGN.CENTER
    p2.space_after = Pt(6)
    
    p3 = pal_frame.add_paragraph()
    bpi_color = COLOR_GREEN if data['bpi'] < 1.0 else COLOR_RED
    bpi_label = get_bpi_label(data.get('calculation_method', 'standard_input'))
    p3.text = f"{bpi_label} {data['bpi']:.2f} (基準 {data['pal_standard']} に対して{'良好' if data['bpi'] < 1.0 else '要改善'})"
    p3.font.size = Pt(11)
    p3.font.bold = True
    p3.font.color.rgb = bpi_color
    p3.font.name = 'Noto Sans JP'
    p3.alignment = PP_ALIGN.CENTER
    
    worst_box = slide.shapes.add_textbox(Inches(3.8), Inches(1.2), Inches(5.9), Inches(4))
    worst_frame = worst_box.text_frame
    worst_frame.word_wrap = True
    
    p1 = worst_frame.paragraphs[0]
    p1.text = "負荷要因の分解分析: ワースト室の特定"
    p1.font.size = Pt(15)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_MAIN
    p1.font.name = 'Noto Sans JP'
    p1.space_after = Pt(10)
    
    if data['worst_rooms']:
        worst_room = data['worst_rooms'][0]
        
        p2 = worst_frame.add_paragraph()
        p2.text = f"最大ワースト室: {worst_room['name']}"
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_RED
        p2.font.name = 'Noto Sans JP'
        p2.space_after = Pt(6)
        
        p3 = worst_frame.add_paragraph()
        p3.text = f"• 単位面積あたり負荷: {worst_room['q_per_a']:.1f} MJ/m²年\n• 基準値超過率: +{((worst_room['q_per_a'] - data['pal_standard']) / data['pal_standard'] * 100):.0f}%\n• 空調負荷: {worst_room['load']:.0f} MJ/年"
        p3.font.size = Pt(11)
        p3.font.name = 'Noto Sans JP'
        p3.space_after = Pt(10)
        
        p4 = worst_frame.add_paragraph()
        p4.text = "推定される負荷要因:"
        p4.font.size = Pt(12)
        p4.font.bold = True
        p4.font.color.rgb = COLOR_MAIN
        p4.font.name = 'Noto Sans JP'
        p4.space_after = Pt(6)
        
        p5 = worst_frame.add_paragraph()
        if 'ミーティング' in worst_room['name'] or '会議' in worst_room['name']:
            p5.text = "• 南側または西側の大きな窓面積による日射負荷\n• 会議室特有の高い人員密度と内部発熱\n• 外皮性能（窓の遮熱性能）の不足"
        elif '事務' in worst_room['name']:
            p5.text = "• OA機器による内部発熱\n• 窓面積比率が高い可能性\n• 断熱性能の不足"
        else:
            p5.text = "• 窓の性能（遮熱・断熱）の不足\n• 外壁断熱材の不足\n• 日射負荷の増大"
        p5.font.size = Pt(10)
        p5.font.name = 'Noto Sans JP'
        p5.space_after = Pt(10)
        
        p6 = worst_frame.add_paragraph()
        p6.text = "改善アドバイス:"
        p6.font.size = Pt(12)
        p6.font.bold = True
        p6.font.color.rgb = COLOR_MAIN
        p6.font.name = 'Noto Sans JP'
        p6.space_after = Pt(6)
        
        p7 = worst_frame.add_paragraph()
        p7.text = f"この室の窓の遮熱性能を向上させる（Low-Eガラスへの交換、外付けブラインド設置）ことで、建物全体のPAL*を約30 MJ/m²年改善可能。投資対効果が高い優先改善箇所です。"
        p7.font.size = Pt(10)
        p7.font.color.rgb = COLOR_BLACK
        p7.font.name = 'Noto Sans JP'
    else:
        p2 = worst_frame.add_paragraph()
        bei_label = get_bei_label(data.get('calculation_method', 'standard_input'))
        p2.text = f"外皮性能は全体的に良好です。設備側の改善に注力することで、{bei_label}目標達成が可能です。"
        p2.font.size = Pt(12)
        p2.font.name = 'Noto Sans JP'

def add_improvement_roadmap_slide(prs, data):
    """データに基づく柔軟な改善提案ロードマップスライドを作成"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_title(slide, "3. ZEB化への改善ロードマップ")
    
    roadmap = generate_improvement_roadmap(data)
    
    x_positions = [0.5, 2.8, 5.1, 7.4]
    
    for i, step_info in enumerate(roadmap):
        box_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x_positions[i]), Inches(1.7), Inches(2.1), Inches(1.7)
        )
        box_shape.fill.solid()
        box_shape.fill.fore_color.rgb = COLOR_BASE1
        box_shape.line.color.rgb = COLOR_MAIN
        box_shape.line.width = Pt(2)
        
        text_frame = box_shape.text_frame
        text_frame.word_wrap = True
        text_frame.margin_left = Inches(0.1)
        text_frame.margin_right = Inches(0.1)
        text_frame.margin_top = Inches(0.12)
        
        p1 = text_frame.paragraphs[0]
        p1.text = step_info['step']
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_MAIN
        p1.font.name = 'Noto Sans JP'
        p1.alignment = PP_ALIGN.CENTER
        p1.space_after = Pt(6)
        
        p2 = text_frame.add_paragraph()
        p2.text = step_info['title']
        p2.font.size = Pt(12)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_BLACK
        p2.font.name = 'Noto Sans JP'
        p2.alignment = PP_ALIGN.CENTER
        p2.space_after = Pt(6)
        
        p3 = text_frame.add_paragraph()
        p3.text = step_info['desc']
        p3.font.size = Pt(9)
        p3.font.color.rgb = COLOR_BLACK
        p3.font.name = 'Noto Sans JP'
        p3.alignment = PP_ALIGN.CENTER
    
    effect_box = slide.shapes.add_textbox(Inches(1.5), Inches(3.8), Inches(7), Inches(1.0))
    effect_frame = effect_box.text_frame
    effect_frame.word_wrap = True
    
    p1 = effect_frame.paragraphs[0]
    p1.text = "期待される総合効果"
    p1.font.size = Pt(15)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_MAIN
    p1.font.name = 'Noto Sans JP'
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(8)
    
    p2 = effect_frame.add_paragraph()
    target_bei = 0.70
    reduction = data['bei_total'] - target_bei
    bei_label = get_bei_label(data.get('calculation_method', 'standard_input'))
    p2.text = f"現在の{bei_label} {data['bei_total']:.2f} から ZEB Oriented 要件 {target_bei:.2f} まで、約 {reduction:.2f} の削減が必要です。\n上記4ステップの実施により、段階的に{bei_label}目標達成を目指します。"
    p2.font.size = Pt(11)
    p2.font.name = 'Noto Sans JP'
    p2.alignment = PP_ALIGN.CENTER

def main():
    """メイン処理"""
    import sys
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = '/home/ubuntu/upload/gemini_response_markdown.txt'
    chart_stacked_file = '/home/ubuntu/energy_stacked_bar_chart_v6.png'
    chart_pie_file = '/home/ubuntu/energy_pie_charts_v6.png'
    chart_bei_file = '/home/ubuntu/bei_chart_with_total_v6.png'
    output_file = '/home/ubuntu/Energy_Diagnosis_Report_v6.pptx'
    
    print("=" * 70)
    print("省エネ診断レポート生成ツール v6.0（最終版）")
    print("=" * 70)
    
    print("\n[1/4] データ抽出中...")
    data = extract_data_from_file(input_file)
    print(f"  建物名称: {data['building_name']}")
    print(f"  総合BEI: {data['bei_total']}")
    print(f"  外皮性能 PAL*: {data['pal_design']} MJ/m²・年")
    print(f"  BPI: {data['bpi']}")
    if data['worst_rooms']:
        print(f"  ワースト室: {data['worst_rooms'][0]['name']} ({data['worst_rooms'][0]['q_per_a']:.1f} MJ/m²年)")
    
    print("\n[2/4] グラフ生成中（Noto Sans JP・デザインガイドライン準拠）...")
    calc_method = data.get('calculation_method', 'standard_input')
    if calc_method == 'standard_input':
        print("　- エネルギー消費量積み上げ棒グラフ（レイアウト完全調整）")
        create_stacked_bar_chart_improved(data, chart_stacked_file)
        print("　- 設備別一次エネルギー消費量パイチャート（品のある色合い）")
        create_pie_charts(data, chart_pie_file)
    else:
        print("　- モデル建物法のためエネルギー消費量詳細グラフをスキップ")
    print("　- BEI比較グラフ（レイアウト調整版）")
    create_bei_comparison_chart_with_total(data, chart_bei_file)
    
    print("\n[3/4] PowerPointファイル生成中（16:9ワイドスクリーン）...")
    print("  - フォント: Noto Sans JP統一")
    print("  - デザインガイドライン準拠の品のある色合い")
    create_presentation(data, chart_stacked_file, chart_pie_file, chart_bei_file, output_file)
    
    print(f"\n[4/4] 完了")
    print(f"  PowerPointファイル: {output_file}")
    
    print("\n" + "=" * 70)
    print("処理が完了しました!")
    print("=" * 70)

if __name__ == '__main__':
    main()
