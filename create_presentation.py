# -*- coding: utf-8 -*-
"""
마케터 장상균 포트폴리오 — PPTX 생성 스크립트 (create_presentation.py)
Green Grid Editorial Theme & PPTX Skill Guidelines Applied
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# 1. Colors & Constants (Green Grid Editorial System)
COLOR_PRIMARY = RGBColor(38, 140, 58)    # #268C3A
COLOR_MID     = RGBColor(133, 194, 143)  # #85C28F
COLOR_LIGHT   = RGBColor(213, 235, 216)  # #D5EBD8
COLOR_TINT    = RGBColor(239, 247, 240)  # #EFF7F0
COLOR_BG      = RGBColor(249, 250, 248)  # #F9FAF8
COLOR_PANEL   = RGBColor(242, 242, 240)  # #F2F2F0
COLOR_WHITE   = RGBColor(255, 255, 255)
COLOR_DARK    = RGBColor(17, 17, 17)     # #111111
COLOR_TEXT    = RGBColor(51, 51, 51)     # #333333
COLOR_MUTED   = RGBColor(102, 102, 102)  # #666666

FONT_HEADING = "Pretendard"
FONT_BODY    = "Pretendard"
FONT_MONO    = "JetBrains Mono"
FONT_SERIF   = "Georgia"

BASE_DIR = r"C:\Users\ABC\.gemini\antigravity\scratch\portfolio-profile"
IMG_PROFILE = os.path.join(BASE_DIR, "assets", "images", "profile.png")
IMG_SMARTBIN = os.path.join(BASE_DIR, "assets", "images", "smartbin.png")
IMG_CHERRY = os.path.join(BASE_DIR, "assets", "images", "cherrycoffee.jpg")
IMG_POSTER = os.path.join(BASE_DIR, "assets", "images", "ai-mvp-poster.jpg")
IMG_TRASH = os.path.join(BASE_DIR, "assets", "images", "trash-cups-overflow.jpg")
IMG_STAFF = os.path.join(BASE_DIR, "assets", "images", "staff-sorting-interview.jpg")
IMG_FESTIVAL = os.path.join(BASE_DIR, "assets", "images", "songdo-beer-festival.jpg")
IMG_CUPS = os.path.join(BASE_DIR, "assets", "images", "transparent-pp-cups.jpg")
IMG_KDB = os.path.join(BASE_DIR, "assets", "images", "kdb-award-ceremony.jpg")
IMG_CHERRY_REELS = os.path.join(BASE_DIR, "assets", "images", "cherry-reels-garden.jpg")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

def set_slide_background(slide, color):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    return bg

def add_three_squares(slide, left, top, size=Inches(0.24)):
    colors = [COLOR_PRIMARY, COLOR_MID, COLOR_LIGHT]
    for i, c in enumerate(colors):
        sq = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left + i * size, top, size, size)
        sq.fill.solid()
        sq.fill.fore_color.rgb = c
        sq.line.fill.background()

def add_green_top_bar(slide, badge_text, title_text, right_sub_1="", right_sub_2=""):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.5), Inches(12.133), Inches(0.92))
    bar.fill.solid()
    bar.fill.fore_color.rgb = COLOR_PRIMARY
    bar.line.fill.background()

    # Pill badge
    if badge_text:
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(0.7), Inches(1.3), Inches(0.52))
        badge.fill.background()
        badge.line.color.rgb = COLOR_WHITE
        badge.line.width = Pt(1.5)
        tf = badge.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = badge_text
        run.font.name = FONT_HEADING
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = COLOR_WHITE

    # Title text
    tx_box = slide.shapes.add_textbox(Inches(2.3) if badge_text else Inches(0.85), Inches(0.65), Inches(7.0), Inches(0.6))
    tf = tx_box.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.name = FONT_HEADING
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = COLOR_WHITE

    # Right meta
    if right_sub_1 or right_sub_2:
        rt_box = slide.shapes.add_textbox(Inches(9.2), Inches(0.65), Inches(3.3), Inches(0.6))
        rt_tf = rt_box.text_frame
        rt_tf.margin_left = rt_tf.margin_top = rt_tf.margin_right = rt_tf.margin_bottom = 0
        p = rt_tf.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        if right_sub_1:
            r1 = p.add_run()
            r1.text = right_sub_1 + ("\n" if right_sub_2 else "")
            r1.font.name = FONT_BODY
            r1.font.size = Pt(10.5)
            r1.font.bold = True
            r1.font.color.rgb = COLOR_WHITE
        if right_sub_2:
            p2 = rt_tf.add_paragraph() if not right_sub_1 else p
            r2 = p2.add_run() if not right_sub_1 else p.add_run()
            r2.text = right_sub_2
            r2.font.name = FONT_BODY
            r2.font.size = Pt(10)
            r2.font.color.rgb = COLOR_WHITE

def add_main_frame_box(slide, top=Inches(1.42), height=Inches(5.4)):
    frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), top, Inches(12.133), height)
    frame.fill.solid()
    frame.fill.fore_color.rgb = COLOR_WHITE
    frame.line.color.rgb = COLOR_PRIMARY
    frame.line.width = Pt(3.5)
    return frame

# ==============================================================================
# SLIDE 1 : COVER
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
set_slide_background(s1, COLOR_BG)

# Upper Box (Expanded to Full Height, No Bottom Green Bar)
up1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.6), Inches(12.133), Inches(6.3))
up1.fill.solid()
up1.fill.fore_color.rgb = COLOR_WHITE
up1.line.color.rgb = COLOR_PRIMARY
up1.line.width = Pt(3.5)

# Top Left Meta
t_box = s1.shapes.add_textbox(Inches(1.0), Inches(0.95), Inches(4.0), Inches(0.8))
tf = t_box.text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]
r = p.add_run()
r.text = "2023 - 2026\nBRAND & CRM MARKETER"
r.font.name = FONT_HEADING
r.font.size = Pt(12)
r.font.bold = True
r.font.color.rgb = COLOR_DARK

# Giant Title: PORTFOLIO
t_box = s1.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(1.4))
tf = t_box.text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "PORTFOLIO"
r.font.name = FONT_HEADING
r.font.size = Pt(76)
r.font.bold = True
r.font.color.rgb = COLOR_DARK

# Subtitle
sub_box = s1.shapes.add_textbox(Inches(1.0), Inches(3.75), Inches(11.333), Inches(0.5))
tf = sub_box.text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "“브랜드의 본질을 파악하고, 데이터로 입증하는 마케터 장상균입니다.”"
r.font.name = FONT_HEADING
r.font.size = Pt(17.5)
r.font.bold = True
r.font.color.rgb = COLOR_PRIMARY

# Chips
chips_box = s1.shapes.add_textbox(Inches(1.0), Inches(4.45), Inches(11.333), Inches(0.45))
tf = chips_box.text_frame
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "#KDB IR 제안&피칭    #체리커피_매출25%상승    #포스터MVP_전환율7.7%"
r.font.name = FONT_BODY
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = COLOR_PRIMARY

# Bottom Footer Meta inside Box (Name Only, Right-Aligned)
f_line = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(6.12), Inches(11.333), Inches(0.015))
f_line.fill.solid()
f_line.fill.fore_color.rgb = COLOR_LIGHT
f_line.line.fill.background()

f_name = s1.shapes.add_textbox(Inches(7.0), Inches(6.22), Inches(5.333), Inches(0.45))
tf = f_name.text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.RIGHT
r = p.add_run()
r.text = "Jang, sang-gyun"
r.font.name = FONT_HEADING
r.font.size = Pt(16)
r.font.bold = True
r.font.color.rgb = COLOR_DARK

# ==============================================================================
# SLIDE 2 : 이력사항 (PROFILE)
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
set_slide_background(s2, COLOR_BG)
add_green_top_bar(s2, "", "이력사항", "JANG SANGGYUN", "PROFILE")
add_main_frame_box(s2)

# Left Gray Profile Panel
p_left = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(1.42), Inches(3.2), Inches(5.4))
p_left.fill.solid()
p_left.fill.fore_color.rgb = COLOR_PANEL
p_left.line.fill.background()

# Profile Photo
if os.path.exists(IMG_PROFILE):
    s2.shapes.add_picture(IMG_PROFILE, Inches(1.2), Inches(1.75), width=Inches(2.0))

p_info = s2.shapes.add_textbox(Inches(0.9), Inches(4.55), Inches(2.6), Inches(1.8))
tf = p_info.text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]
r = p.add_run()
r.text = "장상균\n"
r.font.name = FONT_HEADING
r.font.size = Pt(20)
r.font.bold = True
r.font.color.rgb = COLOR_DARK

r = p.add_run()
r.text = "Brand & CRM Marketer\n\n"
r.font.name = FONT_HEADING
r.font.size = Pt(11)
r.font.bold = True
r.font.color.rgb = COLOR_PRIMARY

contact_info = [
    ("T.", "010-5779-0849"),
    ("E.", "wkdtkdrbs11@naver.com"),
    ("U.", "숭실대 벤처중소기업학과")
]
for label, val in contact_info:
    p = tf.add_paragraph()
    p.space_before = Pt(4)
    r1 = p.add_run()
    r1.text = label + " "
    r1.font.name = FONT_HEADING
    r1.font.size = Pt(10.5)
    r1.font.bold = True
    r1.font.color.rgb = COLOR_PRIMARY

    r2 = p.add_run()
    r2.text = val
    r2.font.name = FONT_BODY
    r2.font.size = Pt(10.5)
    r2.font.color.rgb = COLOR_TEXT

# Center Column: Education & Experience
c_box = s2.shapes.add_textbox(Inches(4.1), Inches(1.65), Inches(4.0), Inches(5.0))
tf = c_box.text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

p = tf.paragraphs[0]
r = p.add_run()
r.text = "학력 및 병역사항"
r.font.name = FONT_HEADING
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = COLOR_PRIMARY

edu_items = [
    ("숭실대학교 (2021.03 - 2027.02 졸업예정)", "벤처중소기업학과 전공 (비즈니스 모델링, 스타트업 마케팅)"),
    ("육군 39보병사단 대테러특수임무대 3팀 (SDT)", "병장 만기전역 · 대테러 초동조치 및 특수임무 수행 (강한 체력과 현장 돌파력)")
]
for title, desc in edu_items:
    p = tf.add_paragraph()
    p.space_before = Pt(6)
    r = p.add_run()
    r.text = "○ " + title + "\n   " + desc
    r.font.name = FONT_BODY
    r.font.size = Pt(9.5)
    r.font.color.rgb = COLOR_TEXT

p = tf.add_paragraph()
p.space_before = Pt(16)
r = p.add_run()
r.text = "주요 경력 및 프로젝트"
r.font.name = FONT_HEADING
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = COLOR_PRIMARY

exp_items = [
    ("대학연합 창업동아리 (2026.07 - 현재)", "숭실지부장 / AI 교차검증 문제 정의 및 포스터 MVP 검증 총괄"),
    ("KDB 산업은행 창업 프로젝트 (2026.03 - 2026.07)", "스마트빈 기획 / 선별장 45건 인터뷰 리서치 총괄 및 IR 피칭"),
    ("카페 체리커피 군자점 (2023.07 - 2024.06 / 1년)", "지점 운영 총괄 / 위치 강조 릴스 제작(팔로워 +200) 및 매출 25% 상승")
]
for title, desc in exp_items:
    p = tf.add_paragraph()
    p.space_before = Pt(6)
    r = p.add_run()
    r.text = "○ " + title + "\n   " + desc
    r.font.name = FONT_BODY
    r.font.size = Pt(9.5)
    r.font.color.rgb = COLOR_TEXT

# Right Column: Awards, Certificates & Skills
r_box = s2.shapes.add_textbox(Inches(8.4), Inches(1.65), Inches(4.0), Inches(5.0))
tf = r_box.text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

p = tf.paragraphs[0]
r = p.add_run()
r.text = "수상 및 자격사항"
r.font.name = FONT_HEADING
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = COLOR_PRIMARY

cert_items = [
    ("KDB 산업은행 창업 프로젝트 장려상 수상 (2026.07)", "선별장 Skip형 고순도 r-PP 원료 직송 밸류체인 구축 및 피칭"),
    ("TOEIC 825점 (ETS 공인 어학성적)", "비즈니스 영어 커뮤니케이션 및 글로벌 마케팅 레퍼런스 리서치"),
    ("자동차운전면허 1종 보통", "오프라인 현장 기동 및 매장·팝업 물류 직접 운용 가능")
]
for title, desc in cert_items:
    p = tf.add_paragraph()
    p.space_before = Pt(6)
    r = p.add_run()
    r.text = "○ " + title + "\n   " + desc
    r.font.name = FONT_BODY
    r.font.size = Pt(9.5)
    r.font.color.rgb = COLOR_TEXT

p = tf.add_paragraph()
p.space_before = Pt(16)
r = p.add_run()
r.text = "업무 툴 및 스킬 역량"
r.font.name = FONT_HEADING
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = COLOR_PRIMARY

skills_list = [
    ("ChatGPT & AI 프롬프트", "5.0 / 5.0", "시장 리서치, 카피라이팅, 데이터 구조화"),
    ("Notion / DB 구축", "4.0 / 5.0", "일정 및 관계형 데이터베이스 관리"),
    ("Google Sheets / Excel", "4.0 / 5.0", "매출 지표 분석, 수치 트래킹, 피벗테이블"),
    ("Figma (콘텐츠/기획)", "3.0 / 5.0", "SNS 콘텐츠 디자인, 와이어프레임 기획"),
    ("Photoshop (포토샵)", "2.0 / 5.0", "이미지 보정, 포스터 편집, 비주얼 애셋 가공")
]
for name, score, desc in skills_list:
    p = tf.add_paragraph()
    p.space_before = Pt(5)
    r = p.add_run()
    r.text = f"• {name}  [{score}]\n   {desc}"
    r.font.name = FONT_BODY
    r.font.size = Pt(9)
    r.font.color.rgb = COLOR_TEXT

# ==============================================================================
# SLIDE 3 : INDEX.
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
set_slide_background(s3, COLOR_BG)

# Full Green Box
idx_box = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.6), Inches(12.133), Inches(6.3))
idx_box.fill.solid()
idx_box.fill.fore_color.rgb = COLOR_PRIMARY
idx_box.line.fill.background()

# Left Title
l_box = s3.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(3.5), Inches(5.0))
tf = l_box.text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]
r = p.add_run()
r.text = "JANG SANGGYUN\nPORTFOLIO\n\n\n"
r.font.name = FONT_HEADING
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = COLOR_WHITE

p2 = tf.add_paragraph()
r = p2.add_run()
r.text = "INDEX."
r.font.name = FONT_HEADING
r.font.size = Pt(64)
r.font.bold = True
r.font.color.rgb = COLOR_WHITE

# Right Entries
entries = [
    ("Project 1", "가상 브랜드 ‘스마트빈’ 기획", "축제·행사장 플라스틱 컵 선별 비용 절감 단일 PP 공급 및 직송 순환경제 플랫폼 (선별장 45건 인터뷰 & KDB 장려상)"),
    ("Project 2", "카페 ‘체리커피 군자점’ 운영 총괄", "차별화 릴스 소구점(위치 강조) 발굴, 매장 넛지 효과 적용 및 인근 식당 5곳 B2B 제휴로 매출 25% 상승"),
    ("Project 3", "대학생 AI 교차검증 문제 정의 및 MVP 실험", "‘AI 답변의 불필요한 교차검증 시간’으로 문제 쪼개기 및 오프라인 포스터 MVP로 사전 가입 전환율 7.7% 달성"),
    ("Story 3", "마케터 장상균의 3가지 이야기 (Philosophy & Attitude)", "F&B와 마케팅에 빠진 계기 · 본질을 파고드는 가설 검증 · 남들이 꺼리는 현장에 뛰어드는 협업 역량")
]

for idx, (pill, title, desc) in enumerate(entries):
    top_pos = Inches(1.2 + idx * 1.3)
    card = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.7), top_pos, Inches(7.5), Inches(1.18))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(46, 155, 67)
    card.line.color.rgb = COLOR_WHITE
    card.line.width = Pt(1)

    # Pill
    badge = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.95), top_pos + Inches(0.3), Inches(1.3), Inches(0.55))
    badge.fill.background()
    badge.line.color.rgb = COLOR_WHITE
    badge.line.width = Pt(1.5)
    tf = badge.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = pill
    r.font.name = FONT_HEADING
    r.font.size = Pt(11.5)
    r.font.bold = True
    r.font.color.rgb = COLOR_WHITE

    # Text
    e_box = s3.shapes.add_textbox(Inches(6.45), top_pos + Inches(0.14), Inches(5.6), Inches(0.9))
    tf = e_box.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title + "\n"
    r.font.name = FONT_HEADING
    r.font.size = Pt(15.5)  # Enlarged & Bolder
    r.font.bold = True
    r.font.color.rgb = COLOR_WHITE

    r = p.add_run()
    r.text = desc
    r.font.name = FONT_BODY
    r.font.size = Pt(9.5)
    r.font.color.rgb = COLOR_LIGHT

# ==============================================================================
# SLIDE 4 : Project 1 (스마트빈 — 문제정의 및 피벗)
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
set_slide_background(s4, COLOR_BG)
add_green_top_bar(s4, "Project 1", "문제정의 및 피벗", "기간: 2026.03.09 - 07.05", "")
add_main_frame_box(s4)

# 3 Columns
cols_s4 = [
    ("고객 및 시장 배경", "Customer & Market", IMG_FESTIVAL, [
        "워터밤·맥주축제 등 대형 행사장과 무인카페에서 단시간에 대량의 일회용 플라스틱 컵 배출",
        "컵 내부의 약 90%가 빈 공기 상태로 버려져 수거 부피가 폭증하고 잔여 음료 부패로 악취 유발",
        "주최 측은 막대한 청소 용역비와 분리 비용을 부담하나 실질 재활용률은 16.4%에 그침"
    ]),
    ("문제 상황", "Problem Definition", IMG_TRASH, [
        "초기에는 부피만 줄이는 ‘단순 3단 압축 쓰레기통’을 구상했으나 실제 처리 과정 검증 착수",
        "재질 혼합 문제: PET, PP, PS 등 서로 다른 재질이 섞여 압축되면 선별 불가로 원료 가치 0원",
        "세척 한계: 남은 음료 오염이 플라스틱 표면에 스며들면 사후 세척으로도 재생 원료화 불가"
    ]),
    ("선별장 45건 인터뷰", "Field Insight & Pivot", IMG_STAFF, [
        "쓰레기 선별장을 직접 방문해 작업자 및 축제 관계자 45건 심층 인터뷰 주도",
        "“수거함 판매가 아니라 투입 단계부터 단일 재질(PP)로 통제해야 돈이 되는 원료가 된다”는 인사이트 도출",
        "단순 하드웨어에서 ‘선별장 Skip형 고순도 r-PP 직송 밸류체인’으로 비즈니스 모델 전면 피벗"
    ])
]

for i, (title, sub, img_path, bullets) in enumerate(cols_s4):
    left_x = Inches(0.85 + i * 3.95)
    
    # Visual Box with Image
    vbox = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_x, Inches(1.7), Inches(3.65), Inches(1.4))
    vbox.fill.solid()
    vbox.fill.fore_color.rgb = COLOR_PANEL
    vbox.line.color.rgb = COLOR_PRIMARY
    vbox.line.width = Pt(1.5)

    if img_path and os.path.exists(img_path):
        s4.shapes.add_picture(img_path, left_x + Inches(0.04), Inches(1.74), width=Inches(3.57), height=Inches(1.32))

    # Title
    tbox = s4.shapes.add_textbox(left_x, Inches(3.22), Inches(3.65), Inches(0.75))
    tf = tbox.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = title + "\n"
    r.font.name = FONT_HEADING
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = COLOR_DARK

    r = p.add_run()
    r.text = sub
    r.font.name = FONT_HEADING
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY

    # Bullets
    bbox = s4.shapes.add_textbox(left_x, Inches(4.05), Inches(3.65), Inches(2.6))
    tf = bbox.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    for b_idx, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if b_idx == 0 else tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = "• " + bullet
        r.font.name = FONT_BODY
        r.font.size = Pt(12)
        r.font.color.rgb = COLOR_TEXT

# ==============================================================================
# SLIDE 5 : Project 1 (스마트빈 — 순환경제 솔루션 및 성과)
# Directly modeled after Reference 2 (AFTER NOISE campaign structure)
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
set_slide_background(s5, COLOR_BG)
add_green_top_bar(s5, "Project 1", "‘스마트빈’ 순환경제 솔루션 및 성과", "기간: 2026.03.09 - 07.05", "")
add_main_frame_box(s5)

# --- 3 PHASE STEP CARDS (Reference 2 Step Showcase) ---
cards_s5 = [
    ("1", "단일 재질(PP) 통제", "원료 순도 100% 확보", IMG_CUPS,
     "무지 PP컵 통일 ➔ 부스 일괄 공급 ➔ 혼입 원천 차단", [
         "사전에 로고 없는 무지 투명 PP 컵 3종만 행사장에 일괄 공급하여 타 재질 혼입 원천 차단",
         "입점 푸드트럭 및 부스 가이드라인 배포로 원료 순도 100% 회수 환경 조성"
     ]),
    ("2", "스마트빈 압축 & 수거", "수거 부피 70% 절감", IMG_SMARTBIN,
     "음료 분리 배출 ➔ 발 페달 3단 압축 ➔ Turn-key 직수거", [
         "음료 분리 배출 유도 후 발 페달식 물리적 압축 시스템으로 적재 부피 70% 이상 절감",
         "행사 종료 후 주최 측 부담 없이 전량 자체 회수(Turn-key 직수거) 전담"
     ]),
    ("3", "선별장 Skip 직송 & 성과", "KDB 장려상 & 고순도 원료화", IMG_KDB,
     "선별장 건너뛰기 ➔ r-PP 펠릿 직공급 ➔ KDB 장려상", [
         "복잡한 중간 선별장을 전면 건너뛰고 재생 플라스틱 공장에 고순도 r-PP 펠릿 원료 직공급",
         "KDB 산업은행 창업프로젝트 장려상 수상\n(IR 피칭을 통한 순환경제 BM 검증 및 성과 입증)"
     ])
]

for idx, (num_str, title_main, title_sub, img_p, flow_str, bullets) in enumerate(cards_s5):
    left_x = Inches(0.85 + idx * 3.95)
    
    # Outer Card Shape
    card = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_x, Inches(1.56), Inches(3.75), Inches(3.75))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_WHITE
    card.line.color.rgb = COLOR_PRIMARY
    card.line.width = Pt(1.5)

    # Number Circle
    num_c = s5.shapes.add_shape(MSO_SHAPE.OVAL, left_x + Inches(0.12), Inches(1.68), Inches(0.38), Inches(0.38))
    num_c.fill.solid()
    num_c.fill.fore_color.rgb = COLOR_PRIMARY
    num_c.line.fill.background()
    tf_n = num_c.text_frame
    tf_n.margin_left = tf_n.margin_top = tf_n.margin_right = tf_n.margin_bottom = 0
    p = tf_n.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = num_str
    r.font.name = FONT_HEADING
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_WHITE

    # Title Text
    tbox = s5.shapes.add_textbox(left_x + Inches(0.58), Inches(1.60), Inches(3.05), Inches(0.55))
    tf_t = tbox.text_frame
    tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
    p = tf_t.paragraphs[0]
    r = p.add_run()
    r.text = title_main + "\n"
    r.font.name = FONT_HEADING
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = COLOR_DARK
    r_sub = p.add_run()
    r_sub.text = title_sub
    r_sub.font.name = FONT_HEADING
    r_sub.font.size = Pt(10.5)
    r_sub.font.bold = True
    r_sub.font.color.rgb = COLOR_PRIMARY

    # Image Thumbnail
    if img_p and os.path.exists(img_p):
        s5.shapes.add_picture(img_p, left_x + Inches(0.12), Inches(2.2), width=Inches(3.51), height=Inches(1.18))

    # Micro Flow Pill
    pill = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_x + Inches(0.12), Inches(3.45), Inches(3.51), Inches(0.34))
    pill.fill.solid()
    pill.fill.fore_color.rgb = RGBColor(239, 247, 240)
    pill.line.color.rgb = COLOR_PRIMARY
    pill.line.width = Pt(1)
    tf_p = pill.text_frame
    tf_p.margin_left = tf_p.margin_top = tf_p.margin_right = tf_p.margin_bottom = 0
    p = tf_p.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = flow_str
    r.font.name = FONT_HEADING
    r.font.size = Pt(9)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY

    # Bullets
    bbox = s5.shapes.add_textbox(left_x + Inches(0.12), Inches(3.88), Inches(3.51), Inches(1.35))
    tf_b = bbox.text_frame
    tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0
    for b_idx, bullet in enumerate(bullets):
        p = tf_b.paragraphs[0] if b_idx == 0 else tf_b.add_paragraph()
        p.space_before = Pt(6)
        r = p.add_run()
        r.text = "• " + bullet
        r.font.name = FONT_BODY
        r.font.size = Pt(12)
        r.font.color.rgb = COLOR_TEXT


# --- FLOW SUMMARY LINE (Reference 2 Style) ---
flow_line = s5.shapes.add_textbox(Inches(0.85), Inches(5.40), Inches(11.65), Inches(0.38))
tf_fl = flow_line.text_frame
tf_fl.margin_left = tf_fl.margin_top = tf_fl.margin_right = tf_fl.margin_bottom = 0
p = tf_fl.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "무지 PP컵 공급   ➔   현장 70% 압축수거   ➔   중간 선별장 Skip   ➔   고순도 플라스틱 컵원료화"
r.font.name = FONT_HEADING
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = COLOR_DARK


# --- BIG BOTTOM GOAL PILL CONTAINER (Reference 2 Signature Component) ---
goal_pill = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(5.84), Inches(11.65), Inches(0.85))
goal_pill.fill.solid()
goal_pill.fill.fore_color.rgb = RGBColor(244, 248, 244)
goal_pill.line.color.rgb = COLOR_PRIMARY
goal_pill.line.width = Pt(2)

tf_g = goal_pill.text_frame
tf_g.margin_left = tf_g.margin_top = tf_g.margin_right = tf_g.margin_bottom = 0
p1 = tf_g.paragraphs[0]
p1.alignment = PP_ALIGN.CENTER
r1 = p1.add_run()
r1.text = "CORE VALUE & GOAL\n"
r1.font.name = FONT_HEADING
r1.font.size = Pt(9)
r1.font.bold = True
r1.font.color.rgb = COLOR_PRIMARY

p2 = tf_g.add_paragraph()
p2.alignment = PP_ALIGN.CENTER
r2 = p2.add_run()
r2.text = "“선별 비용을 없애고, 버려지는 폐기물을 돈이 되는 고순도 재생 원료로 전환”"
r2.font.name = FONT_HEADING
r2.font.size = Pt(14)
r2.font.bold = True
r2.font.color.rgb = COLOR_DARK
goal_pill.line.color.rgb = COLOR_PRIMARY
goal_pill.line.width = Pt(2)

tf_g = goal_pill.text_frame
tf_g.margin_left = tf_g.margin_top = tf_g.margin_right = tf_g.margin_bottom = 0
p1 = tf_g.paragraphs[0]
p1.alignment = PP_ALIGN.CENTER
r1 = p1.add_run()
r1.text = "CORE VALUE & GOAL\n"
r1.font.name = FONT_HEADING
r1.font.size = Pt(8.5)
r1.font.bold = True
r1.font.color.rgb = COLOR_PRIMARY

p2 = tf_g.add_paragraph()
p2.alignment = PP_ALIGN.CENTER
r2 = p2.add_run()
r2.text = "“선별 비용을 없애고, 버려지는 폐기물을 돈이 되는 고순도 재생 원료로 전환”"
r2.font.name = FONT_HEADING
r2.font.size = Pt(12.5)
r2.font.bold = True
r2.font.color.rgb = COLOR_DARK

# ==============================================================================
# SLIDE 6 : Project 2 (체리커피 군자점 — 업무 배경과 차별점 분석)
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout)
set_slide_background(s6, COLOR_BG)
add_green_top_bar(s6, "Project 2", "카페 ‘체리커피 군자점’ 운영 총괄", "기간: 2023.07.15 - 2024.06.15 (1년)", "")
add_main_frame_box(s6)

# Left 3 Vertical Bar Items
vbars_s6 = [
    ("고객 및 상권 과제", "Customer & Market", [
        "군자역 인근 다수의 프랜차이즈 및 개인 카페 사이에서 인지도와 재방문율을 높여야 하는 과제 직면",
        "피크 타임 외 유휴 시간대 유입 증대 및 지역 주민들의 단골 전환이 절실했음"
    ]),
    ("기존 홍보의 한계", "Problem", [
        "기존 카페 SNS 홍보는 단순히 공간 인테리어나 커피 사진 위주여서 주변 매장들과 차별점이 없었음",
        "“고객이 수많은 카페 중 왜 우리 매장을 찾아와야 하는가?”를 고객 관점에서 다시 정의함"
    ]),
    ("본질적 공간 강점 포착", "Core Insight", [
        "‘지하철역에서 도보 3분 거리의 뛰어난 역세권 접근성을 가지면서도, 도심 속에서 마치 교외 대형 정원 카페에 온 듯한 여유로운 자연 채광과 테라스를 누릴 수 있다’는 핵심 가치를 차별점으로 도출"
    ])
]

for idx, (title, sub, bullets) in enumerate(vbars_s6):
    top_pos = Inches(1.68 + idx * 1.62)
    # Green Bar
    bar = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), top_pos, Inches(0.08), Inches(1.4))
    bar.fill.solid()
    bar.fill.fore_color.rgb = COLOR_PRIMARY
    bar.line.fill.background()

    tbox = s6.shapes.add_textbox(Inches(1.15), top_pos, Inches(5.4), Inches(1.4))
    tf = tbox.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title + "  "
    r.font.name = FONT_HEADING
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = COLOR_DARK

    r = p.add_run()
    r.text = sub
    r.font.name = FONT_HEADING
    r.font.size = Pt(10.5)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY

    for bullet in bullets:
        p = tf.add_paragraph()
        p.space_before = Pt(6)
        r = p.add_run()
        r.text = "• " + bullet
        r.font.name = FONT_BODY
        r.font.size = Pt(12)
        r.font.color.rgb = COLOR_TEXT

# Right Showcase
r_frame = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.0), Inches(1.6), Inches(5.4), Inches(4.9))
r_frame.fill.solid()
r_frame.fill.fore_color.rgb = COLOR_PANEL
r_frame.line.color.rgb = COLOR_PRIMARY
r_frame.line.width = Pt(1.5)

c_title = s6.shapes.add_textbox(Inches(7.2), Inches(1.8), Inches(5.0), Inches(0.4))
tf = c_title.text_frame
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "핵심 포지셔닝 : “역세권 도보 3분, 도심 속 교외 정원 카페”"
r.font.name = FONT_HEADING
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = COLOR_PRIMARY

# Centered Garden Reels Frame (Extended Height)
insta_bg = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.65), Inches(2.15), Inches(2.1), Inches(3.1))
insta_bg.fill.solid()
insta_bg.fill.fore_color.rgb = RGBColor(0, 0, 0)
insta_bg.line.color.rgb = COLOR_PRIMARY
insta_bg.line.width = Pt(1.5)

if os.path.exists(IMG_CHERRY_REELS):
    s6.shapes.add_picture(IMG_CHERRY_REELS, Inches(8.85), Inches(2.2), height=Inches(3.0))

post_box = s6.shapes.add_textbox(Inches(7.2), Inches(5.35), Inches(5.0), Inches(1.0))
tf = post_box.text_frame
p = tf.paragraphs[0]
r = p.add_run()
r.text = "차별화 전략 (Positioning Shift) :\n"
r.font.name = FONT_HEADING
r.font.size = Pt(11.5)
r.font.bold = True
r.font.color.rgb = COLOR_PRIMARY

r = p.add_run()
r.text = "단순 감성 인테리어 사진 지양 → 지하철역 출구부터 매장 정원 테라스로 이어지는 접근성과 공간감을 강조한 숏폼 릴스로 전환"
r.font.name = FONT_BODY
r.font.size = Pt(11)
r.font.color.rgb = COLOR_TEXT

# ==============================================================================
# SLIDE 7 : Project 2 결과물 (체리커피 릴스 & 매출 25% 상승)
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout)
set_slide_background(s7, COLOR_BG)
add_green_top_bar(s7, "Project 2", "팔로워 200명 증가 및 전월 대비 매출 25% 상승", "기간: 2023.07.15 - 2024.06.15", "")
add_main_frame_box(s7)

# Left Showcase (Reels Photo + Concise Comparison Graph)
c_panel = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(1.6), Inches(5.4), Inches(5.0))
c_panel.fill.solid()
c_panel.fill.fore_color.rgb = COLOR_PANEL
c_panel.line.color.rgb = COLOR_PRIMARY
c_panel.line.width = Pt(1.5)

# Profile Picture in Dark Phone Frame
phone_frame = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.1), Inches(1.8), Inches(2.4), Inches(4.55))
phone_frame.fill.solid()
phone_frame.fill.fore_color.rgb = RGBColor(12, 15, 20)
phone_frame.line.color.rgb = COLOR_PRIMARY
phone_frame.line.width = Pt(1.5)

if os.path.exists(IMG_CHERRY):
    s7.shapes.add_picture(IMG_CHERRY, Inches(1.15), Inches(2.75), width=Inches(2.3))

# Graph & Stat on Right Side of Left Panel
g_title = s7.shapes.add_textbox(Inches(3.65), Inches(1.8), Inches(2.55), Inches(0.4))
tf = g_title.text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]
r = p.add_run()
r.text = "캠페인 성과"
r.font.name = FONT_HEADING
r.font.size = Pt(12)
r.font.bold = True
r.font.color.rgb = COLOR_DARK

# Metric 1: Revenue +25%
m1_box = s7.shapes.add_textbox(Inches(3.65), Inches(2.35), Inches(2.55), Inches(0.8))
tf = m1_box.text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]
r = p.add_run()
r.text = "매장 매출 : "
r.font.name = FONT_BODY
r.font.size = Pt(10)
r.font.bold = True
r2 = p.add_run()
r2.text = "+25% UP\n"
r2.font.name = FONT_HEADING
r2.font.size = Pt(13)
r2.font.bold = True
r2.font.color.rgb = COLOR_PRIMARY
r3 = p.add_run()
r3.text = "전월 대비 매출 25% 상승 달성"
r3.font.name = FONT_BODY
r3.font.size = Pt(8.5)
r3.font.color.rgb = COLOR_MUTED

# Bar 1 Track & Fill
b1_track = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.65), Inches(3.2), Inches(2.4), Inches(0.12))
b1_track.fill.solid(); b1_track.fill.fore_color.rgb = RGBColor(220, 220, 220); b1_track.line.fill.background()
b1_fill = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.65), Inches(3.2), Inches(2.4), Inches(0.12))
b1_fill.fill.solid(); b1_fill.fill.fore_color.rgb = COLOR_PRIMARY; b1_fill.line.fill.background()

# Metric 2: Followers +200
m2_box = s7.shapes.add_textbox(Inches(3.65), Inches(3.55), Inches(2.55), Inches(0.8))
tf = m2_box.text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]
r = p.add_run()
r.text = "인스타 팔로워 : "
r.font.name = FONT_BODY
r.font.size = Pt(10)
r.font.bold = True
r2 = p.add_run()
r2.text = "+200명 (646명)\n"
r2.font.name = FONT_HEADING
r2.font.size = Pt(13)
r2.font.bold = True
r2.font.color.rgb = COLOR_PRIMARY
r3 = p.add_run()
r3.text = "광고비 0원 100% 바이럴 유입"
r3.font.name = FONT_BODY
r3.font.size = Pt(8.5)
r3.font.color.rgb = COLOR_MUTED

# Bar 2 Track & Fill
b2_track = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.65), Inches(4.4), Inches(2.4), Inches(0.12))
b2_track.fill.solid(); b2_track.fill.fore_color.rgb = RGBColor(220, 220, 220); b2_track.line.fill.background()
b2_fill = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.65), Inches(4.4), Inches(2.0), Inches(0.12))
b2_fill.fill.solid(); b2_fill.fill.fore_color.rgb = COLOR_PRIMARY; b2_fill.line.fill.background()

# Stat Pills
pill1 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.65), Inches(4.75), Inches(2.45), Inches(0.52))
pill1.fill.solid(); pill1.fill.fore_color.rgb = COLOR_WHITE; pill1.line.color.rgb = COLOR_PRIMARY; pill1.line.width = Pt(1)
p = pill1.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "⚡ 광고비 0원 숏폼 도달"; r.font.name = FONT_BODY; r.font.size = Pt(9.5); r.font.bold = True; r.font.color.rgb = COLOR_DARK

pill2 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.65), Inches(5.42), Inches(2.45), Inches(0.52))
pill2.fill.solid(); pill2.fill.fore_color.rgb = COLOR_WHITE; pill2.line.color.rgb = COLOR_PRIMARY; pill2.line.width = Pt(1)
p = pill2.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "🤝 인근 식당 5곳 B2B 제휴"; r.font.name = FONT_BODY; r.font.size = Pt(9.5); r.font.bold = True; r.font.color.rgb = COLOR_DARK


# Right 3 Action Pillars
vbars_s7 = [
    ("공간 강점을 살린 숏폼 릴스 제작", "Content", [
        "‘역세권 도심 속 숨은 숲속 정원 카페’를 테마로 접근성과 공간감을 강조한 릴스를 제작·배포해 1개월 만에 팔로워 200명 증가"
    ]),
    ("매장 경험(CX) 및 넛지(Nudge) 진열", "Store CX", [
        "고객이 주문 카운터로 향하는 시선 동선에 당일 구운 시그니처 베이커리를 전략적으로 배치해 음료+베이커리 동반 구매 및 객단가 상승 유도"
    ]),
    ("인근 인기 식당 5곳 B2B 상생 제휴", "Promotion", [
        "주변 식당 5곳과 협력해 식사 영수증 지참 시 음료 할인 및 매장 내 상호 쿠폰 비치를 진행, 점심 피크 직후 유휴 시간대 유입 극대화"
    ])
]

for idx, (title, sub, bullets) in enumerate(vbars_s7):
    top_pos = Inches(1.68 + idx * 1.62)
    bar = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.8), top_pos, Inches(0.08), Inches(1.4))
    bar.fill.solid(); bar.fill.fore_color.rgb = COLOR_PRIMARY; bar.line.fill.background()

    tbox = s7.shapes.add_textbox(Inches(7.05), top_pos, Inches(5.4), Inches(1.4))
    tf = tbox.text_frame; tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = title + "  "; r.font.name = FONT_HEADING; r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = COLOR_DARK
    r = p.add_run(); r.text = sub; r.font.name = FONT_HEADING; r.font.size = Pt(10.5); r.font.bold = True; r.font.color.rgb = COLOR_PRIMARY
    for bullet in bullets:
        p = tf.add_paragraph(); p.space_before = Pt(6)
        r = p.add_run(); r.text = "• " + bullet; r.font.name = FONT_BODY; r.font.size = Pt(12); r.font.color.rgb = COLOR_TEXT

# ==============================================================================
# SLIDE 8 : Project 3 (대학생 AI 교차검증 문제 정의 및 포스터 실험)
# ==============================================================================
s8 = prs.slides.add_slide(blank_layout)
set_slide_background(s8, COLOR_BG)
add_green_top_bar(s8, "Project 3", "대학생 AI 활용 문제 정의 쪼개기와 병목 돌파", "기간: 2026.07.10 - 현재", "역할: 대학연합 창업동아리 숭실지부장")
add_main_frame_box(s8)

# Left 3 Vertical Bar Items
vbars_s8 = [
    ("초기 문제 정의와 멘토 피드백", "Initial Hypothesis", [
        "대학생의 AI 활용 문제를 정의하며, 처음에는 ‘AI 답변의 신뢰성이 낮아 여러 AI를 교차 검증을 하느라 시간을 낭비한다’라고 보았음",
        "그러나 멘토님의 “신뢰란 무엇인가”라는 피드백을 받고 문제를 더 쪼개야겠다고 판단함"
    ]),
    ("‘신뢰’와 ‘시간 낭비’의 본질 분해", "Problem Breakdown", [
        "팀원들과 논의해 ‘신뢰’는 답변의 정확성, ‘불필요한 시간’은 정확성을 확인하는 교차검증 시간으로 뾰족하게 구체화함",
        "전공 교재 풀이과정이 제공되지 않아 80%의 학생이 겪는 불안감을 타겟 문제로 정밀 타겟팅"
    ]),
    ("1달간의 실험 준비 병목 돌파", "Action Leadership", [
        "실험 준비가 한 달가량 이어져 병목이 생기자, 완성도를 더 높이기보다 고객 반응을 먼저 검증하자고 제안해 팀의 동의를 얻음",
        "막연한 개발 준비를 멈추고 현장 오프라인 포스터를 통해 신속한 가설 검증 단행"
    ])
]

for idx, (title, sub, bullets) in enumerate(vbars_s8):
    top_pos = Inches(1.7 + idx * 1.55)
    bar = s8.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), top_pos, Inches(0.08), Inches(1.3))
    bar.fill.solid(); bar.fill.fore_color.rgb = COLOR_PRIMARY; bar.line.fill.background()

    tbox = s8.shapes.add_textbox(Inches(1.15), top_pos, Inches(5.4), Inches(1.3))
    tf = tbox.text_frame; tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = title + "  "; r.font.name = FONT_HEADING; r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = COLOR_DARK
    r = p.add_run(); r.text = sub; r.font.name = FONT_HEADING; r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = COLOR_PRIMARY
    for bullet in bullets:
        p = tf.add_paragraph(); p.space_before = Pt(4)
        r = p.add_run(); r.text = "• " + bullet; r.font.name = FONT_BODY; r.font.size = Pt(9.5); r.font.color.rgb = COLOR_TEXT

# Right Poster Showcase
p_frame = s8.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.0), Inches(1.7), Inches(5.4), Inches(4.7))
p_frame.fill.solid(); p_frame.fill.fore_color.rgb = COLOR_WHITE; p_frame.line.color.rgb = COLOR_PRIMARY; p_frame.line.width = Pt(1.5)

tf = s8.shapes.add_textbox(Inches(7.2), Inches(1.85), Inches(5.0), Inches(0.4)).text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "실제 배포한 오프라인 실험 포스터 MVP"; r.font.name = FONT_HEADING; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COLOR_PRIMARY

if os.path.exists(IMG_POSTER):
    s8.shapes.add_picture(IMG_POSTER, Inches(7.3), Inches(2.3), width=Inches(4.8))

tf = s8.shapes.add_textbox(Inches(7.2), Inches(5.95), Inches(5.0), Inches(0.35)).text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "▲ 교수님을 통해 실제 배포한 오프라인 포스터 (QR 유입 39명 중 3명 가입, 전환율 7.7%)"; r.font.size = Pt(8.5); r.font.color.rgb = COLOR_MUTED

# ==============================================================================
# SLIDE 9 : 경험 3 결과물 (포스터 실험 전환율 7.7% 및 러닝)
# ==============================================================================
s9 = prs.slides.add_slide(blank_layout)
set_slide_background(s9, COLOR_BG)
add_green_top_bar(s9, "결과물", "포스터 MVP 실험 전환율 7.7% 달성 및 실행 기준 정립", "JANG SANGGYUN", "PORTFOLIO")
add_main_frame_box(s9)

# Left 3 Execution Steps
vbars_s9 = [
    ("오프라인 포스터 MVP 설계", "MVP Design", [
        "“AI마다 답이 달라서 못 믿겠어요 / 풀이과정이 없어 내 풀이가 맞는지 확인할 수 없어요”라는 실제 학생들의 언어로 카피 작성",
        "핵심 페인포인트와 선배 Q&A QR코드를 결합하여 직관적으로 배치"
    ]),
    ("교수님 협조를 통한 타겟 배포", "Channel Strategy", [
        "타겟 학생들이 모이는 핵심 이동 동선인 전공 강의실 및 게시판에 포스터를 직접 배포",
        "지도 교수님의 공식 협조를 이끌어내어 정보 신뢰도와 참여율을 크게 향상시킴"
    ]),
    ("개인정보 입력 전환으로 진성 수요 검증", "Validation Metric", [
        "단순 QR 스캔에 그치지 않고 ‘이름, 학번, 연락처’를 직접 입력하고 가입하는 높은 전환 장벽을 설정",
        "39명 유입 중 3명이 실제 개인정보를 입력하고 가입하여 7.7%의 유의미한 사전 전환율 입증"
    ])
]

for idx, (title, sub, bullets) in enumerate(vbars_s9):
    top_pos = Inches(1.7 + idx * 1.45)
    bar = s9.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), top_pos, Inches(0.08), Inches(1.2))
    bar.fill.solid(); bar.fill.fore_color.rgb = COLOR_PRIMARY; bar.line.fill.background()

    tbox = s9.shapes.add_textbox(Inches(1.15), top_pos, Inches(5.4), Inches(1.2))
    tf = tbox.text_frame; tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = title + "  "; r.font.name = FONT_HEADING; r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = COLOR_DARK
    r = p.add_run(); r.text = sub; r.font.name = FONT_HEADING; r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = COLOR_PRIMARY
    for bullet in bullets:
        p = tf.add_paragraph(); p.space_before = Pt(4)
        r = p.add_run(); r.text = "• " + bullet; r.font.name = FONT_BODY; r.font.size = Pt(9.5); r.font.color.rgb = COLOR_TEXT

# Right Funnel Showcase
f_frame = s9.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.0), Inches(1.7), Inches(5.4), Inches(4.3))
f_frame.fill.solid(); f_frame.fill.fore_color.rgb = COLOR_PANEL; f_frame.line.color.rgb = COLOR_PRIMARY; f_frame.line.width = Pt(1.5)

tf = s9.shapes.add_textbox(Inches(7.2), Inches(1.9), Inches(5.0), Inches(0.4)).text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "오프라인 포스터 실험 퍼널 : 가입 전환율 7.7% 달성"; r.font.name = FONT_HEADING; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COLOR_PRIMARY

# Funnel Steps
fn_steps = [
    ("STEP 1. 타겟 접점 확보", "교수님 협조 기반 학과 수강생 대상 포스터 배포", "타겟 수강생 전원"),
    ("STEP 2. 포스터 QR 스캔 유입", "문제 메시지에 공감해 랜딩 페이지로 유입", "39명"),
    ("STEP 3. 개인정보 입력 및 가입", "높은 전환 허들을 넘어 실제 사전 가입 완료", "3명 (전환율 7.7%)")
]

for s_idx, (st_t, st_d, st_v) in enumerate(fn_steps):
    step_y = Inches(2.4 + s_idx * 1.1)
    card = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.3), step_y, Inches(4.8), Inches(0.95))
    if s_idx == 2:
        card.fill.solid(); card.fill.fore_color.rgb = COLOR_PRIMARY; card.line.fill.background()
        t_col = COLOR_WHITE; v_col = COLOR_WHITE
    else:
        card.fill.solid(); card.fill.fore_color.rgb = COLOR_WHITE; card.line.color.rgb = COLOR_PRIMARY; card.line.width = Pt(1.5)
        t_col = COLOR_DARK; v_col = COLOR_PRIMARY

    tf = s9.shapes.add_textbox(Inches(7.5), step_y + Inches(0.12), Inches(3.2), Inches(0.7)).text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = st_t + "\n"; r.font.name = FONT_HEADING; r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = t_col
    r = p.add_run(); r.text = st_d; r.font.name = FONT_BODY; r.font.size = Pt(9); r.font.color.rgb = t_col

    tf2 = s9.shapes.add_textbox(Inches(10.7), step_y + Inches(0.2), Inches(1.3), Inches(0.5)).text_frame
    tf2.margin_left = tf2.margin_top = tf2.margin_right = tf2.margin_bottom = 0
    p2 = tf2.paragraphs[0]; p2.alignment = PP_ALIGN.RIGHT
    r2 = p2.add_run(); r2.text = st_v; r2.font.name = FONT_MONO; r2.font.size = Pt(12 if s_idx==2 else 13); r2.font.bold = True; r2.font.color.rgb = v_col

# Bottom Gray Summary Strip
strip = s9.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(6.15), Inches(12.133), Inches(0.65))
strip.fill.solid(); strip.fill.fore_color.rgb = COLOR_PANEL; strip.line.fill.background()

s_box = s9.shapes.add_textbox(Inches(1.0), Inches(6.28), Inches(11.333), Inches(0.4))
tf = s_box.text_frame; tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]
r = p.add_run(); r.text = "배운 점 : "; r.font.name = FONT_HEADING; r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = COLOR_DARK
r = p.add_run(); r.text = "브랜드의 정체성과 메시지를 실험할 때, 문제를 깊이 정의하되 적시에 검증해 다음 판단의 근거를 만드는 것"; r.font.name = FONT_HEADING; r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = COLOR_PRIMARY
r = p.add_run(); r.text = "이 중요함을 배웠습니다."; r.font.name = FONT_BODY; r.font.size = Pt(11); r.font.color.rgb = COLOR_TEXT

# ==============================================================================
# SLIDE 10 : Story 3 (마케터 장상균의 3가지 이야기)
# ==============================================================================
s10 = prs.slides.add_slide(blank_layout)
set_slide_background(s10, COLOR_BG)
add_green_top_bar(s10, "Story 3", "마케터 장상균의 3가지 이야기", "JANG SANGGYUN", "PHILOSOPHY")
add_main_frame_box(s10)

# Left Gray Panel
p_left = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(1.42), Inches(3.4), Inches(5.4))
p_left.fill.solid(); p_left.fill.fore_color.rgb = COLOR_PANEL; p_left.line.fill.background()

tf = s10.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(2.6), Inches(4.5)).text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]
r = p.add_run(); r.text = "WORK ETHIC\n\n"; r.font.name = FONT_HEADING; r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = COLOR_PRIMARY

r = p.add_run(); r.text = "본질을 묻고,\n현장과 데이터로\n답합니다\n\n"; r.font.name = FONT_HEADING; r.font.size = Pt(18); r.font.bold = True; r.font.color.rgb = COLOR_DARK

r = p.add_run(); r.text = "눈앞의 현상에 머무르지 않고 고객이 행동하는 진짜 이유를 파고들며, 탁상공론 대신 빠른 실험과 현장 협업으로 실질적인 비즈니스 임팩트를 만듭니다."
r.font.name = FONT_BODY; r.font.size = Pt(10.5); r.font.color.rgb = COLOR_TEXT

# Right 3 Stories
stories = [
    ("01", "1. F&B와 마케팅에 빠져들게 된 계기",
     "1년간 카페를 총괄 운영하며 마케팅의 작은 디테일 하나가 실제 매장의 매출 곡선을 바꾼다는 사실을 체감했습니다. 베이커리 진열에 넛지(Nudge) 효과를 적용하고 동선을 재배치해 매장 경험(CX)을 최적화했으며, 주변 매장과 똑같은 인테리어 사진 대신 ‘지하철역과 가까우면서도 교외 대형 카페처럼 여유로운 분위기를 즐길 수 있다’는 역세권 정원 강점을 잡아 릴스를 제작해 1달 만에 팔로워 200명 증가, 매장 매출 25% 상승을 이끌어냈습니다."),
    ("02", "2. 본질을 파고드는 집요함과 적시 가설 검증",
     "제 강점은 문제의 본질을 파악하고 실행으로 옮기는 추진력입니다. 연합동아리 창업 프로젝트에서 대학생의 AI 활용 문제를 정의할 때, 멘토님의 “신뢰란 무엇인가”라는 피드백을 계기로 ‘신뢰’는 답변의 정확성, ‘불필요한 시간’은 정확성을 확인하는 교차검증 시간으로 구체화했습니다. 실험 준비가 한 달가량 길어지자 완성도를 높이기보다 고객 반응을 먼저 검증하자고 제안해 교수님을 통한 포스터 실험으로 39명 중 3명 가입(7.7% 전환율)을 확인하며 다음 판단의 근거를 만들었습니다."),
    ("03", "3. 남들이 꺼리는 현장에 뛰어드는 협업 역량",
     "스마트빈 프로젝트 초기, 열악한 환경 탓에 누구나 꺼리던 재활용 선별장 방문 인터뷰(총 45건)에 먼저 나서 작업자분들의 목소리를 확보했습니다. 현장에서 ‘단순 압축 시 플라스틱 재질(PET/PP)이 섞여 재활용 가치가 0원이 된다’는 치명적 병목을 발견해 팀과 투명하게 공유했고, 팀은 ‘무지 PP컵 공급 + 압축 수거 + 재생 원료 직공급’의 순환경제 밸류체인으로 피벗하여 KDB 창업프로젝트 장려상을 수상했습니다.")
]

for idx, (num, title, body) in enumerate(stories):
    top_pos = Inches(1.65 + idx * 1.7)
    
    # Icon box
    ib = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.3), top_pos, Inches(0.65), Inches(0.65))
    ib.fill.solid(); ib.fill.fore_color.rgb = COLOR_TINT; ib.line.color.rgb = COLOR_PRIMARY; ib.line.width = Pt(1.5)
    tf = ib.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = num; r.font.name = FONT_HEADING; r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = COLOR_PRIMARY

    # Story text
    s_box = s10.shapes.add_textbox(Inches(5.15), top_pos, Inches(7.3), Inches(1.55))
    tf = s_box.text_frame; tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = title + "\n"; r.font.name = FONT_HEADING; r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = COLOR_PRIMARY

    r = p.add_run(); r.text = body; r.font.name = FONT_BODY; r.font.size = Pt(9.5); r.font.color.rgb = COLOR_TEXT

# ==============================================================================
# SLIDE 11 : THANK YOU & CONTACT
# ==============================================================================
s11 = prs.slides.add_slide(blank_layout)
set_slide_background(s11, COLOR_BG)

up11 = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.6), Inches(12.133), Inches(4.9))
up11.fill.solid(); up11.fill.fore_color.rgb = COLOR_WHITE; up11.line.color.rgb = COLOR_PRIMARY; up11.line.width = Pt(3.5)

tf = s11.shapes.add_textbox(Inches(1.0), Inches(0.9), Inches(4.0), Inches(0.8)).text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]
r = p.add_run(); r.text = "JANG SANGGYUN\nBRAND & CRM MARKETER"; r.font.name = FONT_HEADING; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COLOR_DARK

# Giant Title: THANK YOU
tf = s11.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(11.333), Inches(1.4)).text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "THANK YOU"; r.font.name = FONT_HEADING; r.font.size = Pt(72); r.font.bold = True; r.font.color.rgb = COLOR_DARK

# Subtitle
tf = s11.shapes.add_textbox(Inches(1.0), Inches(3.3), Inches(11.333), Inches(0.5)).text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "검토해 주셔서 감사합니다. 함께 더 큰 성장을 만들 준비가 되어 있습니다."; r.font.name = FONT_HEADING; r.font.size = Pt(17); r.font.bold = True; r.font.color.rgb = COLOR_PRIMARY

# Chips
tf = s11.shapes.add_textbox(Inches(1.0), Inches(3.95), Inches(11.333), Inches(0.4)).text_frame
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "#즉시_실무_투입_가능    #TOEIC_825점_1종보통    #SDT_특수임무대_만기전역"; r.font.name = FONT_BODY; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COLOR_PRIMARY

# Bottom Green Bar
btm11 = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(5.7), Inches(12.133), Inches(1.2))
btm11.fill.solid(); btm11.fill.fore_color.rgb = COLOR_PRIMARY; btm11.line.fill.background()

tf = s11.shapes.add_textbox(Inches(1.0), Inches(6.05), Inches(3.0), Inches(0.6)).text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]
r = p.add_run(); r.text = "Contact"; r.font.name = FONT_HEADING; r.font.size = Pt(22); r.font.bold = True; r.font.color.rgb = COLOR_WHITE

tf = s11.shapes.add_textbox(Inches(4.2), Inches(6.1), Inches(8.1), Inches(0.5)).text_frame
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT

r = p.add_run(); r.text = "📞 010-5779-0849   "; r.font.name = FONT_HEADING; r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = COLOR_WHITE
r = p.add_run(); r.text = "  ✉️ wkdtkdrbs11@naver.com   "; r.font.name = FONT_BODY; r.font.size = Pt(12); r.font.color.rgb = COLOR_WHITE
r = p.add_run(); r.text = "  🌐 서울특별시 · 숭실대학교 벤처중소기업학과"; r.font.name = FONT_BODY; r.font.size = Pt(12); r.font.color.rgb = COLOR_WHITE

# Save Output
out_path = os.path.join(BASE_DIR, "마케터_장상균_포트폴리오.pptx")
prs.save(out_path)
print(f"SUCCESS: Successfully generated PPTX deck at {out_path}")
