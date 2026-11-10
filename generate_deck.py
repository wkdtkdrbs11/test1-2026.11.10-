# -*- coding: utf-8 -*-
"""
Generate Marketer Jang Sanggyun's Portfolio PPTX
Based on Anthropic pptx skill design guidelines:
- 16:9 Widescreen (13.333" x 7.5")
- Toss Blue (#2972E8) & Slate Gray Palette
- High contrast "Sandwich" structure (Dark cover & closing, light content)
- Embedding actual project images (profile.png, smartbin.png, cherrycoffee.jpg, ai-mvp-poster.jpg)
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

# Color Palette
COLOR_DARK_BG = RGBColor(28, 31, 37)      # #1C1F25
COLOR_LIGHT_BG = RGBColor(246, 247, 249)  # #F6F7F9
COLOR_CARD_BG = RGBColor(255, 255, 255)   # #FFFFFF
COLOR_PRIMARY = RGBColor(41, 114, 232)    # #2972E8 (Toss Blue)
COLOR_TEXT_MAIN = RGBColor(28, 31, 37)    # #1C1F25
COLOR_TEXT_MUTED = RGBColor(110, 122, 147)# #6E7A93
COLOR_TEXT_WHITE = RGBColor(246, 247, 249)# #F6F7F9
COLOR_BORDER = RGBColor(225, 228, 235)    # #E1E4EB
COLOR_ACCENT_BG = RGBColor(235, 242, 255) # #EBF2FF

BASE_DIR = r"C:\Users\ABC\.gemini\antigravity\scratch\portfolio-profile"
IMG_PROFILE = os.path.join(BASE_DIR, "assets", "images", "profile.png")
IMG_SMARTBIN = os.path.join(BASE_DIR, "assets", "images", "smartbin.png")
IMG_CHERRY = os.path.join(BASE_DIR, "assets", "images", "cherrycoffee.jpg")
IMG_POSTER = os.path.join(BASE_DIR, "assets", "images", "ai-mvp-poster.jpg")
OUTPUT_PPTX = os.path.join(BASE_DIR, "마케터_장상균_포트폴리오.pptx")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

def add_bg(slide, color):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    return bg

def add_card(slide, left, top, width, height, fill_color=COLOR_CARD_BG, border_color=COLOR_BORDER):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = fill_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(1)
    else:
        card.line.fill.background()
    return card

def add_header(slide, tag_text, title_text, is_dark=False):
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(11.7), Inches(1.1))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p0 = tf.paragraphs[0]
    p0.text = tag_text
    p0.font.name = "Pretendard"
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = COLOR_PRIMARY if is_dark else COLOR_PRIMARY
    p0.space_after = Pt(4)
    
    p1 = tf.add_paragraph()
    p1.text = title_text
    p1.font.name = "Pretendard"
    p1.font.size = Pt(22)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_TEXT_WHITE if is_dark else COLOR_TEXT_MAIN

# ==========================================
# SLIDE 1: Cover (Dark)
# ==========================================
s1 = prs.slides.add_slide(blank_layout)
add_bg(s1, COLOR_DARK_BG)

# Accent badge
badge = add_card(s1, Inches(0.8), Inches(1.5), Inches(3.2), Inches(0.42), fill_color=RGBColor(38, 45, 60), border_color=None)
tf_badge = badge.text_frame
tf_badge.word_wrap = False
p_badge = tf_badge.paragraphs[0]
p_badge.text = "● BRAND & CRM MARKETER"
p_badge.font.name = "JetBrains Mono"
p_badge.font.size = Pt(11)
p_badge.font.bold = True
p_badge.font.color.rgb = COLOR_PRIMARY
p_badge.alignment = PP_ALIGN.CENTER

# Title & Subtitle
tb_title = s1.shapes.add_textbox(Inches(0.8), Inches(2.1), Inches(7.5), Inches(2.8))
tf = tb_title.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

p1 = tf.paragraphs[0]
p1.text = "브랜드의 본질을 파악하고,\n데이터로 입증하는 마케터"
p1.font.name = "Pretendard"
p1.font.size = Pt(36)
p1.font.bold = True
p1.font.color.rgb = COLOR_TEXT_WHITE
p1.space_after = Pt(16)

p2 = tf.add_paragraph()
p2.text = "고객이 선택하는 진짜 이유와 불편을 파고들어 가설을 세우고,\n현장에서 발로 뛰는 실험과 데이터로 비즈니스 성과를 만듭니다."
p2.font.name = "Pretendard"
p2.font.size = Pt(15)
p2.font.color.rgb = RGBColor(160, 170, 190)

# Profile Info Card
tb_meta = s1.shapes.add_textbox(Inches(0.8), Inches(5.4), Inches(7.5), Inches(1.4))
tf_m = tb_meta.text_frame
tf_m.margin_left = tf_m.margin_top = tf_m.margin_right = tf_m.margin_bottom = 0
p_name = tf_m.paragraphs[0]
p_name.text = "장상균  |  Jang Sanggyun"
p_name.font.name = "Pretendard"
p_name.font.size = Pt(18)
p_name.font.bold = True
p_name.font.color.rgb = COLOR_TEXT_WHITE
p_name.space_after = Pt(6)

p_info = tf_m.add_paragraph()
p_info.text = "숭실대학교 벤처중소기업학과  •  wkdtkdrbs11@naver.com"
p_info.font.name = "Pretendard"
p_info.font.size = Pt(13)
p_info.font.color.rgb = COLOR_TEXT_MUTED

# Profile Image (Right)
if os.path.exists(IMG_PROFILE):
    s1.shapes.add_picture(IMG_PROFILE, Inches(9.2), Inches(1.6), Inches(3.3), Inches(4.3))

# ==========================================
# SLIDE 2: Key Metrics & Philosophy (Light)
# ==========================================
s2 = prs.slides.add_slide(blank_layout)
add_bg(s2, COLOR_LIGHT_BG)
add_header(s2, "KEY METRICS & PHILOSOPHY", "숫자로 증명된 실행력과 3가지 핵심 태도")

# 4 Stat Cards
metrics_data = [
    ("+25%", "매출 성장 견인", "체리커피 릴스 마케팅 1개월 만에 달성"),
    ("200명+", "오가닉 팔로워 증가", "상권 차별화 강점 발굴 및 릴스 제작"),
    ("45건", "현장 인터뷰", "선별장 및 대형 축제 이해관계자 심층 리서치"),
    ("7.7%", "MVP 가입 전환율", "오프라인 포스터 가설 검증 (39명 중 3명 가입)")
]
card_w = Inches(2.78)
card_h = Inches(1.7)
start_x = Inches(0.8)
start_y = Inches(1.8)
gap_x = Inches(0.2)

for i, (num, label, desc) in enumerate(metrics_data):
    cx = start_x + i * (card_w + gap_x)
    add_card(s2, cx, start_y, card_w, card_h)
    tb = s2.shapes.add_textbox(cx + Inches(0.2), start_y + Inches(0.2), card_w - Inches(0.4), card_h - Inches(0.4))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p = tf.paragraphs[0]
    p.text = num
    p.font.name = "Pretendard"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY
    p.space_after = Pt(4)
    
    p_lbl = tf.add_paragraph()
    p_lbl.text = label
    p_lbl.font.name = "Pretendard"
    p_lbl.font.size = Pt(13)
    p_lbl.font.bold = True
    p_lbl.font.color.rgb = COLOR_TEXT_MAIN
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "Pretendard"
    p_d.font.size = Pt(10.5)
    p_d.font.color.rgb = COLOR_TEXT_MUTED

# 3 Philosophy Cards
phil_data = [
    ("1. 문제의 본질 정의", "눈앞에 보이는 현상에 머무르지 않고, 고객이 선택하는 이유와 불편을 느끼는 지점을 구체적으로 파고듭니다.", "단순 인테리어 홍보 대신 '역세권 도심 속 교외 정원'으로 본질 가치 재정의"),
    ("2. 데이터와 실험으로 검증", "막연한 추측 대신 직접 가설을 세우고, 빠른 실험을 통해 숫자로 명확히 증명합니다.", "개발 전 포스터 MVP 실험으로 고객 반응(전환율 7.7%) 먼저 검증 후 실행"),
    ("3. 발로 뛰는 현장 협업", "모두가 꺼리는 현장이라도 꼭 필요한 정보라면 직접 발로 뛰어 확인하고 팀과 투명하게 나눕니다.", "쓰레기 선별장을 직접 방문해 인터뷰하며 플라스틱 혼입 문제를 발견하고 피벗 주도")
]
pcard_w = Inches(3.77)
pcard_h = Inches(2.9)
pstart_y = Inches(3.8)

for i, (p_title, p_desc, p_foot) in enumerate(phil_data):
    cx = start_x + i * (pcard_w + Inches(0.2))
    add_card(s2, cx, pstart_y, pcard_w, pcard_h)
    tb = s2.shapes.add_textbox(cx + Inches(0.25), pstart_y + Inches(0.25), pcard_w - Inches(0.5), pcard_h - Inches(0.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p = tf.paragraphs[0]
    p.text = p_title
    p.font.name = "Pretendard"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY
    p.space_after = Pt(10)
    
    p2 = tf.add_paragraph()
    p2.text = f'"{p_desc}"'
    p2.font.name = "Pretendard"
    p2.font.size = Pt(12)
    p2.font.color.rgb = COLOR_TEXT_MAIN
    p2.space_after = Pt(14)
    
    p3 = tf.add_paragraph()
    p3.text = f"실행 사례:\n{p_foot}"
    p3.font.name = "Pretendard"
    p3.font.size = Pt(11)
    p3.font.color.rgb = COLOR_TEXT_MUTED

# ==========================================
# SLIDE 3: Project 1 - Smartbin
# ==========================================
s3 = prs.slides.add_slide(blank_layout)
add_bg(s3, COLOR_LIGHT_BG)
add_header(s3, "PROJECT 01  •  ESG & CIRCULAR ECONOMY", "가상 브랜드 ‘스마트빈’ 기획 및 비즈니스 모델 구축")

# Image Box (Left)
if os.path.exists(IMG_SMARTBIN):
    s3.shapes.add_picture(IMG_SMARTBIN, Inches(0.8), Inches(1.8), Inches(4.3), Inches(4.9))

# Content Card (Right)
c_right = add_card(s3, Inches(5.3), Inches(1.8), Inches(7.2), Inches(4.9))
tb = s3.shapes.add_textbox(Inches(5.6), Inches(2.0), Inches(6.6), Inches(4.5))
tf = tb.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

# Tag
p_tag = tf.paragraphs[0]
p_tag.text = "KDB 산업은행 창업 프로젝트 장려상 수상  |  3인 팀 (기여도 40%)"
p_tag.font.name = "Pretendard"
p_tag.font.size = Pt(11)
p_tag.font.bold = True
p_tag.font.color.rgb = COLOR_PRIMARY
p_tag.space_after = Pt(10)

def add_section_block(tf, label, content):
    p_lbl = tf.add_paragraph()
    p_lbl.text = label
    p_lbl.font.name = "Pretendard"
    p_lbl.font.size = Pt(12.5)
    p_lbl.font.bold = True
    p_lbl.font.color.rgb = COLOR_TEXT_MAIN
    p_lbl.space_after = Pt(2)
    
    p_txt = tf.add_paragraph()
    p_txt.text = content
    p_txt.font.name = "Pretendard"
    p_txt.font.size = Pt(11.5)
    p_txt.font.color.rgb = COLOR_TEXT_MUTED
    p_txt.space_after = Pt(10)

add_section_block(tf, "1. 배경 및 문제 정의 (Problem)", 
    "대형 축제 및 무인카페의 일회용 플라스틱 컵은 약 90%가 빈 공기 상태로 버려져 부피가 폭증하고 잔여 음료 부패로 실질 재활용률이 16.4%에 불과하며 주최 측 청소 비용이 과다 발생함.")

add_section_block(tf, "2. 현장 리서치 및 핵심 인사이트 (Insight)",
    "단순 압축 쓰레기통 구상에서 벗어나, 직접 쓰레기 선별장 45건 인터뷰를 진행. 플라스틱 재질(PET/PP) 혼합 압축 시 선별 불가로 가치 0원 및 오염 침투 문제를 발견하여 사업 방향을 피벗함.")

add_section_block(tf, "3. 솔루션 (Solution): 선별장 Skip형 고순도 r-PP 직송 밸류체인",
    "• 단일 재질 통제: 무지 투명 PP컵 3종 일괄 사전 공급으로 혼입 원천 차단\n• 스마트빈 인프라: 음료 분리 및 발 페달식 물리 압축으로 부피 70% 절감\n• 턴키 직수거: 행사 후 전량 자체 회수하여 재활용 공장에 고순도 펠릿 직공급")

p_res_title = tf.add_paragraph()
p_res_title.text = "4. 주요 성과 (Results)"
p_res_title.font.name = "Pretendard"
p_res_title.font.size = Pt(12.5)
p_res_title.font.bold = True
p_res_title.font.color.rgb = COLOR_TEXT_MAIN

p_res = tf.add_paragraph()
p_res.text = "✔ KDB 산업은행 창업 프로젝트 최종 IR 피칭 및 장려상 수상\n✔ 대형 축제 단기 집약 거점 모델 및 순환식 '밀크런(Milk Run)' 물류 전략 수립"
p_res.font.name = "Pretendard"
p_res.font.size = Pt(11.5)
p_res.font.bold = True
p_res.font.color.rgb = COLOR_PRIMARY

# ==========================================
# SLIDE 4: Project 2 - Cherry Coffee
# ==========================================
s4 = prs.slides.add_slide(blank_layout)
add_bg(s4, COLOR_LIGHT_BG)
add_header(s4, "PROJECT 02  •  F&B BRAND & LOCAL MARKETING", "카페 ‘체리커피 군자점’ 지점 운영 총괄 및 로컬 마케팅")

# Image Box (Left)
if os.path.exists(IMG_CHERRY):
    s4.shapes.add_picture(IMG_CHERRY, Inches(0.8), Inches(1.8), Inches(4.3), Inches(4.9))

# Content Card (Right)
add_card(s4, Inches(5.3), Inches(1.8), Inches(7.2), Inches(4.9))
tb = s4.shapes.add_textbox(Inches(5.6), Inches(2.0), Inches(6.6), Inches(4.5))
tf = tb.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

p_tag = tf.paragraphs[0]
p_tag.text = "1개월 만에 매출 25% 상승 견인  |  1인 단독 총괄 (기여도 100%)"
p_tag.font.name = "Pretendard"
p_tag.font.size = Pt(11)
p_tag.font.bold = True
p_tag.font.color.rgb = COLOR_PRIMARY
p_tag.space_after = Pt(10)

add_section_block(tf, "1. 배경 및 과제 (Problem)",
    "군자역 인근 다수 프랜차이즈/개인 카페 사이에서 차별점 부재. 기존 SNS 홍보는 단순 인테리어 사진 위주였으며 피크 타임 외 유휴 시간대 유입 및 지역 단골 확보가 절실했음.")

add_section_block(tf, "2. 핵심 강점 발굴 및 인사이트 (Insight)",
    "단순한 예쁜 인테리어가 아니라 '지하철역 도보 3분 역세권이면서도 도심 속에서 교외 대형 정원 카페의 여유로운 자연 채광과 테라스를 누릴 수 있다'는 공간적 가치를 핵심 강점으로 도출.")

add_section_block(tf, "3. 3대 실행 전략 (Solution)",
    "• 공간 강점 기반 릴스 제작: 역 출구부터 정원 테라스로 이어지는 접근성 강조 숏폼 배포\n• 매장 경험(CX) 넛지 적용: 주문 동선에 당일 구운 시그니처 베이커리를 전략 배치해 객단가 상승\n• 로컬 B2B 상생 제휴: 인근 인기 식당 5곳과 상호 쿠폰 비치로 점심시간 유입 극대화")

p_res_title = tf.add_paragraph()
p_res_title.text = "4. 주요 성과 (Results)"
p_res_title.font.name = "Pretendard"
p_res_title.font.size = Pt(12.5)
p_res_title.font.bold = True
p_res_title.font.color.rgb = COLOR_TEXT_MAIN

p_res = tf.add_paragraph()
p_res.text = "✔ 인스타그램 릴스 런칭 후 1개월 만에 오가닉 팔로워 200명 증가\n✔ 릴스 바이럴 및 로컬 제휴 연계를 통해 전월 대비 매장 매출 25% 상승\n✔ 인근 식당 5곳과의 지속 가능한 B2B 제휴 생태계 구축"
p_res.font.name = "Pretendard"
p_res.font.size = Pt(11.5)
p_res.font.bold = True
p_res.font.color.rgb = COLOR_PRIMARY

# ==========================================
# SLIDE 5: Project 3 - AI MVP Sprint (Actual Poster!)
# ==========================================
s5 = prs.slides.add_slide(blank_layout)
add_bg(s5, COLOR_LIGHT_BG)
add_header(s5, "PROJECT 03  •  PROBLEM DEFINITION & LEAN MVP", "대학생 AI 교차검증 문제 정의 및 포스터 MVP 가설 검증")

# Poster Image (Left)
if os.path.exists(IMG_POSTER):
    s5.shapes.add_picture(IMG_POSTER, Inches(0.8), Inches(1.8), Inches(4.3), Inches(4.9))

# Content Card (Right)
add_card(s5, Inches(5.3), Inches(1.8), Inches(7.2), Inches(4.9))
tb = s5.shapes.add_textbox(Inches(5.6), Inches(2.0), Inches(6.6), Inches(4.5))
tf = tb.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

p_tag = tf.paragraphs[0]
p_tag.text = "오프라인 포스터 가설 검증 전환율 7.7% 달성  |  숭실지부장 총괄"
p_tag.font.name = "Pretendard"
p_tag.font.size = Pt(11)
p_tag.font.bold = True
p_tag.font.color.rgb = COLOR_PRIMARY
p_tag.space_after = Pt(10)

add_section_block(tf, "1. 배경 및 문제 정의 (Problem)",
    "초기에는 'AI 답변의 신뢰성이 낮아 여러 AI를 교차검증하느라 시간을 낭비한다'고 보았으나, 멘토로부터 '신뢰란 무엇인가'라는 피드백을 받고 문제를 더 날카롭게 쪼개기로 결정함.")

add_section_block(tf, "2. 문제 쪼개기 및 린 의사결정 (Insight)",
    "• 신뢰와 시간 낭비 분해: '신뢰'는 답변의 정확성, '불필요한 시간'은 정확성을 재확인하기 위해 여러 툴을 오가는 교차검증 시간으로 구체화\n• 실행 병목 극복: 실험 준비가 한 달 넘게 지연되자 완성도를 높이기보다 실제 고객 반응부터 먼저 검증하자고 팀을 설득")

add_section_block(tf, "3. 포스터 기반 오프라인 MVP 실험 (Solution)",
    "• '교차검증 시간 단축' 메시지와 QR코드를 담은 오프라인 포스터 기획·제작\n• 교수님들의 협조를 확보하여 강의실 및 학과 게시판 주요 동선에 포스터 배포\n• 단순 조회가 아닌 '이름/연락처 입력 및 사전 가입'이라는 높은 허들로 진성 수요 측정")

p_res_title = tf.add_paragraph()
p_res_title.text = "4. 주요 성과 및 배움 (Results)"
p_res_title.font.name = "Pretendard"
p_res_title.font.size = Pt(12.5)
p_res_title.font.bold = True
p_res_title.font.color.rgb = COLOR_TEXT_MAIN

p_res = tf.add_paragraph()
p_res.text = "✔ 포스터 유입 39명 중 3명이 개인정보를 입력하고 가입 (전환율 7.7% 달성)\n✔ 막연한 개발 준비를 멈추고 실제 고객 데이터로 다음 단계 판단 근거 마련\n✔ '문제를 깊이 정의하되, 적시에 가볍게 검증하여 판단 근거를 만드는 실행 기준' 체득"
p_res.font.name = "Pretendard"
p_res.font.size = Pt(11.5)
p_res.font.bold = True
p_res.font.color.rgb = COLOR_PRIMARY

# ==========================================
# SLIDE 6: Professional Stories (3가지 이야기)
# ==========================================
s6 = prs.slides.add_slide(blank_layout)
add_bg(s6, COLOR_LIGHT_BG)
add_header(s6, "PROFESSIONAL STORY", "마케터 장상균의 3가지 이야기 (일하는 방식과 태도)")

stories = [
    ("01. F&B와 마케팅에 빠져들게 된 계기",
     "디테일이 만드는 매출의 변화",
     "1년간 카페를 총괄 운영하며 마케팅의 작은 디테일 하나가 실제 매장의 매출 곡선을 바꾼다는 사실을 체감하고 마케팅의 매력에 깊이 빠져들었습니다.\n\n단순히 음료를 파는 데 머물지 않고 '고객이 수많은 카페 중 왜 우리 매장을 찾아와야 하는가?'를 매일 치열하게 고민했고, 베이커리 넛지 효과 적용 및 동선 최적화, 역세권 대형 정원 릴스 제작으로 1달 만에 팔로워 200명 증가와 매출 25% 상승을 달성했습니다."),
    
    ("02. 본질을 파고드는 집요함과 가설 검증",
     "행동의 진짜 이유를 규명하는 실행력",
     "제 가장 큰 강점은 겉으로 드러난 현상에 머물지 않고, 고객이 그렇게 행동하는 진짜 이유를 끝까지 파고드는 것입니다.\n\n창업 프로젝트 당시 'AI 답변 신뢰 부족' 피드백을 '정확성 재검증에 소모되는 시간 비용'으로 뾰족하게 재정의했습니다. 준비가 길어지자 탁상공론을 멈추고 포스터 MVP 실험을 제안해 39명 중 3명의 가입(전환율 7.7%)을 입증하며 팀의 불확실성을 해소했습니다."),
    
    ("03. 남들이 꺼리는 현장에 뛰어드는 협업",
     "발로 뛰어 얻은 날것의 정보 공유",
     "협업할 때 서로의 상황을 먼저 배려하고, 각자 현장에서 발로 뛰며 얻은 생생한 정보를 투명하게 나눌 때 가장 올바른 결정이 나온다고 믿습니다.\n\n스마트빈 프로젝트 당시 열악한 재활용 선별장 방문 조사는 누구나 꺼리는 일이었지만, 직접 찾아가 인터뷰하며 '재질 혼합 시 가치 0원'이라는 치명적 병목을 발견했습니다. 이를 공유해 '순환경제 직송 밸류체인'으로 사업 피벗을 이끌었습니다.")
]

card_w = Inches(3.77)
card_h = Inches(4.9)
start_x = Inches(0.8)
start_y = Inches(1.8)

for i, (title, sub, body) in enumerate(stories):
    cx = start_x + i * (card_w + Inches(0.2))
    add_card(s6, cx, start_y, card_w, card_h)
    tb = s6.shapes.add_textbox(cx + Inches(0.28), start_y + Inches(0.28), card_w - Inches(0.56), card_h - Inches(0.56))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = "Pretendard"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY
    p.space_after = Pt(4)
    
    p_sub = tf.add_paragraph()
    p_sub.text = sub
    p_sub.font.name = "Pretendard"
    p_sub.font.size = Pt(12)
    p_sub.font.bold = True
    p_sub.font.color.rgb = COLOR_TEXT_MAIN
    p_sub.space_after = Pt(12)
    
    p_body = tf.add_paragraph()
    p_body.text = body
    p_body.font.name = "Pretendard"
    p_body.font.size = Pt(11)
    p_body.font.color.rgb = COLOR_TEXT_MUTED

# ==========================================
# SLIDE 7: Timeline & Skillset
# ==========================================
s7 = prs.slides.add_slide(blank_layout)
add_bg(s7, COLOR_LIGHT_BG)
add_header(s7, "TIMELINE & SKILLSET", "주요 이력, 활동 및 실무 역량 요약")

# Left: Timeline Card
add_card(s7, Inches(0.8), Inches(1.8), Inches(5.7), Inches(4.9))
tb_t = s7.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(5.1), Inches(4.5))
tf_t = tb_t.text_frame
tf_t.word_wrap = True
tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0

p = tf_t.paragraphs[0]
p.text = "주요 활동 및 이력"
p.font.name = "Pretendard"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_PRIMARY
p.space_after = Pt(12)

timeline_items = [
    ("2026. 07 ~ 현재", "대학연합 창업동아리 | 숭실지부장", "문제 정의 스프린트 총괄, 린 가설 검증 실험 리딩"),
    ("2026. 03 ~ 2026. 07", "KDB 산업은행 창업 프로젝트 | IR 피칭 (장려상)", "가상 브랜드 '스마트빈' 순환경제 비즈니스 모델 기획"),
    ("2023. 07 ~ 2024. 06", "카페 체리커피 군자점 | 지점 운영 총괄 (1년)", "매출 25% 상승, 오가닉 팔로워 200명 증가, B2B 제휴"),
    ("2021. 03 ~ 2027. 02", "숭실대학교 벤처중소기업학과 (졸업예정)", "비즈니스 모델링, 스타트업 마케팅 및 기업가정신 전공"),
    ("병역 사항", "39보병사단 대테러특수임무대 3팀 (SDT)", "병장 만기전역 (강한 체력, 위기 대응력, 책임감 배양)")
]

for date, title, desc in timeline_items:
    p_d = tf_t.add_paragraph()
    p_d.text = f"• {title} ({date})"
    p_d.font.name = "Pretendard"
    p_d.font.size = Pt(12)
    p_d.font.bold = True
    p_d.font.color.rgb = COLOR_TEXT_MAIN
    
    p_c = tf_t.add_paragraph()
    p_c.text = f"   {desc}"
    p_c.font.name = "Pretendard"
    p_c.font.size = Pt(10.5)
    p_c.font.color.rgb = COLOR_TEXT_MUTED
    p_c.space_after = Pt(6)

# Right: Skills Card
add_card(s7, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.9))
tb_s = s7.shapes.add_textbox(Inches(7.1), Inches(2.0), Inches(5.1), Inches(4.5))
tf_s = tb_s.text_frame
tf_s.word_wrap = True
tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0

p = tf_s.paragraphs[0]
p.text = "보유 역량 & 실무 도구"
p.font.name = "Pretendard"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_PRIMARY
p.space_after = Pt(12)

skills_data = [
    ("AI & 생산성", "ChatGPT / Claude (5.0/5.0) - 시장 리서치, 카피라이팅, 데이터 구조화\nNotion (4.0/5.0) - 프로젝트 및 관계형 데이터베이스 관리\nGoogle Sheets / Excel (4.0/5.0) - 매출 분석, 함수 및 피벗테이블"),
    ("디자인 & 분석", "Figma (3.0/5.0) - 와이어프레임, SNS 배너 및 포스터 기획\nPhotoshop (2.0/5.0) - 이미지 보정 및 비주얼 애셋 가공\nGA4 (1.5/5.0) - 웹 유입 경로 분석 및 전환 이벤트 트래킹"),
    ("어학 & 기타 자격", "TOEIC 825점 (비즈니스 영어 커뮤니케이션 가능)\n자동차운전면허 1종 보통 (오프라인 현장 기동 및 물류 직접 운용)")
]

for cat, details in skills_data:
    p_c = tf_s.add_paragraph()
    p_c.text = f"■ {cat}"
    p_c.font.name = "Pretendard"
    p_c.font.size = Pt(12)
    p_c.font.bold = True
    p_c.font.color.rgb = COLOR_TEXT_MAIN
    
    p_d = tf_s.add_paragraph()
    p_d.text = details
    p_d.font.name = "Pretendard"
    p_d.font.size = Pt(10.5)
    p_d.font.color.rgb = COLOR_TEXT_MUTED
    p_d.space_after = Pt(10)

# ==========================================
# SLIDE 8: Closing & Contact (Dark)
# ==========================================
s8 = prs.slides.add_slide(blank_layout)
add_bg(s8, COLOR_DARK_BG)

tb_c = s8.shapes.add_textbox(Inches(1.5), Inches(2.0), Inches(10.3), Inches(3.5))
tf_c = tb_c.text_frame
tf_c.word_wrap = True
tf_c.margin_left = tf_c.margin_top = tf_c.margin_right = tf_c.margin_bottom = 0

p_c1 = tf_c.paragraphs[0]
p_c1.text = "함께 더 큰 성장을 만들 준비가 되었습니다"
p_c1.font.name = "Pretendard"
p_c1.font.size = Pt(32)
p_c1.font.bold = True
p_c1.font.color.rgb = COLOR_TEXT_WHITE
p_c1.space_after = Pt(14)

p_c2 = tf_c.add_paragraph()
p_c2.text = "브랜드 마케팅 및 CRM 직무 기회, 커피챗이나 포트폴리오에 관한\n어떤 질문이든 언제든 편하게 연락해 주세요."
p_c2.font.name = "Pretendard"
p_c2.font.size = Pt(16)
p_c2.font.color.rgb = RGBColor(160, 170, 190)
p_c2.space_after = Pt(36)

p_c3 = tf_c.add_paragraph()
p_c3.text = "Email: wkdtkdrbs11@naver.com   |   Phone: 010-5779-****   |   Location: 서울시 동작구 (숭실대학교)"
p_c3.font.name = "Pretendard"
p_c3.font.size = Pt(14)
p_c3.font.bold = True
p_c3.font.color.rgb = COLOR_PRIMARY

# Save Presentation
prs.save(OUTPUT_PPTX)
print(f"Presentation saved successfully to {OUTPUT_PPTX}")
