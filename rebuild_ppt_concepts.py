# -*- coding: utf-8 -*-
"""
1. 기존 PPT에서 슬라이드 17~22 제대로 삭제
2. 비전공자용 모델 개념 슬라이드 6장 새로 삽입 (슬라이드 10번 위치)
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import copy
from lxml import etree

NSMAP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

prs = Presentation('01_stock_prediction_report.pptx')
print('로드 시 슬라이드 수:', len(prs.slides))

# ──────────────────────────────────────
# STEP 1: 슬라이드 17~22 제대로 삭제
#  → sldIdLst에서 제거 + 관계(rel)도 제거
# ──────────────────────────────────────
def delete_slide(prs, idx):
    sldIdLst = prs.slides._sldIdLst
    sld_id_elems = sldIdLst.findall(
        '{http://schemas.openxmlformats.org/presentationml/2006/main}sldId')
    target = sld_id_elems[idx]
    rId = target.get(f'{{{NSMAP}}}id')
    sldIdLst.remove(target)
    # _rels 내부 딕셔너리에서 직접 제거
    if rId in prs.part._rels._rels:
        del prs.part._rels._rels[rId]

# 뒤에서부터 삭제 (인덱스 밀림 방지)
for i in reversed(range(16, 22)):
    delete_slide(prs, i)

print('삭제 후 슬라이드 수:', len(prs.slides))  # 16이어야 함

# ──────────────────────────────────────
# 공통 헬퍼 함수
# ──────────────────────────────────────
W = prs.slide_width
H = prs.slide_height

def rgb(r, g, b):
    return RGBColor(r, g, b)

COLOR_BG    = rgb(15, 23, 42)
COLOR_WHITE = rgb(255, 255, 255)
COLOR_GRAY  = rgb(148, 163, 184)
COLOR_GREEN = rgb(74, 222, 128)
COLOR_RED   = rgb(248, 113, 113)

CAT_COLORS = {
    'linear':  rgb(234, 168, 23),
    'tree':    rgb(167, 139, 250),
    'knn':     rgb(34, 211, 238),
}

def add_rect(slide, left, top, width, height, fill):
    s = slide.shapes.add_shape(1,
        Inches(left), Inches(top), Inches(width), Inches(height))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s

def add_tb(slide, left, top, width, height,
           text, size=12, bold=False, color=None,
           align=PP_ALIGN.LEFT, wrap=True):
    tb = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color or COLOR_WHITE
    return tb

def set_bg(slide, color):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = color

def make_slide(prs, title, sub, tagline,
               an_title, an_body,
               hw_title, hw_items,
               pros, cons,
               page, accent, badge=''):
    """
    레이아웃:
    ┌─[배지] 제목────────────────────────┐  헤더
    │ 한줄요약                            │  tagline
    ├──────────────────┬─────────────────┤
    │ 비유로 이해하기   │ 어떻게 작동하나  │
    ├──────────────────┴─────────────────┤
    │ 장점 ✅         │ 단점 ⚠            │
    └─────────────────────────────────────┘
    """
    layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(layout)
    # 기본 placeholder 제거
    for ph in slide.placeholders:
        sp = ph._element
        sp.getparent().remove(sp)

    set_bg(slide, COLOR_BG)

    # 헤더 바
    add_rect(slide, 0, 0, 10, 1.1, accent)
    if badge:
        add_rect(slide, 0.18, 0.16, 1.3, 0.42, rgb(15, 23, 42))
        add_tb(slide, 0.18, 0.16, 1.3, 0.42,
               badge, size=9, bold=True, color=accent,
               align=PP_ALIGN.CENTER)
    add_tb(slide, 1.62, 0.08, 7.0, 0.58,
           title, size=21, bold=True,
           color=rgb(15, 23, 42))
    add_tb(slide, 1.62, 0.64, 7.0, 0.36,
           sub, size=10, color=rgb(30, 41, 59))
    add_tb(slide, 8.7, 0.08, 1.1, 0.5,
           page, size=10, color=rgb(30, 41, 59),
           align=PP_ALIGN.RIGHT)

    # 한줄요약 바
    add_rect(slide, 0, 1.1, 10, 0.52, rgb(30, 41, 59))
    add_tb(slide, 0.25, 1.15, 9.5, 0.42,
           f'한 줄 요약   |   {tagline}',
           size=11, color=accent)

    # 왼쪽 패널: 비유
    add_rect(slide, 0.18, 1.75, 4.65, 2.9, rgb(22, 33, 55))
    add_tb(slide, 0.32, 1.82, 4.35, 0.38,
           an_title, size=11, bold=True, color=accent)
    add_tb(slide, 0.32, 2.24, 4.35, 2.32,
           an_body, size=10.5, color=COLOR_WHITE)

    # 오른쪽 패널: 작동방식
    add_rect(slide, 5.17, 1.75, 4.65, 2.9, rgb(22, 33, 55))
    add_tb(slide, 5.32, 1.82, 4.35, 0.38,
           hw_title, size=11, bold=True, color=accent)
    y = 2.24
    for item in hw_items:
        add_tb(slide, 5.32, y, 4.35, 0.4,
               item, size=10, color=COLOR_WHITE)
        y += 0.44

    # 장점
    add_rect(slide, 0.18, 4.78, 4.65, 1.82, rgb(16, 38, 28))
    add_tb(slide, 0.32, 4.85, 2.0, 0.36,
           '✅  장점', size=11, bold=True, color=COLOR_GREEN)
    y2 = 5.24
    for p_text in pros:
        add_tb(slide, 0.32, y2, 4.4, 0.38,
               f'• {p_text}', size=10, color=COLOR_WHITE)
        y2 += 0.4

    # 단점
    add_rect(slide, 5.17, 4.78, 4.65, 1.82, rgb(38, 16, 16))
    add_tb(slide, 5.32, 4.85, 2.0, 0.36,
           '⚠  단점', size=11, bold=True, color=COLOR_RED)
    y3 = 5.24
    for c_text in cons:
        add_tb(slide, 5.32, y3, 4.4, 0.38,
               f'• {c_text}', size=10, color=COLOR_WHITE)
        y3 += 0.4

    # 하단 바
    add_rect(slide, 0, 6.78, 10, 0.22, rgb(22, 33, 55))
    add_tb(slide, 0.2, 6.8, 7.5, 0.18,
           '주식 등락률 예측 AI 모델 개발', size=8, color=COLOR_GRAY)
    add_tb(slide, 8.0, 6.8, 1.8, 0.18,
           page, size=8, color=COLOR_GRAY, align=PP_ALIGN.RIGHT)

    return slide


# ──────────────────────────────────────
# STEP 2: 6개 모델 슬라이드 생성
# (add_slide는 항상 맨 뒤에 추가됨)
# ──────────────────────────────────────

make_slide(
    prs,
    title   = '선형 회귀 (Linear Regression)',
    sub     = '가장 단순한 예측 공식 — 다른 모델 성능의 기준선(Baseline) 역할',
    tagline = '"각 지표에 가중치를 곱해 더하면 등락률이 나온다"',
    an_title = '비유로 이해하기',
    an_body  = (
        '키와 몸무게의 관계처럼,\n'
        '"RSI가 1 오르면 등락률 0.03% 변화"\n'
        '이런 공식을 데이터에서 자동으로 찾습니다.\n\n'
        '수만 개의 과거 날짜를 보면서\n'
        '"어떤 가중치 조합이 정답에 가장\n'
        ' 가까운지"를 계산합니다.\n\n'
        '예:  y = 0.03×RSI + 0.02×MACD\n'
        '       + 0.01×거래량 + ...'
    ),
    hw_title = '어떻게 작동하나',
    hw_items = [
        '① 11개 지표 각각에 가중치를 부여',
        '② 가중치×지표를 모두 더함 = 예측값',
        '③ 오차를 최소화하는 가중치를 학습',
        '④ 가중치를 보면 어떤 지표가',
        '   중요한지 바로 알 수 있음',
    ],
    pros    = [
        '학습 시간 0.1초 미만 — 9개 중 가장 빠름',
        '가중치를 보면 지표별 영향력 즉시 파악',
        '다른 모델의 성능 비교 기준선 역할',
    ],
    cons    = [
        '"RSI 높고 거래량 급증" 같은 조합 효과 미반영',
        '비선형 복잡한 시장 패턴 표현 한계',
        '규제 없으면 과적합 위험 (→ Ridge/Lasso)',
    ],
    page    = '10 / 22',
    accent  = CAT_COLORS['linear'],
    badge   = '선형 계열',
)

make_slide(
    prs,
    title   = 'Ridge 회귀',
    sub     = '선형 회귀 + "너무 극단적인 답을 내지 마" 제약',
    tagline = '"모든 가중치를 조금씩 줄여 안정적인 예측을 만든다"',
    an_title = '비유로 이해하기',
    an_body  = (
        '"답안을 너무 극단적으로 쓰지 마세요"\n'
        '라는 채점 규정을 상상해 보세요.\n\n'
        '선형 회귀는 가중치가 지나치게 클 수 있어요.\n'
        '→ "이 지표가 무조건 1,000% 중요해!"\n\n'
        'Ridge는 가중치가 커질수록 페널티를 주어\n'
        '"모든 지표를 균형 있게 활용"하도록 합니다.\n\n'
        'Alpha(α)가 클수록 가중치를 더 강하게 억제\n'
        '→ 탐색 결과: α = 0.1 선택'
    ),
    hw_title = '어떻게 작동하나',
    hw_items = [
        '① 선형 회귀처럼 가중치 학습',
        '② 가중치² 합산에 페널티 추가',
        '③ Alpha(α)로 페널티 강도 조절',
        '④ 모든 가중치를 0이 아닌',
        '   작은 값으로 유지 (완전 제거 X)',
    ],
    pros    = [
        '과적합 방지 — 선형 회귀보다 안정적',
        '모든 지표를 균형 있게 조금씩 활용',
        '학습 시간 1초 미만',
    ],
    cons    = [
        '불필요한 지표도 완전히 제거하지 않음',
        'Alpha 최적값을 별도 탐색해야 함',
        '비선형 관계 표현 불가',
    ],
    page    = '11 / 22',
    accent  = CAT_COLORS['linear'],
    badge   = '선형 계열',
)

make_slide(
    prs,
    title   = 'Lasso 회귀',
    sub     = '중요하지 않은 지표를 완전히 0으로 만들어 자동 정리',
    tagline = '"불필요한 지표를 자동으로 걸러내 진짜 중요한 것만 남긴다"',
    an_title = '비유로 이해하기',
    an_body  = (
        '11명이 투자 의견을 말하는데,\n'
        '"설득력 없는 의견은 아예 발언권 0으로"\n'
        '만드는 사회자를 상상해 보세요.\n\n'
        'Ridge: 11명 모두 발언, 극단적 주장만 줄임\n'
        'Lasso: 핵심 멤버만 남기고 나머지 침묵\n\n'
        '→ Lasso 결과: 예측에 중요한 지표만 남고\n'
        '   나머지 가중치가 정확히 0이 됩니다.\n\n'
        '→ 탐색 결과: α = 0.001 선택'
    ),
    hw_title = '어떻게 작동하나',
    hw_items = [
        '① 선형 회귀처럼 가중치 학습',
        '② 가중치 절댓값 합산에 페널티 추가',
        '③ 페널티가 강해지면 일부 가중치 → 0',
        '④ 가중치 0인 지표 = 예측에 불필요',
        '   → 자동으로 변수 선택',
    ],
    pros    = [
        '자동 변수 선택 — 중요한 지표가 무엇인지 밝힘',
        '불필요한 지표를 완전 제거해 모델 단순화',
        '과적합 방지 + 해석 가능성 동시 확보',
    ],
    cons    = [
        '상관관계 높은 지표 중 하나를 임의 제거 가능',
        'Alpha가 너무 크면 모든 가중치가 0이 됨',
        '비선형 관계 표현 불가',
    ],
    page    = '12 / 22',
    accent  = CAT_COLORS['linear'],
    badge   = '선형 계열',
)

make_slide(
    prs,
    title   = 'ElasticNet 회귀',
    sub     = 'Ridge + Lasso를 섞은 하이브리드 모델',
    tagline = '"Ridge의 안정성 + Lasso의 변수 선택을 동시에 얻는다"',
    an_title = '비유로 이해하기',
    an_body  = (
        '두 가지 다이어트 방법이 있다고 해요.\n\n'
        'Ridge식: 모든 음식을 조금씩 줄이기\n'
        'Lasso식: 나쁜 음식은 완전히 끊기\n\n'
        'ElasticNet은 "둘을 반반씩" 하는 방식.\n\n'
        'l1_ratio = 0.5\n'
        '→ Ridge 50% + Lasso 50% 동시 적용\n\n'
        '→ 탐색 결과: α=0.001, l1_ratio=0.5 선택'
    ),
    hw_title = '어떻게 작동하나',
    hw_items = [
        '① Ridge 페널티 + Lasso 페널티 동시 적용',
        '② l1_ratio로 두 방식의 비율 조절',
        '③ 일부 계수 → 0 (Lasso 효과)',
        '④ 나머지 계수는 고르게 억제 (Ridge 효과)',
        '⑤ 두 하이퍼파라미터(α, l1_ratio) 탐색',
    ],
    pros    = [
        'Ridge + Lasso 장점 동시에 누림',
        '지표 간 상관관계가 높을 때 특히 효과적',
        '가장 유연한 선형 모델',
    ],
    cons    = [
        '탐색해야 할 하이퍼파라미터가 2개',
        'Ridge나 Lasso보다 해석이 약간 복잡',
        '비선형 관계 표현 불가',
    ],
    page    = '13 / 22',
    accent  = CAT_COLORS['linear'],
    badge   = '선형 계열',
)

make_slide(
    prs,
    title   = '결정 트리 (Decision Tree)',
    sub     = '조건 분기를 반복해 예측하는 규칙 기반 모델',
    tagline = '"RSI가 낮고 거래량이 높으면 상승" — 사람이 읽을 수 있는 규칙',
    an_title = '비유로 이해하기',
    an_body  = (
        '병원 응급실 환자 분류를 떠올려 보세요.\n\n'
        '"체온 > 38도?"\n'
        '  Yes → "혈압 정상?"\n'
        '          No → "즉시 입원"\n\n'
        '주식도 마찬가지:\n'
        '"RSI < 30?" (과매도)\n'
        '  Yes → "거래량 급증?"\n'
        '          Yes → "상승 +2.3% 예측"\n\n'
        'max_depth=8: 분기를 8번으로 제한\n'
        '→ 너무 깊어지면 훈련 데이터만 외움'
    ),
    hw_title = '어떻게 작동하나',
    hw_items = [
        '① 어떤 지표를 어떤 기준으로 나눌지 학습',
        '② 나눌 때마다 오차가 가장 줄어드는 분기',
        '③ max_depth=8 에서 분기 중단',
        '④ 마지막 그룹(Leaf)의 평균 = 예측값',
        '⑤ 규칙을 출력하면 사람이 직접 읽기 가능',
    ],
    pros    = [
        '예측 규칙을 직접 읽을 수 있어 해석성 최고',
        '데이터 전처리(정규화 등) 불필요',
        '비선형 관계 및 지표 간 조합 포착 가능',
    ],
    cons    = [
        '깊어질수록 훈련 데이터를 외워버림 (과적합)',
        '9개 모델 중 2025년 실전 과적합 갭 가장 큼 (+0.7%p)',
        'RF/XGB보다 성능 낮음 (트리 1개의 한계)',
    ],
    page    = '14 / 22',
    accent  = CAT_COLORS['tree'],
    badge   = '트리 계열',
)

make_slide(
    prs,
    title   = 'KNN 회귀 (K-Nearest Neighbors)',
    sub     = '"오늘과 가장 비슷했던 과거 날들의 결과 평균" = 예측',
    tagline = '"오늘 날씨 같은 날 다음엔 어땠지?" — 유사한 과거 k개를 참조',
    an_title = '비유로 이해하기',
    an_body  = (
        '동네 부동산 가격 예측을 생각해 보세요.\n\n'
        '"이 집과 비슷한 집 20채를 찾아보자.\n'
        ' 그 집들의 실거래가 평균이 이 집 가격."\n\n'
        '주식도 마찬가지:\n'
        '오늘: RSI=42, 거래량=1.2배, MACD=0.3...\n'
        '→ 과거에 이런 상황이었던 날 20개 검색\n'
        '→ 그 날들의 5일 후 등락률 평균 = 예측\n\n'
        'k=20: 가장 비슷한 20개 날 참조\n'
        '(탐색에서 k=20일 때 MAE 최소)'
    ),
    hw_title = '어떻게 작동하나',
    hw_items = [
        '① 학습 = 훈련 데이터를 통째로 저장',
        '② 예측 = 오늘과 거리 가장 가까운',
        '   k=20개 날 탐색 (11개 지표 기준)',
        '③ 20개 날의 5일 후 등락률 평균 = 예측',
        '④ 데이터 많을수록 참조 후보 증가',
    ],
    pros    = [
        '학습 과정 없음 — 즉시 사용 가능',
        '"비슷한 과거 = 비슷한 미래" 직관적 원리',
        '데이터 추가만 하면 자동으로 모델 개선',
    ],
    cons    = [
        '예측마다 전체 훈련 데이터 탐색 → 느림',
        '11개 지표가 모두 동등하게 취급됨',
        '지표 수가 늘면 거리 의미 희석 (차원의 저주)',
    ],
    page    = '15 / 22',
    accent  = CAT_COLORS['knn'],
    badge   = '거리 기반',
)

print('슬라이드 생성 후 총 수:', len(prs.slides))  # 22

# ──────────────────────────────────────
# STEP 3: 새 슬라이드(16~21)를 슬라이드 9 뒤로 이동
# 목표 순서: 0~8 (기존), 16~21 (새 6장), 9~15 (기존 나머지)
# ──────────────────────────────────────
NS_PML = 'http://schemas.openxmlformats.org/presentationml/2006/main'
sldIdLst = prs.slides._sldIdLst
all_ids  = sldIdLst.findall(f'{{{NS_PML}}}sldId')

old_first9  = all_ids[:9]
new_six     = all_ids[16:22]
old_rest    = all_ids[9:16]

for elem in all_ids:
    sldIdLst.remove(elem)
for elem in old_first9 + new_six + old_rest:
    sldIdLst.append(elem)

print('순서 조정 후 총 슬라이드:', len(prs.slides))

# ──────────────────────────────────────
# STEP 4: 저장
# ──────────────────────────────────────
prs.save('01_stock_prediction_report.pptx')
print('PPT saved OK')
