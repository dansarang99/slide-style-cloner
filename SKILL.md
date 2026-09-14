---
name: slide-style-cloner
description: "Reverse engineer and clone executive presentation styles from any input PDF or PPTX slide deck, extract DESIGN.md, STYLE.md, and PROMPT.md, and generate brand-new, high-fidelity executive presentation slides matching that exact aesthetic. Use whenever the user provides a reference PDF or PPTX deck, says '이 슬라이드 스타일로 만들어줘', '슬라이드 스타일 추출해줘', 'DESIGN.md STYLE.md PROMPT.md 추출', '슬라이드 디자인 복제', or wants to produce McKinsey/BCG-caliber executive strategy decks on a new topic."
---

# Slide Style Cloner (슬라이드 스타일 복제 및 제작 스킬)

PDF 또는 PPTX 슬라이드로부터 최고급 디자인 언어를 역공학(Reverse Engineering)하여 디자인 시스템 문서(`DESIGN.md`, `STYLE.md`, `PROMPT.md`)를 도출하고, 이를 바탕으로 완전히 새로운 주제의 고품질 슬라이드 덱을 복제·생성하는 전문 스킬입니다.

---

## 1. 워크플로우 개요 (Execution Workflow)

```
[입력: 레퍼런스 PDF / PPTX]
       │
       ▼
Phase 1: 비주얼 & 메타데이터 디컴포지션 (Visual Decomposition)
  - `scripts/extract_style.py` 실행하여 캔버스, 폰트, 색상, 레이아웃 좌표 추출
  - 핵심 슬라이드를 고해상도 이미지로 변환하여 시각적 그리드 분석
       │
       ▼
Phase 2: 3대 마스터 가이드 생성
  - templates/를 참조하여 [002]_DESIGN.md, [003]_STYLE.md, [004]_PROMPT.md 작성
       │
       ▼
Phase 3: 사용자 검토 및 신규 주제 기획 (Review & Topic Plan)
  - 추출된 3대 문서를 사용자에게 브리핑하고, 신규 슬라이드 주제 및 목차 확정
       │
       ▼
Phase 4: 신규 슬라이드 생성 및 렌더링 (Generation & Rendering)
  - PROMPT.md를 기반으로 6대 아키타입 슬라이드 코드 생성
  - `scripts/render_deck.py`를 통해 1920x1080 고해상도 렌더링
  - 원본과 1:1 시각적 일치성(Sync) 검증
```

---

## 2. 6대 핵심 슬라이드 아키타입 (Slide Archetypes)

어떤 프레젠테이션을 분석하더라도 슬라이드는 다음 6가지 핵심 패턴으로 분류하여 표준화합니다:

1. **Cover (표지)**:
   - 배경 틴트 + 대형 오가닉 곡선 그래픽 + Kicker + 2줄 메인 타이틀 + 작성자 메타데이터
2. **Inquiry 3-Card (질문 카드)**:
   - Q1~Q3 컬러 배지 (Coral, Teal, Amber) + 볼드 질문 + 하단 전문가 인용구 틴트 박스
3. **Dark Divider (섹션 간지)**:
   - 딥 네이비 배경 + 슬레이트 네이비 곡선 + 화이트 대형 타이틀 + 코럴 불릿 리스트
4. **Metric 4-Card (지표 통계)**:
   - 52pt 초대형 통계 숫자 (4색 악센트 분할) + 2줄 설명 + 하단 전폭 딥 네이비 결론 배너
5. **Process Flow (프로세스 체브론)**:
   - 5단계 가로형 카드 + 화살표 연결선 + 주요 병목/강조 단계 딥 네이비 반전
6. **Quadrant Matrix (사분면 매트릭스)**:
   - X/Y 좌표축 + 우상단 위험 영역 살구색 하이라이트 + 우측 2단 인사이트 사이드바

---

## 3. 핵심 디자인 원칙 (Do's & Don'ts)

### Do's (반드시 지켜야 할 규칙)
- **Kicker + Title Pairing**: 모든 본문 슬라이드 상단에는 Kicker 태그(Coral, 11pt, 자간 +2px)와 타이틀(30pt 볼드)을 반드시 세트로 배치합니다.
- **Card Enclosure**: 본문의 모든 핵심 데이터는 `background: #FFFFFF; border: 1px solid #D4DFE9; border-radius: 12px;` 카드 안에 담습니다.
- **Bottom Takeaway Banner**: 분석 슬라이드 하단에는 카드 너비 전체를 덮는 전폭 딥 네이비 배너(`background: #0F2742; color: #FFFFFF;`)를 배치하여 경영진 결론을 전달합니다.
- **Strict Palette**: Primary Navy(`#0F2742`), Slate(`#1E4260`), Coral(`#D9481F`), Teal(`#14707E`), Amber(`#B07D18`)만 사용합니다.

### Don'ts (절대 하지 말아야 할 것)
- 클립아트나 3D 그래픽, 원색 계열(순수 빨강/파랑) 사용 금지
- 한 카드 안에 4문장 이상의 장문 텍스트 밀어넣기 금지
- 하단 테이크어웨이 배너 누락 금지

---

## 4. 스크립트 실행 가이드 (Bundled Scripts)

- **PDF 스타일 추출**:
  ```bash
  python scripts/extract_style.py path/to/deck.pdf
  ```
- **신규 슬라이드 덱 렌더링**:
  ```bash
  python scripts/render_deck.py --deck path/to/deck_definition.json --output result/
  ```
