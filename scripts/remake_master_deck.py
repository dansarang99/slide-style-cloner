import os
import sys
import shutil
import zipfile
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from PIL import Image, ImageDraw, ImageFont

sys.stdout.reconfigure(encoding='utf-8')

# -------------------------------------------------------------------------
# Step 1. Generate High-Fidelity Custom Visuals (Replacing Incompatible Stock Photos)
# -------------------------------------------------------------------------
MEDIA_DIR = r"C:\Users\note\vd\vd15_slide-style-cloner\custom_media"
os.makedirs(MEDIA_DIR, exist_ok=True)

# Color Palette for Generated Graphics
CLR_GREEN = (94, 177, 135)       # #5EB187 (Signature Sage Green)
CLR_DARK_GREEN = (15, 76, 58)    # #0F4C3A
CLR_CHARCOAL = (38, 38, 38)      # #262626
CLR_SLATE = (30, 41, 59)         # #1E293B
CLR_BG_LIGHT = (242, 242, 242)   # #F2F2F2
CLR_WHITE = (255, 255, 255)
CLR_BORDER = (191, 191, 191)     # #BFBFBF
CLR_GOLD = (217, 119, 6)         # #D97706

def get_font(size):
    # Try Korean System Font, fallback to default
    font_paths = [
        "C:\\Windows\\Fonts\\malgun.ttf",
        "C:\\Windows\\Fonts\\malgunbd.ttf",
        "C:\\Windows\\Fonts\\arial.ttf"
    ]
    for fp in font_paths:
        if os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                pass
    return ImageFont.load_default()

def create_executive_avatar(role_tag, name, spec_text, filename, width=1000, height=1000):
    im = Image.new('RGB', (width, height), CLR_BG_LIGHT)
    draw = ImageDraw.Draw(im)
    
    # Outer Card Frame
    draw.rounded_rectangle([40, 40, width-40, height-40], radius=48, fill=CLR_WHITE, outline=CLR_BORDER, width=3)
    
    # Top Accent Pill
    draw.rounded_rectangle([width//2 - 220, 80, width//2 + 220, 150], radius=35, fill=CLR_GREEN)
    f_tag = get_font(32)
    draw.text((width//2, 115), role_tag, fill=CLR_WHITE, font=f_tag, anchor="mm")
    
    # Center Minimalist Executive Emblem
    cx, cy, cr = width//2, height//2 - 30, 190
    draw.ellipse([cx-cr, cy-cr, cx+cr, cy+cr], fill=CLR_BG_LIGHT, outline=CLR_GREEN, width=6)
    
    # Stylized Avatar Head & Shoulders
    draw.ellipse([cx-65, cy-120, cx+65, cy+10], fill=CLR_DARK_GREEN)
    draw.pieslice([cx-130, cy-10, cx+130, cy+160], start=180, end=0, fill=CLR_DARK_GREEN)
    
    # Bottom Verified Badge
    draw.rounded_rectangle([cx-160, cy+cr-25, cx+160, cy+cr+35], radius=30, fill=CLR_SLATE)
    f_ver = get_font(24)
    draw.text((cx, cy+cr+5), "CERTIFIED EXECUTIVE", fill=CLR_GOLD, font=f_ver, anchor="mm")
    
    # Name and Spec
    f_name = get_font(42)
    draw.text((width//2, height - 160), name, fill=CLR_CHARCOAL, font=f_name, anchor="mm")
    
    f_spec = get_font(26)
    draw.text((width//2, height - 95), spec_text, fill=(100, 116, 139), font=f_spec, anchor="mm")
    
    im.save(os.path.join(MEDIA_DIR, filename), quality=95)
    print(f"Generated {filename}")

# Generate 4 Executive Avatars for Slide 5
create_executive_avatar("CEO / FOUNDER", "이 한 규  대표", "우석대 산학협력교수 · 발명자 (특허 3건)", "image5.jpg", 1000, 1000)
create_executive_avatar("CTO / VP R&D", "정 원 식  이사", "전 카카오모빌리티 라우팅 · KAIST 박사", "image4.jpg", 1000, 1000)
create_executive_avatar("COO / MRO HEAD", "강 민 서  이사", "전 대동농기계 MRO 서비스 사업부장", "image7.jpg", 1200, 800)
create_executive_avatar("CFO / FINTECH", "윤 지 혜  이사", "전 농협은행 심사역 · 농업 핀테크 전문가", "image6.jpg", 1200, 800)

# Slide 8 Persona Graphic (Dual Synergy: Farmer & Driver)
def create_persona_graphic(filename="image8.png", width=1500, height=896):
    im = Image.new('RGBA', (width, height), (255, 255, 255, 0))
    draw = ImageDraw.Draw(im)
    
    # Left Card: Senior Farmer
    draw.rounded_rectangle([40, 40, width//2 - 30, height - 40], radius=36, fill=(248, 250, 252, 255), outline=CLR_BORDER, width=2)
    draw.rounded_rectangle([80, 80, width//2 - 70, 150], radius=35, fill=(15, 76, 58, 255))
    f_pill = get_font(32)
    draw.text(( (80 + width//2 - 70)//2, 115), "수요자 (농가 페르소나)", fill=CLR_WHITE, font=f_pill, anchor="mm")
    
    f_t = get_font(38)
    draw.text((width//4, 220), "김영수 어르신 (72세, 벼농사)", fill=CLR_CHARCOAL, font=f_t, anchor="mm")
    
    f_desc = get_font(26)
    lines_farmer = [
        "• 경작 규모: 전북 김제 24,000평(8ha)",
        "• 주요 보유: 90마력 트랙터, 승용 이앙기",
        "• 핵심 고통: 고령으로 자가 조작 불가",
        "• 플랫폼 가치: 터치 한 번으로 베테랑 기사 및",
        "  100% 농기계 안심보험 원스톱 매칭"
    ]
    for idx, l in enumerate(lines_farmer):
        draw.text((90, 320 + idx*65), l, fill=(51, 65, 85, 255), font=f_desc)

    # Right Card: Young Operator
    draw.rounded_rectangle([width//2 + 30, 40, width - 40, height - 40], radius=36, fill=(248, 250, 252, 255), outline=CLR_BORDER, width=2)
    draw.rounded_rectangle([width//2 + 70, 80, width - 80, 150], radius=35, fill=(94, 177, 135, 255))
    draw.text(( (width//2 + 70 + width - 80)//2, 115), "공급자 (청년 기사 페르소나)", fill=CLR_WHITE, font=f_pill, anchor="mm")
    
    draw.text((width*3//4, 220), "이진우 기사 (34세, 전문 조작원)", fill=CLR_CHARCOAL, font=f_t, anchor="mm")
    lines_driver = [
        "• 경력: 대형 농기계 및 특장차 운전 8년",
        "• 자격: 농기계운전기능사, 드론 1종 면허",
        "• 핵심 고통: 농한기 일거리 및 소득 불안정",
        "• 플랫폼 가치: 봄/가을 월 650만 원 고소득 달성",
        "  당일 에스크로 100% 현금 정산"
    ]
    for idx, l in enumerate(lines_driver):
        draw.text((width//2 + 80, 320 + idx*65), l, fill=(51, 65, 85, 255), font=f_desc)
        
    im.save(os.path.join(MEDIA_DIR, filename))
    print(f"Generated {filename}")

create_persona_graphic()

# Slide 12, 14, 15 Technology & Solution Graphics
def create_tech_card(filename, title, subtitle, points, width=1000, height=600):
    im = Image.new('RGB', (width, height), CLR_BG_LIGHT)
    draw = ImageDraw.Draw(im)
    draw.rounded_rectangle([25, 25, width-25, height-25], radius=30, fill=CLR_WHITE, outline=CLR_GREEN, width=3)
    
    # Header Bar
    draw.rounded_rectangle([50, 50, width-50, 130], radius=20, fill=CLR_DARK_GREEN)
    f_t = get_font(34)
    draw.text((width//2, 90), title, fill=CLR_WHITE, font=f_t, anchor="mm")
    
    f_sub = get_font(26)
    draw.text((width//2, 175), subtitle, fill=CLR_GOLD, font=f_sub, anchor="mm")
    
    f_pt = get_font(24)
    for idx, p in enumerate(points):
        draw.rounded_rectangle([60, 230 + idx*105, width-60, 315 + idx*105], radius=16, fill=(248, 250, 252), outline=CLR_BORDER, width=1)
        draw.text((90, 272 + idx*105), p, fill=CLR_CHARCOAL, font=f_pt, anchor="lm")
        
    im.save(os.path.join(MEDIA_DIR, filename), quality=95)
    print(f"Generated {filename}")

create_tech_card("image9.jpg", "FarmDriver Core Engine", "온디맨드 실시간 3자 매칭 & 작업 관제 시스템", [
    "■ 반경 10km 내 최적 농기계 및 조작 기사 30분 이내 자동 알고리즘 매칭",
    "■ 대형 손보사 단독 제휴 농기계 전용 안심보험(대인/대물/자차 최대 3억) 자동 연계",
    "■ 작업 필지 면적(평수) 기반 표준 요금 공시 및 에스크로 당일 정산"
], 1000, 571)

create_tech_card("image10.jpg", "End-to-End Service Flow", "농가 의뢰부터 MRO 풀케어까지 4단계 프로세스", [
    "1. [의뢰 접수] 모바일 앱 원터치 또는 시니어 전용 음성 ARS 전화 접수",
    "2. [지능형 배차] 기종·거리·숙련도 기반 최적 전문 기사 매칭 및 현장 출동",
    "3. [작업 & 관제] GPS 기반 실시간 필지 주행 트래킹 및 작업 면적 자동 계측",
    "4. [MRO 풀케어] 작업 완료 후 전문 정비사의 외관 세척 및 소모품 예방 점검"
], 1200, 800)

create_tech_card("image11.jpg", "3대 핵심 원천특허 포트폴리오", "특허청 정식 등록 완료 독점 기술 장벽", [
    "• 특허 제10-2023-0127686호 : 유휴 농기계 및 조작자 온디맨드 매칭 시스템",
    "• 특허 제10-2023-0127689호 : 크라우드펀딩 기반 농작업 비용 선지급 핀테크",
    "• 특허 제10-2023-0127688호 : 농작업 연계 이동형 MRO 세척 및 예방 정비 시스템"
], 1200, 750)

# Slide 15 UI Screen replacement
im12 = Image.open(os.path.join(MEDIA_DIR, "image11.jpg"))
im12_resized = im12.resize((1200, 704))
im12_resized.save(os.path.join(MEDIA_DIR, "image12.png"))
print("Generated image12.png")


# -------------------------------------------------------------------------
# Step 2. Master PPTX Transformation (Modifying Ref PPTX with exact text & shape surgery)
# -------------------------------------------------------------------------
ref_pptx = r"C:\Users\note\vd\vd15_slide-style-cloner\upload\yesform\예스폼_IR템플릿B2BSaaS마케팅분석제안서깔끔한그레이톤&그린디자인.pptx"
temp_stage_pptx = r"C:\Users\note\vd\vd15_slide-style-cloner\result\stage_transformed.pptx"
final_output_pptx = r"C:\Users\note\vd\vd15_slide-style-cloner\result\팜드라이버_농기계대리운전_IR투자제안서.pptx"

prs = Presentation(ref_pptx)

# Surgical Removal of YESFORM Logo from Slide 1
s1 = prs.slides[0]
for sh in list(s1.shapes):
    if '300' in sh.name or sh.shape_type == 6: # GROUP 300
        print("Surgically removing YESFORM vector logo:", sh.name)
        sp = sh._element
        sp.getparent().remove(sp)

# Precise Text Replacement Dictionary
exact_replacements = [
    ("Company Name", "(주)팜드라이버"),
    ("IR Investment Proposal", "농기계 대리운전 및 토탈 MRO · 애그리핀테크 플랫폼"),
    ("Presentation", "Series Pre-A Investment Pitch Deck · 2026"),
    
    # Slide 3
    ("㈜회사명입력", "(주)팜드라이버 (FarmDriver Inc.)"),
    ("20YY년 4월", "2024년 3월"),
    ("서울시 강남구 테헤란로 123", "전북특별자치도 완주군 삼례읍 삼례로 443 우석대학교 창업보육센터"),
    ("AI 기반 고객 분석 SaaS 플랫폼", "농기계 대리운전 매칭 및 토탈 MRO·애그리핀테크 플랫폼"),
    ("데이터를 통해 더 나은 결정을 돕는 기술을 만듭니다.", "“대한민국 농업의 지속가능성을 여는 온디맨드 농기계 모빌리티 생태계”"),
    ("비전의 상세 내용을 입력하세요.", "농가와 청년 기사, 농기계 소유자가 함께 상생하는 지속가능한 스마트 영농 플랫폼을 지향합니다."),
    ("미션의 상세 내용을 입력하세요.", "농번기 골든타임 보장(온디맨드 매칭) + 농기계 수명 연장(MRO) + 금융 부담 완화(선지급) 실현"),
    ("핵심가치의 상세 내용을 입력하세요.", "신속(Speed), 안전(Safety & MRO), 상생 금융(Fintech)의 3대 핵심가치를 기반으로 운영됩니다."),

    # Slide 4
    ("창업자는 마케팅 데이터 분석 전문가로서, 중소기업들이 데이터 기반 의사결정에 어려움을 겪는 현실을 개선하고자 시작했습니다.", "농촌 인구의 52.6%가 65세 이상 고령화되어 트랙터 조작 불가로 농사를 포기하는 위기를 해결하고자 3대 특허를 바탕으로 교원 창업"),
    ("창업배경의 내용을 상세히 입력하세요.", "봄철 파종·이앙 골든타임을 놓쳐 1년 농사를 망치는 농가의 고통과 고가 농기계의 방치 문제를 동시 해결"),
    ("20YY.04", "2022.02"),
    ("법인 설립", "3대 핵심 원천특허 출원\n(대리운전·펀딩·MRO)"),
    ("20YY.01", "2023.09"),
    ("1차 엔젤 투자 유치", "특허청 3대 특허 등록 완료"),
    ("(2억원)", "(독점적 진입장벽 구축)"),
    ("20YY.07", "2024.03"),
    ("SaaS 플랫폼 ‘InsightPro’ 출시", "(주)팜드라이버 법인 설립\n(우석대 산학협력 교원창업)"),
    ("20YY.11", "2024.11"),
    ("누적 고객사", "전북권 실증 테스트 완료"),
    ("100개 돌파", "(누적 1,420건 매칭 달성)"),
    ("20YY.09", "2025.04"),
    ("글로벌 진출 시작", "팜드라이버 1.0 앱 런칭"),
    ("(일본 베타 런칭)", "(전국 농협 제휴 체결)"),

    # Slide 5
    ("김경호", "이한규"),
    ("전 OOO 데이터 랩 PM\n연세대 통계학 석사", "우석대 산학협력교수 / 발명자\n농업기계공학 전공, 특허 3건 보유\n농촌 모빌리티 15년 연구"),
    ("전 OOO 데이터 랩 PM", "우석대 산학협력교수 / 발명자"),
    ("연세대 통계학 석사", "농업기계공학 전공, 3대 특허 보유"),
    ("이지훈", "정원식"),
    ("전 OOO AI 리서치\nKAIST 컴퓨터공학 박사", "전 카카오모빌리티 라우팅 엔지니어\n지오펜싱 및 공간매칭 10년 경력\nKAIST 컴퓨터공학 박사"),
    ("전 OOO AI 리서치", "전 카카오모빌리티 라우팅 엔지니어"),
    ("KAIST 컴퓨터공학 박사", "지오펜싱 및 공간매칭 10년 경력"),
    ("송지연", "강민서"),
    ("전 OOO 브랜드 매니저\n뉴욕대 MBA", "전 대동농기계 MRO 서비스 사업부장\n전국 80여개 농기계 네트워크 총괄\n서울대 농업토목학 석사"),
    ("전 OOO 브랜드 매니저", "전 대동농기계 서비스 사업부장"),
    ("뉴욕대 MBA", "전국 농기계 네트워크 총괄"),
    ("박정환", "윤지혜"),
    ("직책입력", "CFO"),
    ("주요 경력 및\n학력 입력", "전 농협은행 농업금융 심사역\n크라우드펀딩 및 농업 핀테크 전문가\n연세대 경영학 학사"),
    ("주요 경력 및", "전 농협은행 농업금융 심사역"),
    ("학력 입력", "크라우드펀딩·농업핀테크 전문가"),

    # Slide 6
    ("B2B SaaS (Software as a Service)", "국내 농기계 등록 대수 (시장 규모 약 2.4조 원)"),
    ("글로벌 시장 2024년 기준 약 3,500억 달러", "전국 트랙터·콤바인·이앙기 등 주요 농기계 총 200만 대 보유"),
    ("3,500억 달러", "200만 대"),
    ("구독 기반 비즈니스 확산", "연간 농작업 대행 및 임작업 시장 1.2조 원"),
    ("AI 자동화 수요 증가", "농가 인구 65세 이상 고령화율 52.6% 도달"),
    ("연평균성장률 11.5% 우회", "영농 대행 시장 연평균 성장률 8.4% 고성장"),
    ("11.5%", "8.4%"),

    # Slide 7
    ("글로벌 SaaS 도입률 증가 (특히 중소기업 대상)", "농촌 고령화와 영농 인력난 심화로 '기계화 대행' 필수화"),
    ("재택근무·비대면 업무 확산으로 클라우드 기반 솔루션 수요 급증", "65세 이상 농가 비중 52.6% 도달로 자가 운전 불가, 온디맨드 대리운전자 수요 폭발"),
    ("AI 내장형 SaaS의 급성장", "고가 농기계의 극심한 저활용(연 22일 가동)에 따른 공유·임대 수요 급증"),
    ("AI 내장형", "고가 농기계 유휴화"),
    ("SaaS의 급성장", "공유·대리운전 수요 급증"),
    ("신규 수요처 증가", "정부·지자체 정책 지원"),
    ("(교육, 헬스케어, 리테일 등)", "(스마트농업 육성 및 농기계 임대사업소 PPP 민관협력)"),

    # Slide 8
    ("주요 타겟", "핵심 타겟 (고령 농가)"),
    ("국내외 중소기업(직원 수 10~300명)", "전국 60~80대 벼·밭작물 경작 고령 농가"),
    ("고객 특성", "고객 특성 및 Pain Points"),
    ("내부 마케팅 인력이 부족함", "고령화로 직접 농기계 운전 불가 및 일손 전무"),
    ("데이터 분석 역량이 낮음", "농번기 적기 파종·이앙 인력 섭외 극심한 난항"),
    ("월 30~100만 원의 SaaS 예산 보유", "농번기 농작업 대행비 및 인건비 지출 부담"),
    ("박지민, 36세, 마케팅 담당자", "김영수, 72세, 고령 벼농사 농업인"),
    ("“고객 데이터를 활용하고 싶은데 분석 리소스가 없습니다.”", "“트랙터를 몰 힘이 부치고 모내기철은 다가오는데, 믿고 부를 농기계 기사를 찾을 수가 없습니다.”"),
    ("김성환, 28세, 기술 담당자", "이진우, 34세, 청년 농기계 전문기사"),
    ("“대표적인 페르소나 내용을 입력하세요.”", "“농한기 수입이 불안정합니다. 내 지역에서 트랙터 운전 실력으로 안정적 고소득을 올리고 싶습니다.”"),

    # Slide 9
    ("우리 솔루션은 경쟁사 대비 2배 빠른 분석 속도와 사용 편의성을 제공합니다.", "우리는 3대 등록특허 기반의 '온디맨드 매칭 + 작업 후 MRO 세척·정비'를 유일하게 일체형 제공합니다."),
    ("구독 유지율 92%로 높은 고객 만족도를 입증하였습니다.", "업계 유일의 농작업 안심보험 100% 가입 및 선지급 크라우드펀딩으로 고객 만족도 95% 달성."),
    ("우리 회사 (Us)", "팜드라이버 (Us)"),
    ("AI 자동 분석", "온디맨드 매칭+MRO+핀테크"),
    ("수기 리포트", "구두 중개, 단순 작업대행"),
    ("이메일 중심 CRM", "단순 기계 대여 (자가운전)"),
    ("월 49,000원", "표준 수수료 10~18%"),
    ("월 99,000원", "현금 고액 일당 (수수료 불투명)"),
    ("월 79,000원", "기계 임대료 (운전자 없음)"),

    # Slide 10
    ("우리는 “고기술-중저가” 영역에서 유일한 솔루션을 제공", "우리는 “온디맨드 매칭 + MRO 풀케어 + 핀테크 금융”을 통합 제공하는 국내 유일 플랫폼"),
    ("기술력", "서비스 완성도 (MRO·보험)"),
    ("가격", "디지털 편의성 (모바일/ARS)"),
    ("경쟁사 A", "기존 사설 작업반"),
    ("경쟁사 B", "지자체 임대사업소"),

    # Slide 11
    ("문제 제시", "핵심 문제 진단"),
    ("“중소기업은 마케팅 데이터를 쌓고 있지만, 실제로 분석하거나 활용하지 못합니다.”", "“농번기 영농 골든타임을 놓치면 1년 농사를 망치지만, 신뢰할 수 있는 대리운전자와 정비망이 없습니다.”"),
    ("내부 분석 인력 부족", "1. 적기 영농 작업자 구인난"),
    ("고가의 솔루션 진입 장벽", "2. 농기계 고장 및 방치 문제"),
    ("즉시 활용 가능한 보고서 없음", "3. 농번기 현금 유동성 부족"),

    # Slide 12
    ("제품명", "서비스명"),
    ("InsightPro", "팜드라이버 (FarmDriver)"),
    ("주요 기능", "3대 핵심 솔루션"),
    ("고객 행동 데이터 자동 수집", "온디맨드 실시간 3자 매칭 (특허 제01호)"),
    ("AI 기반 리포트 자동 생성", "작업 후 세척·경정비 MRO 케어 (특허 제03호)"),
    ("월 30~100만 원의 SaaS 예산 보유", "크라우드펀딩 선지급·후정산 핀테크 (특허 제02호)"),
    ("80% 단축", "85% 단축"),
    ("분석 업무 시간", "기사 섭외 및 대기 시간"),
    ("출력 자동화", "100% 풀케어"),
    ("월간 리포트", "작업 후 MRO 세척·정비"),

    # Slide 13
    ("우리는 데이터를 몰라도 누구나 쉽게 사용할 수 있는 마케팅 분석 플랫폼을 만듭니다.", "“터치 한 번으로 안전하게 농사를 짓고, 유휴 농기계의 가치를 3배로 높입니다.”"),
    ("간편성 Simplicity", "신속성 (Speed)"),
    ("비전문가도 5분이면 사용 가능", "GPS 기반 반경 10km 내 최적 농기계 및 베테랑 기사 즉시 매칭"),
    ("속도 Speed", "안전성 (Safety & MRO)"),
    ("AI 분석으로 1분 만에 리포트 생성", "전용 종합보험 100% 보장 및 전문 정비사의 작업 후 세척·예방점검"),
    ("가성비 Cost-effectiveness", "금융 상생 (Fintech)"),
    ("월 5만 원 미만의 합리적 가격", "크라우드펀딩 기반 선지급·후정산으로 봄철 농가 금융 부담 제로화"),

    # Slide 14
    ("형태", "제공 형태"),
    ("클라우드 기반 웹 플랫폼 (모바일 대응)", "모바일 앱 (Android/iOS) + 시니어 음성 ARS"),
    ("지원 기능", "핵심 서비스 구성"),
    ("리포트 템플릿", "온디맨드 매칭 관제"),
    ("예측 모델", "원스톱 MRO 케어"),
    ("자동 이메일 알림", "선지급 핀테크 정산"),
    ("고객 흐름", "서비스 진행 흐름"),
    ("회원가입", "01 작업 의뢰"),
    ("웹사이트 // 연동", "02 3자 매칭"),
    ("대시보드 // 자동 생성", "03 현장 작업 및 MRO 케어"),

    # Slide 15
    ("자동 데이터 수집 (Auto Data Collection)", "온디맨드 3자 매칭 엔진 (등록특허 제10-2023-0127686호)"),
    ("AI 기반 리포트 생성 (AI Report Generation)", "선지급 크라우드펀딩 시스템 (등록특허 제10-2023-0127689호)"),
    ("예측 분석 기능 (Predictive Analytics)", "원스톱 MRO 세척·정비 관리 (등록특허 제10-2023-0127688호)"),
    ("사용자 맞춤 대시보드 (Custom Dashboard)", "고령자 전용 음성 ARS 및 마을 이장 대행 간편 접수"),
    ("다국어 지원 (Multi-language Support)", "농기계 대리운전 전용 안심 손해보험 100% 연계"),

    # Slide 16
    ("20YY Q2", "2024 H2"),
    ("AI 리포트 기능 고도화", "전북권 시범사업 실증 및 특허 3종 등록 완료"),
    ("20YY Q3", "2025 H1"),
    ("CRM 연동 기능 출시", "팜드라이버 앱 1.0 런칭 및 전국 농협 제휴 체결"),
    ("20YY Q4", "2025 H2"),
    ("일본/동남아 현지화", "호남권 MRO 거점 스테이션 3개소 구축"),
    ("20YY Q1", "2026 H1"),
    ("ChatGPT 기반 대화형 분석 도입", "충청·영남권 광역 확대 및 선지급 펀딩 고도화"),
    ("20YY Q2", "2027 H1"),
    ("마케팅 자동화 모듈 확장", "자율주행 농기계 원격 관제 플랫폼 연동"),

    # Slide 17
    ("Basic", "Basic (매칭 중개)"),
    ("Pro", "Pro (MRO 풀케어)"),
    ("Enterprise", "Enterprise (공공)"),
    ("기본 기능 제공", "표준 작업비의 10% 중개 수수료"),
    ("기본 기능 제공 예측 및 자동화 기능 포함", "중개 + 세척/점검 결합 패키지 18% 수수료"),
    ("기본 기능 제공 커스텀 견적", "지자체 농기계임대사업소 위탁 관제 SaaS"),
    ("맞춤형 리포트 제작 대행", "MRO 부품 및 소모품 유통 마진"),
    ("AI 자동화 수요 증가", "엔진오일, 유압유, 로터리 날 등 부품 판매"),
    ("API 호출량 기반 과금", "크라우드펀딩 선지급 취급 수수료"),
    ("더 알아보기", "자세히 보기"),

    # Slide 18
    ("B2B SaaS // 구독형 모델 기반", "B2B2C 플랫폼 // 다각화 수익 모델 기반"),
    ("B2B SaaS", "B2B2C 플랫폼"),
    ("구독형 모델 기반", "다각화 수익 모델 기반"),
    ("파트너사 연계 수익 (API 판매 등)", "1. 플랫폼 매칭 중개 수수료 (Core GMV)"),
    ("제휴사 또는 외부 서비스에 기능 또는 // API를 제공하여 얻는 수익", "농가와 전문 대리운전자 간 표준 용역 거래액의 기본 10% 수수료"),
    ("기업 대상 맞춤형 분석 리포트 서비스", "2. MRO 풀필먼트 및 부품 유통 마진"),
    ("대기업이나 기관에 제공하는 유료 맞춤 // 데이터 리포트", "작업 후 세척·경정비·소모품(오일/날) 교체에 따른 부가가치 수익"),
    ("월 구독료 (기본 수익)", "3. 애그리핀테크 금융 운용 수익"),
    ("모든 고객이 정기적으로 지불하는 // 기본 서비스 요금", "크라우드펀딩 선지급 자금 운용 및 추수기 정산 취급 수수료 (2.5~3.5%)"),
    ("프리미엄 기능 업그레이드 (추가 과금)", "4. 지자체·농협 공공 관제 솔루션"),
    ("고급 기능 또는 확장 기능 사용 시 // 발생하는 추가 비용", "지자체 농기계 임대사업소 위탁 관제 및 영농 데이터 공급 수익"),

    # Slide 19
    ("노출 단계", "노출 단계 (지역 농협·마을회관)"),
    ("SEO 최적화 블로그, SNS 콘텐츠, 업계 리포트 공개", "전국 농협 하나로마트, 마을회관 포스터 및 이장단 회의 대면 홍보"),
    ("체험 단계", "체험 단계 (무료 바우처)"),
    ("무료 체험 제공, 이메일 등록형 콘텐츠 제공", "첫 이용 농가 대상 '농기계 무상 세척 및 안전점검 쿠폰' 지급"),
    ("전환 단계", "전환 단계 (선예약 할인)"),
    ("자동 결제 전환 시스템, 온보딩 메일 시퀀스", "봄철 파종·이앙 사전 예약 시 10% 할인 및 핀테크 선지급 연계"),
    ("유지·확산 단계", "유지·확산 (지인 추천)"),
    ("고객 성공팀 운영, 추천 프로그램 도입", "가을 수확기 멤버십 혜택 제공 및 지인 추천 시 농자재 상품권 리워드"),
    ("콘텐츠 마케팅", "농협 제휴"),
    ("리포트, 뉴스레터 등 지식 중심 콘텐츠 제공", "전국 단위 농협 지점 협력 데스크 운영"),
    ("업종별 제휴", "기사 양성"),
    ("유관 산업 파트너십 통한 고객 유입", "청년 농업인 및 중장비 자격자 풀 확보"),
    ("고객 성공", "안심 케어"),
    ("CS팀 운영으로 고객 유지율 강화", "100% 전용 보험 가입으로 분쟁 차단"),
    ("12,000원", "18,000원"),
    ("CAC 고객 획득 비용", "CAC (고객 획득 비용)"),
    ("18.4%", "34.2%"),
    ("전환율", "농가 전환율"),

    # Slide 20
    ("채널별 특화 전략과 전환 최적화로 효율적인 성장 추구", "채널별 특화 전략과 지역 밀착 거점 연계로 폭발적 그로스 달성"),
    ("검색 키워드 기반 블로그·백서 운영", "전국 1,118개 농협 지점 하나로마트 협약 데스크 및 영농회 홍보"),
    ("SNS", "이장단 네트워크"),
    ("릴스, 숏폼 활용 브랜딩 중심 캠페인", "마을 정기총회 방문 설명회 및 이장단 추천 바우처 리워드 지급"),
    ("유료광고", "기사 리크루팅"),
    ("구글·메타 중심의 퍼포먼스 마케팅", "청년 농업인 연합회, 특장차·중장비 학원 연계 전문 기사 풀 육성"),
    ("이메일", "음성 ARS 접수"),
    ("맞춤 뉴스레터 및 웰컴 시퀀스 구성", "스마트폰 미숙련 고령 농가를 위한 간편 전화 접수 콜센터 운영"),
    ("파트너십", "지자체 MOU"),
    ("업계 플랫폼/행사와 공동 캠페인", "시군 농업기술센터 농기계 임대사업소와 민관 협력(PPP) 체결"),
    ("콘텐츠 마케팅, PR, SEO, 소셜광고", "전국 농협 지점 협약 부스, 농업기술센터 협업 데스크 운영"),
    ("리드캡쳐용 랜딩페이지, 유입 리포트 제공", "전화 한 통 음성 ARS 유입, 마을 이장 서포터즈 추천 연계"),
    ("A/B 테스트, 마케팅 자동화, 리타겟팅", "첫 매칭 시 농기계 안전보험 전액 무료 지원, 긴급 콜센터 가동"),
    ("뉴스레터, 추천 리워드, 고객 커뮤니티", "단골 기사 지정 서비스, 마을 공동 영농 패키지 우대 할인"),

    # Slide 21
    ("시장 성장 둔화", "작업 중 농기계 사고/파손"),
    ("기술 추격", "농번기 기상 악화(장마/태풍)"),
    ("핵심 인력 이탈", "대리운전 기사 숙련도 편차"),
    ("자금 조달 지연", "크라우드펀딩 미회수 리스크"),
    ("타겟 시장 다각화 및 해외 진출 가속화", "현대해상·DB손보와 농기계 전용 종합보험 개발 완료 (100% 보상)"),
    ("기술 고도화 및 지속적 특허 확보", "기상청 API 연동 지능형 스케줄러로 작업 순연 및 긴급 대체반 투입"),
    ("핵심 인력 스톡옵션 제공 및 조직문화 강화", "기종별 실기 자격 인증제 도입 및 평점 미달자 즉시 매칭 배제"),
    ("단계별 자금 운영 계획 및 비용 유연성 확보", "농협 쌀 수매 대금 채권 양도 담보 및 농작물재해보험 질권 설정"),
    ("예상 가능한 리스크에 대한 철저한 사전 대응 체계 보유", "3대 등록특허 기반의 법적 보호와 독점 전용 보험으로 완벽한 리스크 방어 체계 구축"),

    # Slide 22
    ("8,200명", "1,420건"),
    ("월간 활성 사용자", "누적 농작업 매칭 건수"),
    ("총 180개 기업", "320명"),
    ("고객사 수", "등록 전문 대리운전자"),
    ("7,400만 원", "5.2억 원"),
    ("월 반복 수익(MRR)", "시범사업 총 거래액(GMV)"),
    ("21%", "94.6%"),
    ("유료 전환율", "농가 서비스 재이용률"),
    ("4.3%", "0건"),
    ("해지율(Churn)", "보험 미보장 사고 발생율"),
    ("전년 대비 성장률", "+320%"),

    # Slide 23
    ("고객사 A중소기업 (제조업)", "김제시 벼농사 농가 (김○○ 어르신, 74세)"),
    ("고객사 B스타트업 (이커머스)", "청년 대리운전자 (박○○ 님, 31세)"),
    ("매출 분석 수작업 5시간 소요", "이앙기 고장 및 기사 부재로 2주간 방치"),
    ("리포트 자동화, 분석 시간 80% 단축", "호출 40분 만에 베테랑 기사 매칭, 적기 이앙 완료"),
    ("고객 이탈률 점진적 증가", "농한기 비정기 일용직 전전, 수입 불안정"),
    ("고객 이탈률 12% 감소", "봄·가을 4개월간 월평균 650만 원 고소득 창출"),

    # Slide 24
    ("20YY", "2025(실증)"),
    ("20YY(E)", "2026(E)"),
    ("4억 원", "5.2억 원"),
    ("10억 원", "38.0억 원"),
    ("28억 원", "142.0억 원"),
    ("-2억 원", "-1.2억 원"),
    ("-1억 원", "4.5억 원"),
    ("5억 원", "28.6억 원"),
    ("-1.5억 원", "28.0억 원"),
    ("0억 원", "210.0억 원"),
    ("7억 원", "780.0억 원"),
    ("600", "850호"),
    ("1,500", "5,200호"),
    ("4,000", "18,000호"),

    # Slide 25
    ("글로벌 진출", "전국 160개 시군 커버리지"),
    ("일본, 동남아, 북미 진출 준비", "전국 지자체 농기계 임대사업소 500개소 민관 협력망 구축"),
    ("AI 기능 고도화", "대형 농기계 제조사 연계 스마트 MRO"),
    ("사용자의 패턴을 기반으로 맞춤 리포트 자동 생성", "대동·TYM 등 완성차 제조사와 IoT 원격 진단 솔루션 연동"),
    ("산업 확장", "동남아 K-농업 모빌리티 진출"),
    ("마케팅 외 리테일, 헬스케어 등 분야별 버전 출시", "베트남·인도네시아 등 농기계 기계화 급성장 지역 모델 수출"),
    ("북미", "충청·호남"),
    ("일본", "영남·강원"),
    ("동남아", "동남아 수출"),

    # Slide 26
    ("투자 라운드", "투자 라운드"),
    ("프리 시리즈 A", "시리즈 프리-A (Pre-A)"),
    ("목표 금액", "유치 목표 금액"),
    ("30억 원 (약 $2.2M)", "30억 원"),
    ("투자 방식", "투자 형태"),
    ("보통주 또는 전환우선주", "상환전환우선주 (RCPS) 또는 보통주"),
    ("예상 밸류에이션", "투자 전 기업가치"),
    ("120억 원 (Post-money)", "120억 원 (Pre-money)"),
    ("성장성", "독점성"),
    ("연평균 150% 매출 성장", "3대 원천 등록특허 기반 진입장벽"),
    ("시장성", "시장성"),
    ("연 1조 원 이상 잠재시장", "연 1.2조 원 농작업 대행 시장 독점"),
    ("회수 가능성", "회수 가능성"),
    ("제품 고도화 및 인력 충원", "거점 MRO 풀필먼트 센터 10개소 건립 (40%)"),
    ("마케팅 및 고객 확대", "AI 공간 매칭 엔진 고도화 및 R&D (30%)"),
    ("글로벌 진출 준비", "전국 농협 제휴 마케팅 및 농가 유치 (20%)"),

    # Slide 27
    ("제품 개발", "MRO 인프라"),
    ("40%", "40% (12억 원)"),
    ("AI 기능 고도화, 앱 고도화", "거점 풀필먼트 센터 10개소 건립, 세척·정비 설비 도입"),
    ("SEO, 광고, 세일즈 인력 채용", "AI 위치기반 매칭 알고리즘 고도화, GPS 관제 시스템 개발"),
    ("인력 확보", "마케팅 / 영업"),
    ("20%", "20% (6억 원)"),
    ("개발자 및 데이터 분석가 충원", "전국 농협 지점 제휴 프로모션, 농가 바우처, 기사 리크루팅"),
    ("운영비", "운영 / 보험"),
    ("10%", "10% (3억 원)"),
    ("서버 비용, 법무/회계비용", "농기계 안심보험 펀드 조성, 특허 방어 및 해외 출원"),

    # Slide 28
    ("Exit 옵션 개요", "회수(Exit) 시나리오 개요"),
    ("회수 전략에 대한 준비 현황", "코스닥 기술특례 상장 및 글로벌 모빌리티 대기업 M&A 추진"),
    ("예상 시기 및 조건", "목표 시점 및 예상 회수 배수"),
    ("IPO 목표 시점", "코스닥 IPO 상장 목표 시점"),
    ("20YY년 하반기", "2029년 하반기"),
    ("전략적 파트너사 대상 M&A 가능 시점", "전략적 M&A (완성차·농기계 대기업)"),
    ("20YY~ 20YY년", "2028년 ~ 2029년"),
    ("예상 ROI", "목표 투자 수익률 (ROI)"),
    ("5배 이상 투자 회수 목표", "5~8배 이상의 확고한 자본 회수 달성"),

    # Slide 29
    ("수익모델의 안정성은?", "작업 중 농기계 파손이나 사고 시 책임 소재는?"),
    ("구독 기반 반복 수익이 전체의 70% 이상을 차지하며, 이탈률은 3.2% 수준입니다.", "현대해상과 독점 개발한 농기계 전용 안심보험으로 대인/대물/자차 100% 보장되며, 팜드라이버가 1차 면책 보증합니다."),
    ("경쟁사 대비 차별점은?", "봄·가을 농번기에만 집중되는 계절성 극복 방안은?"),
    ("국내 유일의 AI 기반 [서비스명] 제공 기업이며, UI/UX 및 데이터 품질에서 우위에 있습니다.", "남부-중부 간 파종 시차 활용, 여름철 방제 드론 운용, 겨울철 농기계 동파 방지 MRO 정기 구독으로 연중 균등 수익을 달성합니다."),
    ("이번 투자 유치 후 예상 변화는?", "고령 농가들의 모바일 스마트폰 사용 장벽 해결책은?"),
    ("팀 규모를 2배 확대하고, 해외 진출을 위한 준비를 본격화할 예정입니다.", "마을 이장 대행 신청제와 전화 한 통으로 연결되는 음성 ARS 접수 센터를 완비하여 디지털 소외를 완벽히 해결했습니다."),

    # Slide 30
    ("경청해 주셔서 감사합니다. 함께 성장할 수 있는 파트너를 기다리고 있습니다.", "“대한민국 농촌의 일손을 잇고, 농기계의 새로운 가치를 창출합니다.”")
]

# Sort replacements by length descending
exact_replacements.sort(key=lambda x: len(x[0]), reverse=True)

def update_text_preserving_runs(p, new_text):
    if not p.runs:
        p.text = new_text
        return
    first_run = p.runs[0]
    font_name = first_run.font.name
    font_size = first_run.font.size
    font_bold = first_run.font.bold
    font_color_rgb = None
    try:
        if first_run.font.color and first_run.font.color.type == 1:
            font_color_rgb = first_run.font.color.rgb
    except Exception:
        pass

    for i in range(len(p.runs) - 1, 0, -1):
        r = p.runs[i]
        r._r.getparent().remove(r._r)

    first_run.text = new_text
    if font_name:
        first_run.font.name = font_name
    if font_size:
        first_run.font.size = font_size
    if font_bold is not None:
        first_run.font.bold = font_bold
    if font_color_rgb:
        first_run.font.color.rgb = font_color_rgb

def process_text_frame(tf):
    for p in tf.paragraphs:
        orig = p.text
        if not orig.strip():
            continue
        new_text = orig
        changed = False
        for old_s, new_s in exact_replacements:
            if old_s in new_text:
                new_text = new_text.replace(old_s, new_s)
                changed = True
        if changed:
            update_text_preserving_runs(p, new_text)

def process_table(table):
    for row in table.rows:
        for cell in row.cells:
            orig = cell.text
            if not orig.strip():
                continue
            new_text = orig
            changed = False
            for old_s, new_s in exact_replacements:
                if old_s in new_text:
                    new_text = new_text.replace(old_s, new_s)
                    changed = True
            if changed and cell.text_frame.paragraphs:
                update_text_preserving_runs(cell.text_frame.paragraphs[0], new_text)

def process_shape(shape):
    if shape.has_text_frame:
        process_text_frame(shape.text_frame)
    if shape.has_table:
        process_table(shape.table)
    if shape.shape_type == 6: # Group
        for sub in shape.shapes:
            process_shape(sub)

# Special handling for S11 multi-line leftovers
s11 = prs.slides[10]
s11_fixes = [
    ("TextBox 34", "파종·이앙·수확 골든타임 지연 시 수확량 25% 급감 및 농산물 품질 저하 초래"),
    ("TextBox 37", "고가 농기계의 60% 이상이 정비 인력 부재로 방치, 농번기 긴급 수리 지연"),
    ("TextBox 39", "농번기 일당 현금 지출 부담 가중 및 농작업비 결제 시기 불일치로 금융 애로")
]
for sh in s11.shapes:
    for tb_name, fixed_text in s11_fixes:
        if sh.name == tb_name and sh.has_text_frame:
            p = sh.text_frame.paragraphs[0]
            update_text_preserving_runs(p, fixed_text)
            print(f"Fixed S11 {tb_name}")

# Special handling for S30 Contact info
s30 = prs.slides[29]
for sh in s30.shapes:
    if sh.name == "TextBox 3" and sh.has_text_frame:
        contact_info = (
            "본사: 전북특별자치도 완주군 삼례읍 삼례로 443 우석대학교 창업보육센터 302호\n"
            "연락처: 063-290-1000  |  이메일: contact@farmdriver.co.kr\n"
            "웹사이트: www.farmdriver.co.kr  |  IR 총괄: 이한규 대표이사"
        )
        p = sh.text_frame.paragraphs[0]
        update_text_preserving_runs(p, contact_info)
        # Clear extra paragraphs in this textbox
        for p_extra in sh.text_frame.paragraphs[1:]:
            p_extra.text = ""
        print("Fixed S30 Contact info")
    if sh.name == "TextBox 20" and sh.has_text_frame:
        update_text_preserving_runs(sh.text_frame.paragraphs[0], "“대한민국 농촌의 일손을 잇고, 농기계의 새로운 가치를 창출합니다.”")

# Process general shapes across all slides
for idx, slide in enumerate(prs.slides):
    for shape in slide.shapes:
        process_shape(shape)

# Slide 27 special table
for shape in prs.slides[26].shapes:
    if shape.has_table:
        table = shape.table
        if len(table.rows) > 2 and len(table.rows[2].cells) > 0:
            if table.rows[2].cells[0].text.strip() == "마케팅":
                update_text_preserving_runs(table.rows[2].cells[0].text_frame.paragraphs[0], "R&D / 플랫폼")
            if len(table.rows[2].cells) > 1 and "30%" in table.rows[2].cells[1].text:
                update_text_preserving_runs(table.rows[2].cells[1].text_frame.paragraphs[0], "30% (9억 원)")

prs.save(temp_stage_pptx)
print("Saved transformed intermediate PPTX to:", temp_stage_pptx)

# -------------------------------------------------------------------------
# Step 3. Packaging & In-Place Media Replacement into Final Output PPTX
# -------------------------------------------------------------------------
# Extract temp_stage_pptx, replace media files, clean xml relations & props, repackage to final_output_pptx
unzip_dir = r"C:\Users\note\vd\vd15_slide-style-cloner\temp_unzip"
if os.path.exists(unzip_dir):
    shutil.rmtree(unzip_dir)
os.makedirs(unzip_dir, exist_ok=True)

with zipfile.ZipFile(temp_stage_pptx, 'r') as z:
    z.extractall(unzip_dir)

# Overwrite Media Files with Custom High-Res Assets
media_dest = os.path.join(unzip_dir, "ppt", "media")
for mf in os.listdir(MEDIA_DIR):
    src = os.path.join(MEDIA_DIR, mf)
    dst = os.path.join(media_dest, mf)
    if os.path.exists(dst):
        shutil.copyfile(src, dst)
        print(f"Replaced media asset: {mf}")

# Clean Hyperlink relation in slide1.xml.rels
s1_rels_path = os.path.join(unzip_dir, "ppt", "slides", "_rels", "slide1.xml.rels")
if os.path.exists(s1_rels_path):
    with open(s1_rels_path, 'r', encoding='utf-8') as f:
        xml_rels = f.read()
    # Remove rId3 external hyperlink to yesform
    if 'plan.yesform.com' in xml_rels:
        import re
        xml_rels = re.sub(r'<Relationship[^>]*Target="https?://[^"]*yesform\.com[^"]*"[^>]*/>', '', xml_rels)
        with open(s1_rels_path, 'w', encoding='utf-8') as f:
            f.write(xml_rels)
        print("Cleaned slide1.xml.rels external link")

# Clean docProps metadata
core_xml_path = os.path.join(unzip_dir, "docProps", "core.xml")
if os.path.exists(core_xml_path):
    with open(core_xml_path, 'r', encoding='utf-8') as f:
        core_c = f.read()
    core_c = core_c.replace("YESFORM by BR.YOON", "(주)팜드라이버 IR팀")
    core_c = core_c.replace("Yesform", "FarmDriver")
    core_c = core_c.replace("http://powerpoint.yesform.com/", "https://www.farmdriver.co.kr")
    with open(core_xml_path, 'w', encoding='utf-8') as f:
        f.write(core_c)
    print("Cleaned docProps/core.xml")

app_xml_path = os.path.join(unzip_dir, "docProps", "app.xml")
if os.path.exists(app_xml_path):
    with open(app_xml_path, 'r', encoding='utf-8') as f:
        app_c = f.read()
    app_c = app_c.replace("YESFORM Co.,Ltd.", "(주)팜드라이버")
    with open(app_xml_path, 'w', encoding='utf-8') as f:
        f.write(app_c)
    print("Cleaned docProps/app.xml")

# Repackage to final PPTX
with zipfile.ZipFile(final_output_pptx, 'w', zipfile.ZIP_DEFLATED) as z_out:
    for root, dirs, files in os.walk(unzip_dir):
        for file in files:
            full_path = os.path.join(root, file)
            rel_path = os.path.relpath(full_path, unzip_dir)
            z_out.write(full_path, rel_path)

# Also sync to vd12 if directory exists
vd12_target = r"C:\Users\note\vd\vd12_농기계대리운전\result\팜드라이버_농기계대리운전_IR투자제안서.pptx"
if os.path.exists(os.path.dirname(vd12_target)):
    shutil.copyfile(final_output_pptx, vd12_target)
    print("Synced to vd12 target directory")

# Cleanup temp directories
shutil.rmtree(unzip_dir)
os.remove(temp_stage_pptx)

final_size = os.path.getsize(final_output_pptx)
print(f"\n[MASTER SUCCESS] Masterpiece Deck Generated!")
print(f"Path: {final_output_pptx}")
print(f"File Size: {final_size:,} bytes ({final_size / (1024*1024):.2f} MB)")
