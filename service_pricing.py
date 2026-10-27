"""Approved tariffs, shared by server quotes and all delivery paths."""
VERSION = '2026-10-26-v2'
STANDARD = 10
DESIGNER = 20
CHAT = 0.5
VISION = VOICE = 3
WORD_PAGE = 5
WORD_MINIMUM = 20
GENERATED_WORD = frozenset({'doc','referat','report','essay','coursework','notes'})
DETAILED_WORD = frozenset({'dkp_car','dkp','rent','rent_flat_ru','services','services_ru','employment','supply','gift','marriage_contract','lawsuit','alimony','partnership','nda','agency','joint_activity','nonresidential_rent','cession','self_employed'})
EXCEL = {'expense_estimate':20,'price_list':20,'family_budget':20,'calc_table':20,'project_budget':40,'startup_model':40}

def word_price(pages, kind='doc'):
    if kind in DETAILED_WORD: return 40
    if kind not in GENERATED_WORD: return 20
    try: count=max(1,min(40,int(pages)))
    except (ValueError,TypeError): count=1
    return max(WORD_MINIMUM,count*WORD_PAGE)

def excel_price(kind):
    return EXCEL.get(kind,20)

def public_tariffs():
    return dict(free_daily=10,free_reset_timezone="Europe/Moscow",version=VERSION,standard=STANDARD,designer=DESIGNER,chat=CHAT,vision=VISION,voice=VOICE,
                word_page=WORD_PAGE,word_minimum=WORD_MINIMUM,word_detailed=sorted(DETAILED_WORD),
                word_generated=sorted(GENERATED_WORD),excel=EXCEL,
                constructor=dict(base=10,extra_photo=3,pattern=2,maximum=20))
