#!/usr/bin/env python3
"""
عرض مقترح فعالية غداء العمل – مكتب العوفي والحربي
Working Lunch Event Proposal – Al-Awfi & Al-Harbi
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

# ── Brand Colors (Navy + Gold – Professional Audit Firm) ─────────────
NAVY    = RGBColor(0x1B, 0x3A, 0x6B)   # Primary Navy Blue
NAVY2   = RGBColor(0x0D, 0x1F, 0x42)   # Dark Navy
GOLD    = RGBColor(0xC4, 0x8E, 0x1A)   # Primary Gold
GOLD_L  = RGBColor(0xF5, 0xEB, 0xCE)   # Light Gold
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY   = RGBColor(0xF1, 0xF4, 0xF9)   # Light Gray-Blue
MGRAY   = RGBColor(0x9A, 0xA8, 0xC2)   # Medium Gray
DARK    = RGBColor(0x14, 0x1A, 0x2E)   # Near Black
TEAL    = RGBColor(0x1B, 0x7A, 0x9E)   # Accent Teal

FONT = "Arial"

# ── Helpers ───────────────────────────────────────────────────────────
def _set_cs(run):
    rPr = run._r.get_or_add_rPr()
    for e in rPr.findall(qn('a:cs')):
        rPr.remove(e)
    cs = etree.SubElement(rPr, qn('a:cs'))
    cs.set('typeface', FONT)

def R(slide, l, t, w, h, c):
    """Add filled rectangle (in inches, no border)."""
    s = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = c
    s.line.fill.background()
    return s

def T(slide, text, l, t, w, h,
      sz=16, bold=False, col=DARK,
      align=PP_ALIGN.RIGHT, rtl=True, italic=False):
    """Add single-paragraph textbox."""
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    if rtl:
        pPr = p._p.get_or_add_pPr()
        pPr.set(qn('a:rtl'), '1')
    r = p.add_run()
    r.text = text
    r.font.size = Pt(sz)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = col
    r.font.name = FONT
    _set_cs(r)
    return tb

def ML(slide, lines, l, t, w, h,
       sz=15, bold=False, col=DARK,
       align=PP_ALIGN.RIGHT, sp_a=5,
       prefix='', rtl=True,
       hd=None, hd_sz=None):
    """Add multi-line textbox with optional header line."""
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    all_lines  = ([hd] + list(lines)) if hd else list(lines)
    all_sz     = ([hd_sz or sz+2] + [sz]*len(lines)) if hd else [sz]*len(lines)
    all_bold   = ([True] + [bold]*len(lines)) if hd else [bold]*len(lines)
    all_prefix = ([''] + [prefix]*len(lines)) if hd else [prefix]*len(lines)

    for i, (line, lsz, lb, lpfx) in enumerate(
            zip(all_lines, all_sz, all_bold, all_prefix)):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(sp_a)
        if rtl:
            pPr = p._p.get_or_add_pPr()
            pPr.set(qn('a:rtl'), '1')
        r = p.add_run()
        r.text = (lpfx + ' ' if lpfx else '') + line
        r.font.size = Pt(lsz)
        r.font.bold = lb
        r.font.color.rgb = col
        r.font.name = FONT
        _set_cs(r)
    return tb

def footer(slide):
    T(slide,
      "شركة العوفي والحربي محاسبون ومراجعون قانونيون  |  www.cla.sa",
      0.3, 7.17, 12.73, 0.28,
      sz=11, col=WHITE, align=PP_ALIGN.CENTER, rtl=False)

def std_header(slide, title_ar, title_en):
    R(slide, 0, 0, 13.33, 7.5, LGRAY)
    R(slide, 0, 0, 13.33, 1.45, NAVY)
    R(slide, 0, 0, 0.15, 7.5, GOLD)
    R(slide, 0.15, 1.45, 13.18, 0.07, GOLD)
    R(slide, 0, 7.15, 13.33, 0.35, NAVY)
    T(slide, title_ar, 0.35, 0.12, 12.5, 1.0, sz=30, bold=True, col=WHITE)
    T(slide, title_en, 0.35, 0.82, 6,   0.5,  sz=13, col=GOLD,
      align=PP_ALIGN.LEFT, rtl=False)
    footer(slide)

# ══════════════════════════════════════════════════════════════════════
# SESSION DATA – 10 Dialogue Sessions
# ══════════════════════════════════════════════════════════════════════
SESSIONS = [
    {
        "num": 1, "day": 1,
        "title": "دور المراجعة الداخلية في تعزيز حوكمة الشركات",
        "axes": [
            "مفهوم الحوكمة المؤسسية ومتطلباتها وفق الأنظمة السعودية",
            "دور المراجع الداخلي في دعم وتعزيز بيئة الحوكمة",
            "العلاقة التكاملية بين المراجعة الداخلية ومجلس الإدارة",
            "أفضل الممارسات العالمية والتجارب الناجحة في حوكمة المراجعة",
        ],
    },
    {
        "num": 2, "day": 1,
        "title": "معايير المراجعة الداخلية الدولية (IPPF 2024) والبيئة السعودية",
        "axes": [
            "نظرة عامة على إطار IPPF 2024 المُحدَّث وأبرز المستجدات",
            "متطلبات الامتثال للمعايير الدولية وآليات التطبيق المحلي",
            "المقارنة بين الإطار القديم والجديد وأهم التغييرات الجوهرية",
            "آليات تطبيق المعايير بكفاءة في المنشآت السعودية",
        ],
    },
    {
        "num": 3, "day": 1,
        "title": "إدارة المخاطر المؤسسية وأثرها على خطة المراجعة الداخلية",
        "axes": [
            "إطار إدارة المخاطر المؤسسية (ERM) وركائزه الأساسية",
            "تحديد وتقييم المخاطر الجوهرية في المنشآت الكبرى والمتوسطة",
            "بناء خطة المراجعة السنوية على أساس التحليل الدقيق للمخاطر",
            "الربط الفعّال بين خريطة المخاطر وأولويات المراجعة الداخلية",
        ],
    },
    {
        "num": 4, "day": 1,
        "title": "التحول الرقمي وأثره على مهنة المراجعة الداخلية",
        "axes": [
            "الذكاء الاصطناعي وتحليل البيانات الضخمة في عمليات المراجعة",
            "أدوات المراجعة الرقمية الحديثة وتطبيقاتها العملية",
            "مراجعة أمن المعلومات وحماية البنية التحتية للأنظمة",
            "الفرص والتحديات التي يطرحها التحول الرقمي على المراجع الداخلي",
        ],
    },
    {
        "num": 5, "day": 1,
        "title": "الاستقلالية والموضوعية وأخلاقيات المراجعة الداخلية",
        "axes": [
            "ميثاق الأخلاقيات المهنية للمراجع الداخلي وفق معايير IIA",
            "الاستقلالية التنظيمية والموضوعية الفردية وضماناتها",
            "إدارة حالات تعارض المصالح ومعالجتها بمهنية عالية",
            "دور الثقافة المؤسسية في تعزيز الأخلاقيات والنزاهة المهنية",
        ],
    },
    {
        "num": 6, "day": 2,
        "title": "المراجعة الداخلية القائمة على المخاطر – المنهجية والتطبيق",
        "axes": [
            "منهجية المراجعة القائمة على المخاطر خطوةً بخطوة",
            "بناء خطة مهمة المراجعة انطلاقاً من تقييم المخاطر",
            "تحديد حجم العينة والإجراءات بناءً على درجة المخاطر",
            "صياغة تقارير المراجعة القائمة على المخاطر وفق المعايير",
        ],
    },
    {
        "num": 7, "day": 2,
        "title": "دور المراجعة الداخلية في مكافحة الاحتيال والفساد المالي",
        "axes": [
            "أساليب الكشف المبكر عن الغش والاحتيال في المنشآت",
            "إطار الوقاية من الاحتيال ودور المراجع الداخلي المحوري",
            "إجراءات التحقيق الداخلي عند الاشتباه بحالات الاحتيال",
            "الامتثال لأنظمة مكافحة الاحتيال وغسيل الأموال",
        ],
    },
    {
        "num": 8, "day": 2,
        "title": "تقييم فاعلية الرقابة الداخلية وفق إطار COSO",
        "axes": [
            "مكونات إطار COSO للرقابة الداخلية ومبادئه الأساسية الخمسة",
            "منهجيات تقييم فاعلية بيئة الرقابة الداخلية بشكل متكامل",
            "تحديد نقاط الضعف الجوهرية والإخفاقات الرقابية ومعالجتها",
            "بناء خطط العلاج الفعّالة ومتابعة تنفيذ التوصيات",
        ],
    },
    {
        "num": 9, "day": 2,
        "title": "مستقبل مهنة المراجعة الداخلية في ظل رؤية 2030",
        "axes": [
            "الاتجاهات العالمية الراهنة لمهنة المراجعة الداخلية وتطورها",
            "أثر رؤية 2030 والإصلاحات الاقتصادية الكبرى على مسيرة المهنة",
            "متطلبات سوق العمل السعودي للمراجع الداخلي الكفء المستقبلي",
            "مسارات تطوير الكفاءات المهنية والمسار الوظيفي للمراجع الداخلي",
        ],
    },
    {
        "num": 10, "day": 2,
        "title": "التكامل بين المراجعة الداخلية والخارجية ودور لجنة المراجعة",
        "axes": [
            "الفوارق والتكامل المنهجي بين المراجعة الداخلية والخارجية",
            "آليات التنسيق والتعاون الفعّال بين الفريقين في الميدان",
            "دور لجنة المراجعة في الإشراف والحوكمة وتفعيل المساءلة",
            "نماذج تطبيقية ناجحة في تعزيز التنسيق والعمل المشترك",
        ],
    },
]

# ══════════════════════════════════════════════════════════════════════
# BUILD PRESENTATION
# ══════════════════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)
BL = prs.slide_layouts[6]   # Blank layout


# ── SLIDE 1: Cover ────────────────────────────────────────────────────
def add_cover():
    sl = prs.slides.add_slide(BL)

    # Background layers
    R(sl, 0, 0, 13.33, 7.5, NAVY2)
    R(sl, 0, 0, 13.33, 0.13, GOLD)          # top gold strip
    R(sl, 0, 7.37, 13.33, 0.13, GOLD)       # bottom gold strip
    R(sl, 12.8, 0.13, 0.18, 7.24, GOLD)     # right gold bar
    R(sl, 0, 0.13, 0.18, 7.24, GOLD)        # left gold bar

    # Inner panel
    R(sl, 0.35, 0.45, 12.63, 6.65, NAVY)

    # Decorative gold corner accent (top-right quadrant)
    R(sl, 9.5, 0.45, 3.48, 0.08, GOLD)
    R(sl, 12.98, 0.45, 0.0, 2.5, GOLD)

    # Gold horizontal divider
    R(sl, 1.2, 3.8, 10.93, 0.05, GOLD)

    # "Proposal" badge
    R(sl, 4.9, 1.05, 3.53, 0.55, GOLD)
    T(sl, "عـرض مقـترح", 4.9, 1.05, 3.53, 0.55,
      sz=17, bold=True, col=NAVY2, align=PP_ALIGN.CENTER)

    # Main title
    T(sl, "فعـالية غـداء العمل",
      0.55, 1.75, 12.23, 1.2,
      sz=44, bold=True, col=WHITE, align=PP_ALIGN.CENTER)

    # Subtitle
    T(sl, "في مجال المراجعة الداخلية",
      0.55, 2.85, 12.23, 0.75,
      sz=26, bold=False, col=GOLD, align=PP_ALIGN.CENTER)

    # Organizer label
    T(sl, "تُقيمها وترعاها:", 0.55, 3.95, 12.23, 0.45,
      sz=15, col=MGRAY, align=PP_ALIGN.CENTER)

    # Office name
    T(sl, "شركة العوفي والحربي  محاسبون ومراجعون قانونيون",
      0.35, 4.38, 12.63, 0.75,
      sz=28, bold=True, col=WHITE, align=PP_ALIGN.CENTER)

    # CLA badge
    T(sl, "عضو  CLA Global",
      0.55, 5.1, 12.23, 0.45,
      sz=16, col=GOLD, align=PP_ALIGN.CENTER, rtl=False)

    # Details row
    R(sl, 2.0, 5.7, 9.33, 0.52, NAVY2)
    T(sl, "فعالية من يومين   |   10 جلسات حوارية   |   70 – 100 مشارك / يوم",
      2.0, 5.7, 9.33, 0.52,
      sz=14, col=WHITE, align=PP_ALIGN.CENTER, rtl=False)

    # Website
    T(sl, "www.cla.sa",
      5.7, 6.45, 1.93, 0.38,
      sz=14, col=GOLD, align=PP_ALIGN.CENTER, rtl=False)

add_cover()


# ── SLIDE 2: Event Overview ───────────────────────────────────────────
def add_overview():
    sl = prs.slides.add_slide(BL)
    std_header(sl, "نظرة عامة على الفعالية", "Event Overview")

    # 3 KPI cards
    cards = [
        ("مدة الفعالية",     "يومان",       "Day 1  &  Day 2"),
        ("عدد الجلسات",      "10 جلسات",    "5 جلسات في كل يوم"),
        ("المشاركون / يوم",  "70 – 100",    "شخص متوقع"),
    ]
    cw, ch, cy = 3.75, 2.1, 1.65
    for i, (ttl, val, sub) in enumerate(cards):
        cx = 0.4 + i * (cw + 0.32)
        R(sl, cx, cy, cw, ch, WHITE)
        R(sl, cx, cy, cw, 0.09, GOLD)
        T(sl, ttl, cx, cy+0.15, cw, 0.45,
          sz=15, bold=True, col=NAVY, align=PP_ALIGN.CENTER)
        T(sl, val, cx, cy+0.58, cw, 0.75,
          sz=30, bold=True, col=GOLD, align=PP_ALIGN.CENTER)
        T(sl, sub, cx, cy+1.3, cw, 0.55,
          sz=13, col=MGRAY, align=PP_ALIGN.CENTER)

    # Description box
    R(sl, 0.4, 3.95, 12.53, 3.0, WHITE)
    R(sl, 0.4, 3.95, 12.53, 0.07, GOLD)
    T(sl, "عن الفعالية", 0.6, 4.08, 4, 0.44,
      sz=17, bold=True, col=NAVY)

    desc = [
        "تهدف الفعالية إلى استقطاب النخبة من المهنيين والمختصين في مجال المراجعة الداخلية "
        "في بيئة احترافية تفاعلية تجمع الفكر مع العمل.",
        "تشتمل الفعالية على عشر جلسات حوارية متخصصة تغطي أبرز محاور وتحديات المهنة، "
        "موزعة بالتساوي على يومين متتاليين.",
        "تُدار كل جلسة بإشراف مدير حوار متخصص من مكتب العوفي والحربي، بمشاركة راعٍ "
        "مؤسسي ومتحدثَين مستقلَّين لإثراء النقاش.",
        "يُتاح للرعاة والمتحدثين تعزيز حضورهم المهني وتوسيع شبكة علاقاتهم في الوسط المهني.",
    ]
    ML(sl, desc,
       0.55, 4.6, 12.3, 2.3,
       sz=13, col=DARK, sp_a=7)

add_overview()


# ── SLIDE 3: Event Structure ──────────────────────────────────────────
def add_structure():
    sl = prs.slides.add_slide(BL)
    std_header(sl, "هيكل الفعالية", "Event Structure")

    day_sessions = [
        [
            "الجلسة 1  –  المراجعة الداخلية والحوكمة",
            "الجلسة 2  –  معايير المراجعة الدولية IPPF",
            "الجلسة 3  –  إدارة المخاطر المؤسسية",
            "الجلسة 4  –  التحول الرقمي والمراجعة",
            "الجلسة 5  –  الأخلاقيات والاستقلالية",
        ],
        [
            "الجلسة 6  –  المراجعة القائمة على المخاطر",
            "الجلسة 7  –  مكافحة الاحتيال والفساد",
            "الجلسة 8  –  تقييم الرقابة الداخلية COSO",
            "الجلسة 9  –  مستقبل المهنة ورؤية 2030",
            "الجلسة 10 –  تكامل المراجعة الداخلية والخارجية",
        ],
    ]
    day_labels = ["اليوم الأول", "اليوم الثاني"]

    col_x = [0.35, 6.84]
    col_w = 6.3

    for d, (day_x, sessions_list, label) in enumerate(
            zip(col_x, day_sessions, day_labels)):
        R(sl, day_x, 1.65, col_w, 5.35, WHITE)
        R(sl, day_x, 1.65, col_w, 0.58, NAVY)
        T(sl, label, day_x, 1.65, col_w, 0.58,
          sz=19, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        for i, sess in enumerate(sessions_list):
            ry = 2.33 + i * 0.91
            bg = LGRAY if i % 2 == 0 else WHITE
            R(sl, day_x+0.05, ry, col_w-0.1, 0.82, bg)
            # number badge
            bx = day_x + col_w - 0.55
            R(sl, bx, ry+0.16, 0.42, 0.5, GOLD)
            num = i + 1 + (d * 5)
            T(sl, str(num), bx, ry+0.16, 0.42, 0.5,
              sz=13, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
            T(sl, sess, day_x+0.15, ry+0.12, col_w-0.85, 0.6,
              sz=12, col=DARK)

    # Middle arrow/divider
    R(sl, 6.52, 3.1, 0.29, 0.06, GOLD)
    R(sl, 6.52, 4.1, 0.29, 0.06, GOLD)
    R(sl, 6.52, 5.1, 0.29, 0.06, GOLD)

add_structure()


# ── SLIDES 4-13: Session Slides ───────────────────────────────────────
def add_session(s):
    sl = prs.slides.add_slide(BL)

    # Background
    R(sl, 0, 0, 13.33, 7.5, LGRAY)
    # Top header
    R(sl, 0, 0, 13.33, 1.7, NAVY)
    # Left gold bar
    R(sl, 0, 0, 0.15, 7.5, GOLD)
    # Gold separator line
    R(sl, 0.15, 1.7, 13.18, 0.07, GOLD)
    # Footer
    R(sl, 0, 7.15, 13.33, 0.35, NAVY)
    footer(sl)

    # Day badge
    day_ar = "اليوم الأول" if s["day"] == 1 else "اليوم الثاني"
    R(sl, 11.6, 0.12, 1.56, 0.42, GOLD)
    T(sl, day_ar, 11.6, 0.12, 1.56, 0.42,
      sz=13, bold=True, col=NAVY2, align=PP_ALIGN.CENTER)

    # Session number badge
    R(sl, 11.6, 0.62, 1.56, 0.82, GOLD_L)
    T(sl, f"الجلسة  {s['num']}",
      11.6, 0.62, 1.56, 0.82,
      sz=16, bold=True, col=NAVY, align=PP_ALIGN.CENTER)

    # Session title in header
    T(sl, s["title"],
      0.3, 0.12, 11.15, 1.5,
      sz=22, bold=True, col=WHITE)

    # ── Axes card (left / main) ─────────────────────────────────────
    R(sl, 0.25, 1.87, 8.75, 5.1, WHITE)
    R(sl, 0.25, 1.87, 8.75, 0.52, NAVY)
    T(sl, "محاور الجلسة الحوارية",
      0.38, 1.87, 8.5, 0.52,
      sz=15, bold=True, col=WHITE)

    nums_ar = ["أولاً", "ثانياً", "ثالثاً", "رابعاً"]
    for i, (ax, lbl) in enumerate(zip(s["axes"], nums_ar)):
        ay = 2.52 + i * 1.1
        if i % 2 == 1:
            R(sl, 0.25, ay-0.12, 8.75, 1.0, LGRAY)
        R(sl, 0.25, ay-0.08, 0.07, 0.9, GOLD)     # accent line
        R(sl, 0.38, ay+0.08, 0.65, 0.55, GOLD)    # number badge
        T(sl, str(i+1), 0.38, ay+0.08, 0.65, 0.55,
          sz=15, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        T(sl, lbl + " :", 1.12, ay, 1.1, 0.42,
          sz=12, bold=True, col=GOLD)
        T(sl, ax, 1.12, ay+0.38, 7.75, 0.65,
          sz=13, col=DARK)

    # ── Speakers panel (right) ──────────────────────────────────────
    R(sl, 9.15, 1.87, 4.0, 5.1, NAVY)
    R(sl, 9.15, 1.87, 4.0, 0.08, GOLD)

    T(sl, "فريق الجلسة",
      9.15, 1.97, 4.0, 0.5,
      sz=15, bold=True, col=GOLD, align=PP_ALIGN.CENTER)

    # Moderator
    R(sl, 9.25, 2.6, 3.8, 0.05, GOLD)
    T(sl, "مدير الحوار",
      9.25, 2.68, 3.8, 0.38,
      sz=13, bold=True, col=GOLD, align=PP_ALIGN.CENTER)
    T(sl, "من مكتب العوفي والحربي",
      9.25, 3.05, 3.8, 0.38,
      sz=12, col=WHITE, align=PP_ALIGN.CENTER)
    T(sl, "أ. [اسم مدير الحوار]",
      9.25, 3.38, 3.8, 0.38,
      sz=12, col=MGRAY, align=PP_ALIGN.CENTER)

    # Sponsor
    R(sl, 9.25, 3.88, 3.8, 0.05, GOLD)
    T(sl, "راعي الجلسة",
      9.25, 3.97, 3.8, 0.38,
      sz=13, bold=True, col=GOLD, align=PP_ALIGN.CENTER)
    T(sl, "[اسم الجهة الراعية]",
      9.25, 4.35, 3.8, 0.38,
      sz=12, col=WHITE, align=PP_ALIGN.CENTER)
    T(sl, "إيراد الرعاية: 10,000 ريال",
      9.25, 4.68, 3.8, 0.38,
      sz=11, col=GOLD_L, align=PP_ALIGN.CENTER)

    # Independent speakers
    R(sl, 9.25, 5.18, 3.8, 0.05, GOLD)
    T(sl, "المتحدثون المستقلون",
      9.25, 5.27, 3.8, 0.38,
      sz=13, bold=True, col=GOLD, align=PP_ALIGN.CENTER)
    for j in range(2):
        T(sl, f"م. [المتحدث المستقل  {j+1}]",
          9.25, 5.68 + j * 0.46, 3.8, 0.42,
          sz=12, col=WHITE, align=PP_ALIGN.CENTER)
    T(sl, "إيراد كل متحدث: 4,000 ريال",
      9.25, 6.62, 3.8, 0.35,
      sz=11, col=GOLD_L, align=PP_ALIGN.CENTER)

for sess in SESSIONS:
    add_session(sess)


# ── SLIDE 14: Costs ───────────────────────────────────────────────────
def add_costs():
    sl = prs.slides.add_slide(BL)
    std_header(sl, "التكاليف التقديرية للفعالية", "Estimated Event Costs")

    # Table header row
    R(sl, 0.35, 1.62, 12.63, 0.52, NAVY)
    T(sl, "البند",                  5.8, 1.62, 6.8,  0.52, sz=14, bold=True, col=WHITE)
    T(sl, "المبلغ (ريال سعودي)",    0.38, 1.62, 4.3, 0.52, sz=14, bold=True, col=WHITE)
    T(sl, "النوع",                  4.75, 1.62, 1.0, 0.52, sz=14, bold=True,
      col=WHITE, align=PP_ALIGN.CENTER)

    rows = [
        ("رخصة تصريح إقامة الفعالية",
         "",
         "5,000 ريال", "ثابت", GOLD),
        ("تكاليف مكتب الفعاليات",
         "",
         "8,000 ريال", "ثابت", GOLD),
        ("تكاليف الاستضافة – الحد الأدنى",
         "70 شخص × 280 ريال × 2 يوم",
         "39,200 ريال", "متغير", TEAL),
        ("تكاليف الاستضافة – الحد الأعلى",
         "100 شخص × 350 ريال × 2 يوم",
         "70,000 ريال", "متغير", TEAL),
    ]

    for i, (item, sub, amount, kind, badge_c) in enumerate(rows):
        y = 2.25 + i * 1.02
        bg = WHITE if i % 2 == 0 else LGRAY
        R(sl, 0.35, y, 12.63, 0.92, bg)
        T(sl, item,   5.8, y+0.05, 6.8, 0.45, sz=13, bold=True, col=DARK)
        if sub:
            T(sl, sub, 5.8, y+0.48, 6.8, 0.38, sz=11, col=MGRAY)
        T(sl, amount, 0.45, y+0.2,  4.0, 0.52, sz=17, bold=True, col=NAVY)
        R(sl, 4.75, y+0.22, 1.0, 0.45, badge_c)
        T(sl, kind, 4.75, y+0.22, 1.0, 0.45,
          sz=12, bold=True, col=WHITE, align=PP_ALIGN.CENTER)

    # Total row
    R(sl, 0.35, 6.35, 12.63, 0.62, NAVY)
    T(sl, "إجمالي التكاليف المتوقعة",
      3.5, 6.35, 9.1, 0.62, sz=16, bold=True, col=WHITE)
    T(sl, "52,200 – 83,000 ريال",
      0.45, 6.35, 3.0, 0.62, sz=16, bold=True, col=GOLD)

    # Note
    T(sl,
      "* التكاليف المتغيرة محسوبة بناءً على توقع حضور 70 – 100 شخص / يوم × يومين",
      0.35, 7.02, 12.63, 0.38,
      sz=11, col=MGRAY)

add_costs()


# ── SLIDE 15: Revenue ─────────────────────────────────────────────────
def add_revenue():
    sl = prs.slides.add_slide(BL)
    std_header(sl, "الإيرادات المتوقعة للفعالية", "Projected Revenue")

    rev_items = [
        ("رعاة الجلسات الحوارية",
         "10 جلسات  ×  10,000 ريال / جلسة",
         "100,000 ريال"),
        ("المتحدثون المستقلون",
         "10 جلسات  ×  متحدثان  ×  4,000 ريال / متحدث",
         "80,000 ريال"),
    ]

    for i, (source, calc, total) in enumerate(rev_items):
        y = 1.72 + i * 1.7
        R(sl, 0.35, y, 12.63, 1.5, WHITE)
        R(sl, 0.35, y, 0.1,  1.5, GOLD)
        T(sl, source, 0.6, y+0.12, 8.5, 0.55, sz=19, bold=True, col=NAVY)
        T(sl, calc,   0.6, y+0.72, 8.5, 0.55, sz=13, col=MGRAY)
        R(sl, 9.7, y+0.32, 3.0, 0.78, GOLD)
        T(sl, total, 9.7, y+0.32, 3.0, 0.78,
          sz=20, bold=True, col=WHITE, align=PP_ALIGN.CENTER)

    # Total revenue
    R(sl, 0.35, 5.2, 12.63, 0.75, NAVY)
    T(sl, "إجمالي الإيرادات المتوقعة",
      3.0, 5.2, 9.5, 0.75, sz=18, bold=True, col=WHITE)
    T(sl, "180,000 ريال",
      0.45, 5.2, 2.5, 0.75, sz=20, bold=True, col=GOLD)

    # Net surplus
    R(sl, 0.35, 6.1, 12.63, 0.87, GOLD_L)
    R(sl, 0.35, 6.1, 0.12, 0.87, GOLD)
    T(sl, "صافي الفائض المتوقع (بعد خصم التكاليف)",
      2.3, 6.1, 9.3, 0.45, sz=15, bold=True, col=NAVY)
    T(sl, "97,000  –  127,800 ريال سعودي",
      2.3, 6.5, 9.3, 0.4,  sz=14, col=DARK)
    T(sl, "النطاق",
      0.5, 6.3, 1.5, 0.45, sz=13, bold=True, col=GOLD, align=PP_ALIGN.CENTER)

    T(sl,
      "* الأرقام تقديرية وقابلة للمراجعة وفقاً لعدد المشاركين الفعلي وشروط التفاوض مع الرعاة",
      0.35, 7.02, 12.63, 0.35,
      sz=11, col=MGRAY)

add_revenue()


# ── SLIDE 16: Closing ─────────────────────────────────────────────────
def add_closing():
    sl = prs.slides.add_slide(BL)

    R(sl, 0, 0, 13.33, 7.5, NAVY2)
    R(sl, 0, 0, 13.33, 0.13, GOLD)
    R(sl, 0, 7.37, 13.33, 0.13, GOLD)
    R(sl, 0, 0.13, 0.18, 7.24, GOLD)
    R(sl, 12.8, 0.13, 0.18, 7.24, GOLD)

    # Inner content panel
    R(sl, 0.35, 0.5, 12.63, 6.65, NAVY)

    # Decorative gold lines
    R(sl, 1.5, 2.1, 10.33, 0.05, GOLD)
    R(sl, 1.5, 5.3, 10.33, 0.05, GOLD)

    T(sl, "شـكراً لاهتمامكم",
      0.5, 1.0, 12.33, 1.25,
      sz=46, bold=True, col=WHITE, align=PP_ALIGN.CENTER)

    T(sl, "نتطلع إلى شراكة ناجحة وفعالية متميزة",
      0.5, 2.28, 12.33, 0.72,
      sz=22, col=GOLD, align=PP_ALIGN.CENTER)

    T(sl, "شركة العوفي والحربي  محاسبون ومراجعون قانونيون",
      0.5, 3.2, 12.33, 0.68,
      sz=22, bold=True, col=WHITE, align=PP_ALIGN.CENTER)

    T(sl, "عضو  CLA Global",
      0.5, 3.85, 12.33, 0.55,
      sz=17, col=GOLD, align=PP_ALIGN.CENTER, rtl=False)

    T(sl, "www.cla.sa",
      5.7, 4.55, 1.93, 0.45,
      sz=16, col=GOLD, align=PP_ALIGN.CENTER, rtl=False)

    # Contact prompt
    T(sl, "للتواصل والاستفسار يُرجى زيارة الموقع الإلكتروني",
      0.5, 5.5, 12.33, 0.5,
      sz=15, col=MGRAY, align=PP_ALIGN.CENTER)

add_closing()


# ── Save ──────────────────────────────────────────────────────────────
OUTPUT = "/home/user/Abdulsalam-/عرض_مقترح_فعالية_غداء_العمل.pptx"
prs.save(OUTPUT)
print(f"✅  Saved: {OUTPUT}")
print(f"    Slides: {len(prs.slides)}")
