#!/usr/bin/env python3
"""
Big-Four-style presentation – navy + white only
Creates two files: Arabic and English versions
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import copy

# ── Colors: Navy + White only ─────────────────────────────────────────
NAVY     = RGBColor(0x00, 0x30, 0x87)   # Deep Navy (primary)
NAVY2    = RGBColor(0x00, 0x1A, 0x55)   # Darker Navy
NAVY_M   = RGBColor(0x1A, 0x52, 0xA8)   # Mid Navy (accent)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
OFF_W    = RGBColor(0xF5, 0xF7, 0xFB)   # Off-white bg
GRAY_L   = RGBColor(0xE8, 0xEC, 0xF4)   # Light gray
GRAY_M   = RGBColor(0x8A, 0x94, 0xAE)   # Medium gray (secondary text)
DARK     = RGBColor(0x0D, 0x1A, 0x33)   # Near-black text

FONT_AR = "Arial"
FONT_EN = "Calibri"

# ── Low-level helpers ─────────────────────────────────────────────────
def _cs(run, font):
    rPr = run._r.get_or_add_rPr()
    for e in rPr.findall(qn('a:cs')): rPr.remove(e)
    cs = etree.SubElement(rPr, qn('a:cs'))
    cs.set('typeface', font)

def rect(slide, l, t, w, h, c, border_c=None, border_w=None):
    s = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = c
    if border_c:
        s.line.color.rgb = border_c
        if border_w: s.line.width = Pt(border_w)
    else:
        s.line.fill.background()
    return s

def oval(slide, l, t, w, h, c):
    s = slide.shapes.add_shape(9, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = c
    s.line.fill.background()
    return s

def tb(slide, text, l, t, w, h, sz=16, bold=False, col=DARK,
       align=PP_ALIGN.RIGHT, rtl=True, italic=False, font=FONT_AR, wrap=True):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    if rtl:
        pPr = p._p.get_or_add_pPr(); pPr.set(qn('a:rtl'), '1')
    r = p.add_run()
    r.text = text; r.font.size = Pt(sz); r.font.bold = bold
    r.font.italic = italic; r.font.color.rgb = col; r.font.name = font
    _cs(r, font)
    return box

def mltb(slide, lines, l, t, w, h, sz=14, bold=False, col=DARK,
         align=PP_ALIGN.RIGHT, sp_a=5, prefix='', rtl=True, font=FONT_AR,
         header=None, hdr_sz=None, hdr_bold=True, hdr_col=None):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = True
    all_l = ([header] + list(lines)) if header else list(lines)
    all_s = ([hdr_sz or sz+2] + [sz]*len(lines)) if header else [sz]*len(lines)
    all_b = ([hdr_bold] + [bold]*len(lines)) if header else [bold]*len(lines)
    all_c = ([(hdr_col or col)] + [col]*len(lines)) if header else [col]*len(lines)
    all_p = ([''] + [prefix]*len(lines)) if header else [prefix]*len(lines)
    for i, (line, s, b, c, pfx) in enumerate(zip(all_l, all_s, all_b, all_c, all_p)):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_after = Pt(sp_a)
        if rtl:
            pPr = p._p.get_or_add_pPr(); pPr.set(qn('a:rtl'), '1')
        r = p.add_run()
        r.text = (pfx + '  ' if pfx else '') + line
        r.font.size = Pt(s); r.font.bold = b; r.font.color.rgb = c
        r.font.name = font; _cs(r, font)
    return box

# ══════════════════════════════════════════════════════════════════════
# CONTENT DATA
# ══════════════════════════════════════════════════════════════════════
DATA = {
    "ar": {
        "lang": "ar",
        "font": FONT_AR,
        "rtl": True,
        "align_main": PP_ALIGN.RIGHT,
        "align_c": PP_ALIGN.CENTER,
        "align_l": PP_ALIGN.LEFT,
        "event_title": "فعالية غداء العمل",
        "event_sub": "في مجال المراجعة الداخلية",
        "proposal_tag": "عرض مقترح",
        "org_name": "شركة العوفي والحربي",
        "org_sub": "محاسبون ومراجعون قانونيون  •  عضو CLA Global",
        "web": "www.cla.sa",
        "kpi": ["يومان", "10 جلسات", "70 – 100"],
        "kpi_lbl": ["مدة الفعالية", "جلسات حوارية", "مشارك / يوم"],
        "overview_title": "نظرة عامة",
        "structure_title": "هيكل الفعالية",
        "costs_title": "التكاليف التقديرية",
        "revenue_title": "الإيرادات المتوقعة",
        "closing_line1": "نتطلع إلى شراكة ناجحة",
        "closing_line2": "وفعالية متميزة",
        "day_labels": ["اليوم الأول", "اليوم الثاني"],
        "session_word": "الجلسة",
        "axes_title": "محاور الجلسة",
        "team_title": "فريق الجلسة",
        "moderator_lbl": "مدير الحوار",
        "moderator_val": "أ. [اسم مدير الحوار]\nمكتب العوفي والحربي",
        "sponsor_lbl": "راعي الجلسة",
        "sponsor_val": "[اسم الجهة الراعية]",
        "sponsor_fee": "إيراد الرعاية: 10,000 ريال",
        "speakers_lbl": "المتحدثون المستقلون",
        "speaker_val": ["م. [المتحدث الأول]", "م. [المتحدث الثاني]"],
        "speaker_fee": "4,000 ريال / متحدث",
        "day1_list": [
            "الجلسة 1  –  المراجعة الداخلية والحوكمة",
            "الجلسة 2  –  معايير المراجعة الدولية IPPF",
            "الجلسة 3  –  إدارة المخاطر المؤسسية",
            "الجلسة 4  –  التحول الرقمي والمراجعة",
            "الجلسة 5  –  الأخلاقيات والاستقلالية",
        ],
        "day2_list": [
            "الجلسة 6   –  المراجعة القائمة على المخاطر",
            "الجلسة 7   –  مكافحة الاحتيال والفساد",
            "الجلسة 8   –  تقييم الرقابة الداخلية COSO",
            "الجلسة 9   –  مستقبل المهنة ورؤية 2030",
            "الجلسة 10  –  تكامل المراجعتين الداخلية والخارجية",
        ],
        "cost_hdr": ["البند", "المبلغ (ريال سعودي)", "النوع"],
        "costs_rows": [
            ("رخصة تصريح إقامة الفعالية", "5,000 ريال", "ثابت"),
            ("تكاليف مكتب الفعاليات", "8,000 ريال", "ثابت"),
            ("استضافة – الحد الأدنى  (70 × 280 × 2)", "39,200 ريال", "متغير"),
            ("استضافة – الحد الأعلى  (100 × 350 × 2)", "70,000 ريال", "متغير"),
        ],
        "total_costs": "52,200  –  83,000 ريال",
        "total_costs_lbl": "إجمالي التكاليف المتوقعة",
        "rev_rows": [
            ("رعاة الجلسات الحوارية",
             "10 جلسات  ×  10,000 ريال",
             "100,000 ريال"),
            ("المتحدثون المستقلون",
             "10 جلسات  ×  2 متحدث  ×  4,000 ريال",
             "80,000 ريال"),
        ],
        "total_rev": "180,000 ريال",
        "total_rev_lbl": "إجمالي الإيرادات المتوقعة",
        "surplus_lbl": "صافي الفائض المتوقع",
        "surplus_val": "97,000  –  127,800 ريال",
        "cost_note": "* التكاليف المتغيرة بناءً على الحضور المتوقع 70–100 شخص/يوم × يومين",
        "rev_note": "* الأرقام تقديرية وقابلة للتعديل وفقاً للتفاوض مع الرعاة",
        "sessions": [
            {
                "num": 1, "day": 1,
                "title": "دور المراجعة الداخلية في تعزيز حوكمة الشركات",
                "axes": [
                    "مفهوم الحوكمة المؤسسية ومتطلباتها في الأنظمة السعودية",
                    "دور المراجع الداخلي في دعم وتعزيز بيئة الحوكمة",
                    "العلاقة التكاملية بين المراجعة الداخلية ومجلس الإدارة",
                    "أفضل الممارسات والتجارب الناجحة في حوكمة المراجعة",
                ],
            },
            {
                "num": 2, "day": 1,
                "title": "معايير المراجعة الداخلية الدولية (IPPF 2024) والبيئة السعودية",
                "axes": [
                    "نظرة عامة على إطار IPPF 2024 وأبرز المستجدات",
                    "متطلبات الامتثال وآليات التطبيق في البيئة المحلية",
                    "المقارنة بين الإطار القديم والجديد وأهم التغييرات",
                    "تطبيق المعايير بكفاءة في المنشآت السعودية",
                ],
            },
            {
                "num": 3, "day": 1,
                "title": "إدارة المخاطر المؤسسية وأثرها على خطة المراجعة الداخلية",
                "axes": [
                    "إطار إدارة المخاطر المؤسسية (ERM) وركائزه الأساسية",
                    "تحديد وتقييم المخاطر الجوهرية في المنشآت",
                    "بناء خطة المراجعة السنوية على أساس المخاطر",
                    "الربط الفعّال بين خريطة المخاطر وأولويات المراجعة",
                ],
            },
            {
                "num": 4, "day": 1,
                "title": "التحول الرقمي وأثره على مهنة المراجعة الداخلية",
                "axes": [
                    "الذكاء الاصطناعي وتحليل البيانات الضخمة في المراجعة",
                    "أدوات المراجعة الرقمية الحديثة وتطبيقاتها العملية",
                    "مراجعة أمن المعلومات وحماية البنية التحتية للأنظمة",
                    "الفرص والتحديات التي يطرحها التحول الرقمي",
                ],
            },
            {
                "num": 5, "day": 1,
                "title": "الاستقلالية والموضوعية وأخلاقيات المراجعة الداخلية",
                "axes": [
                    "ميثاق الأخلاقيات المهنية للمراجع الداخلي وفق معايير IIA",
                    "الاستقلالية التنظيمية والموضوعية الفردية وضماناتها",
                    "إدارة حالات تعارض المصالح ومعالجتها مهنياً",
                    "دور الثقافة المؤسسية في تعزيز الأخلاقيات والنزاهة",
                ],
            },
            {
                "num": 6, "day": 2,
                "title": "المراجعة الداخلية القائمة على المخاطر – المنهجية والتطبيق",
                "axes": [
                    "منهجية المراجعة القائمة على المخاطر خطوةً بخطوة",
                    "بناء خطة مهمة المراجعة انطلاقاً من تقييم المخاطر",
                    "تحديد حجم العينة والإجراءات بناءً على درجة المخاطر",
                    "صياغة تقارير المراجعة القائمة على المخاطر",
                ],
            },
            {
                "num": 7, "day": 2,
                "title": "دور المراجعة الداخلية في مكافحة الاحتيال والفساد المالي",
                "axes": [
                    "أساليب الكشف المبكر عن الغش والاحتيال في المنشآت",
                    "إطار الوقاية من الاحتيال ودور المراجع الداخلي",
                    "إجراءات التحقيق الداخلي عند الاشتباه بالاحتيال",
                    "الامتثال لأنظمة مكافحة الاحتيال وغسيل الأموال",
                ],
            },
            {
                "num": 8, "day": 2,
                "title": "تقييم فاعلية الرقابة الداخلية وفق إطار COSO",
                "axes": [
                    "مكونات إطار COSO للرقابة الداخلية ومبادئه الخمسة",
                    "منهجيات تقييم فاعلية بيئة الرقابة الداخلية",
                    "تحديد نقاط الضعف الجوهرية والإخفاقات الرقابية",
                    "بناء خطط العلاج ومتابعة تنفيذ التوصيات",
                ],
            },
            {
                "num": 9, "day": 2,
                "title": "مستقبل مهنة المراجعة الداخلية في ظل رؤية 2030",
                "axes": [
                    "الاتجاهات العالمية لمهنة المراجعة الداخلية وتطورها",
                    "أثر رؤية 2030 والإصلاحات الاقتصادية على المهنة",
                    "متطلبات سوق العمل السعودي للمراجع الداخلي المستقبلي",
                    "مسارات تطوير الكفاءات المهنية والمسار الوظيفي",
                ],
            },
            {
                "num": 10, "day": 2,
                "title": "التكامل بين المراجعة الداخلية والخارجية ودور لجنة المراجعة",
                "axes": [
                    "الفوارق والتكامل المنهجي بين المراجعة الداخلية والخارجية",
                    "آليات التنسيق والتعاون الفعّال بين الفريقين",
                    "دور لجنة المراجعة في الإشراف والحوكمة والمساءلة",
                    "نماذج تطبيقية ناجحة في تعزيز التنسيق المشترك",
                ],
            },
        ],
    },
    "en": {
        "lang": "en",
        "font": FONT_EN,
        "rtl": False,
        "align_main": PP_ALIGN.LEFT,
        "align_c": PP_ALIGN.CENTER,
        "align_l": PP_ALIGN.LEFT,
        "event_title": "Working Lunch Event",
        "event_sub": "Internal Audit Dialogue Sessions",
        "proposal_tag": "PROPOSAL",
        "org_name": "Al-Awfi & Al-Harbi",
        "org_sub": "Certified Public Accountants & Legal Auditors  •  CLA Global Member",
        "web": "www.cla.sa",
        "kpi": ["2 Days", "10 Sessions", "70 – 100"],
        "kpi_lbl": ["Event Duration", "Dialogue Sessions", "Attendees / Day"],
        "overview_title": "Event Overview",
        "structure_title": "Event Structure",
        "costs_title": "Estimated Costs",
        "revenue_title": "Projected Revenue",
        "closing_line1": "We look forward to a",
        "closing_line2": "successful partnership",
        "day_labels": ["Day One", "Day Two"],
        "session_word": "Session",
        "axes_title": "Session Agenda",
        "team_title": "Session Team",
        "moderator_lbl": "Dialogue Moderator",
        "moderator_val": "[Moderator Name]\nAl-Awfi & Al-Harbi Office",
        "sponsor_lbl": "Session Sponsor",
        "sponsor_val": "[Sponsor Organization]",
        "sponsor_fee": "Sponsorship Fee: SAR 10,000",
        "speakers_lbl": "Independent Speakers",
        "speaker_val": ["[Speaker 1]", "[Speaker 2]"],
        "speaker_fee": "SAR 4,000 / speaker",
        "day1_list": [
            "Session 1  –  Internal Audit & Corporate Governance",
            "Session 2  –  International Standards IPPF 2024",
            "Session 3  –  Enterprise Risk Management",
            "Session 4  –  Digital Transformation & Audit",
            "Session 5  –  Ethics & Independence",
        ],
        "day2_list": [
            "Session 6   –  Risk-Based Internal Auditing",
            "Session 7   –  Fraud & Financial Corruption",
            "Session 8   –  Internal Control Evaluation (COSO)",
            "Session 9   –  Future of the Profession & Vision 2030",
            "Session 10  –  Internal & External Audit Integration",
        ],
        "cost_hdr": ["Item", "Amount (SAR)", "Type"],
        "costs_rows": [
            ("Event Permit License", "SAR 5,000", "Fixed"),
            ("Event Office Fees", "SAR 8,000", "Fixed"),
            ("Hospitality – Minimum  (70 × 280 × 2)", "SAR 39,200", "Variable"),
            ("Hospitality – Maximum  (100 × 350 × 2)", "SAR 70,000", "Variable"),
        ],
        "total_costs": "SAR 52,200  –  83,000",
        "total_costs_lbl": "Total Estimated Costs",
        "rev_rows": [
            ("Session Sponsors",
             "10 sessions  ×  SAR 10,000",
             "SAR 100,000"),
            ("Independent Speakers",
             "10 sessions  ×  2 speakers  ×  SAR 4,000",
             "SAR 80,000"),
        ],
        "total_rev": "SAR 180,000",
        "total_rev_lbl": "Total Projected Revenue",
        "surplus_lbl": "Expected Net Surplus",
        "surplus_val": "SAR 97,000  –  127,800",
        "cost_note": "* Variable costs based on expected attendance of 70–100 per day × 2 days",
        "rev_note": "* Figures are estimates subject to negotiation with sponsors",
        "sessions": [
            {
                "num": 1, "day": 1,
                "title": "The Role of Internal Audit in Enhancing Corporate Governance",
                "axes": [
                    "Concept of corporate governance and Saudi regulatory requirements",
                    "The internal auditor's role in supporting governance frameworks",
                    "The complementary relationship between internal audit and the board",
                    "Global best practices in internal audit governance",
                ],
            },
            {
                "num": 2, "day": 1,
                "title": "International Internal Auditing Standards (IPPF 2024) in the Saudi Context",
                "axes": [
                    "Overview of the updated IPPF 2024 framework and key changes",
                    "Compliance requirements and local implementation approaches",
                    "Comparison between the old and new frameworks",
                    "Implementing standards efficiently within Saudi organisations",
                ],
            },
            {
                "num": 3, "day": 1,
                "title": "Enterprise Risk Management and Its Impact on the Internal Audit Plan",
                "axes": [
                    "Enterprise Risk Management (ERM) framework and core pillars",
                    "Identifying and assessing material risks in organisations",
                    "Building the annual audit plan on a risk-assessment basis",
                    "Aligning the risk map with internal audit priorities",
                ],
            },
            {
                "num": 4, "day": 1,
                "title": "Digital Transformation and Its Effect on the Internal Audit Profession",
                "axes": [
                    "Artificial intelligence and big data analytics in auditing",
                    "Modern digital audit tools and their practical applications",
                    "Information security auditing and infrastructure protection",
                    "Opportunities and challenges of digital transformation for auditors",
                ],
            },
            {
                "num": 5, "day": 1,
                "title": "Independence, Objectivity, and Ethics in Internal Auditing",
                "axes": [
                    "IIA Code of Ethics and professional conduct standards",
                    "Organisational independence and individual objectivity",
                    "Managing and resolving conflicts of interest professionally",
                    "The role of organisational culture in promoting integrity",
                ],
            },
            {
                "num": 6, "day": 2,
                "title": "Risk-Based Internal Auditing – Methodology and Application",
                "axes": [
                    "Step-by-step methodology for risk-based internal auditing",
                    "Building the engagement plan based on risk assessment",
                    "Determining sample size and procedures by risk level",
                    "Writing risk-based audit reports in line with standards",
                ],
            },
            {
                "num": 7, "day": 2,
                "title": "The Role of Internal Audit in Combating Fraud and Financial Corruption",
                "axes": [
                    "Early detection techniques for fraud and misconduct",
                    "Fraud prevention framework and the internal auditor's role",
                    "Internal investigation procedures when fraud is suspected",
                    "Compliance with anti-fraud and anti-money-laundering regulations",
                ],
            },
            {
                "num": 8, "day": 2,
                "title": "Evaluating Internal Control Effectiveness Based on the COSO Framework",
                "axes": [
                    "COSO internal control framework: five components and principles",
                    "Methodologies for assessing the effectiveness of control environments",
                    "Identifying material weaknesses and control deficiencies",
                    "Developing remediation plans and tracking recommendation implementation",
                ],
            },
            {
                "num": 9, "day": 2,
                "title": "The Future of the Internal Audit Profession Under Vision 2030",
                "axes": [
                    "Global trends shaping the internal audit profession",
                    "Impact of Vision 2030 and economic reforms on the profession",
                    "Saudi labour market requirements for future internal auditors",
                    "Professional development pathways and career progression",
                ],
            },
            {
                "num": 10, "day": 2,
                "title": "Integration Between Internal & External Audit and the Audit Committee",
                "axes": [
                    "Differences and synergies between internal and external audit",
                    "Mechanisms for effective coordination between both teams",
                    "The audit committee's role in oversight and governance",
                    "Successful models of collaboration and joint working",
                ],
            },
        ],
    },
}

# ══════════════════════════════════════════════════════════════════════
# SLIDE BUILDER  (Big-Four / fluid style)
# ══════════════════════════════════════════════════════════════════════

def build(d):
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    BL = prs.slide_layouts[6]
    F  = d["font"]
    AR = d["rtl"]
    AM = d["align_main"]
    AC = d["align_c"]
    AL = d["align_l"]

    # ── helpers scoped to language ────────────────────────────────────
    def T(slide, text, l, t, w, h, sz=16, bold=False, col=DARK,
          align=None, italic=False):
        align = align if align is not None else AM
        return tb(slide, text, l, t, w, h, sz, bold, col, align, AR, italic, F)

    def ML(slide, lines, l, t, w, h, sz=14, bold=False, col=DARK,
           align=None, sp_a=6, prefix='',
           header=None, hdr_sz=None, hdr_col=None):
        align = align if align is not None else AM
        return mltb(slide, lines, l, t, w, h, sz, bold, col, align, sp_a,
                    prefix, AR, F, header, hdr_sz, True, hdr_col)

    def foot(slide):
        rect(slide, 0, 7.2, 13.33, 0.3, NAVY)
        tb(slide, d["org_name"] + "   |   " + d["web"],
           0.4, 7.22, 12.53, 0.26,
           sz=11, col=WHITE, align=AC, rtl=False, font=F)

    # ── diagonal accent helper (parallelogram) ────────────────────────
    def diag_accent(slide, l, t, w, h, c):
        s = slide.shapes.add_shape(20, Inches(l), Inches(t), Inches(w), Inches(h))
        s.fill.solid(); s.fill.fore_color.rgb = c
        s.line.fill.background()
        return s

    # ─────────────────────────────────────────────────────────────────
    # SLIDE 1 – COVER  (full-bleed left panel split)
    # ─────────────────────────────────────────────────────────────────
    def add_cover():
        sl = prs.slides.add_slide(BL)
        # Full bg off-white
        rect(sl, 0, 0, 13.33, 7.5, OFF_W)
        # Left navy panel (55%)
        rect(sl, 0, 0, 7.3, 7.5, NAVY2)
        # Diagonal right edge of navy panel (parallelogram overlap)
        diag_accent(sl, 6.4, 0, 1.5, 7.5, NAVY2)
        # Top accent thin line (full)
        rect(sl, 0, 0, 13.33, 0.06, NAVY_M)
        # Mid Navy stripe inside navy panel (decorative)
        rect(sl, 0, 5.9, 7.3, 0.08, NAVY_M)

        # Proposal tag (small white rounded pill on navy)
        rect(sl, 0.55, 0.65, 2.2, 0.42, NAVY_M)
        tb(sl, d["proposal_tag"], 0.55, 0.65, 2.2, 0.42,
           sz=13, bold=True, col=WHITE, align=AC, rtl=False, font=F)

        # Main event title (large, white, on navy)
        tb(sl, d["event_title"], 0.45, 1.3, 6.6, 1.5,
           sz=46, bold=True, col=WHITE, align=AL, rtl=AR, font=F)
        # Subtitle
        tb(sl, d["event_sub"], 0.45, 2.8, 6.6, 0.65,
           sz=20, bold=False, col=GRAY_L, align=AL, rtl=AR, font=F)

        # Thin white divider
        rect(sl, 0.45, 3.62, 5.5, 0.04, WHITE)

        # Org name
        tb(sl, d["org_name"], 0.45, 3.82, 6.6, 0.68,
           sz=22, bold=True, col=WHITE, align=AL, rtl=AR, font=F)
        tb(sl, d["org_sub"], 0.45, 4.48, 6.6, 0.45,
           sz=12, col=GRAY_M, align=AL, rtl=False, font=F)

        # Web on navy
        tb(sl, d["web"], 0.45, 6.7, 3, 0.42,
           sz=13, col=GRAY_L, align=AL, rtl=False, font=F)

        # RIGHT side: KPI tiles (3 stacked)
        kpi_data = list(zip(d["kpi"], d["kpi_lbl"]))
        for i, (val, lbl) in enumerate(kpi_data):
            ky = 1.0 + i * 1.85
            rect(sl, 8.3, ky, 4.5, 1.55, WHITE)
            # Left navy accent bar
            rect(sl, 8.3, ky, 0.08, 1.55, NAVY)
            tb(sl, val, 8.5, ky+0.12, 4.2, 0.75,
               sz=36, bold=True, col=NAVY2, align=AL, rtl=False, font=F)
            tb(sl, lbl, 8.5, ky+0.88, 4.2, 0.45,
               sz=14, col=GRAY_M, align=AL, rtl=False, font=F)

    add_cover()

    # ─────────────────────────────────────────────────────────────────
    # SLIDE 2 – OVERVIEW
    # ─────────────────────────────────────────────────────────────────
    def add_overview():
        sl = prs.slides.add_slide(BL)
        rect(sl, 0, 0, 13.33, 7.5, WHITE)
        # Top navy bar
        rect(sl, 0, 0, 13.33, 1.3, NAVY)
        # Diagonal tab
        diag_accent(sl, 3.6, 0, 1.0, 1.3, NAVY2)
        # Bottom footer
        foot(sl)

        T(sl, d["overview_title"], 0.45, 0.18, 9, 0.92,
          sz=30, bold=True, col=WHITE, align=AL)

        # 3 KPI boxes
        for i, (val, lbl) in enumerate(zip(d["kpi"], d["kpi_lbl"])):
            bx = 0.4 + i * 4.27
            rect(sl, bx, 1.5, 4.0, 2.1, OFF_W)
            rect(sl, bx, 1.5, 4.0, 0.07, NAVY)
            tb(sl, val, bx, 1.65, 4.0, 0.9,
               sz=36, bold=True, col=NAVY2, align=AC, rtl=False, font=F)
            tb(sl, lbl, bx, 2.52, 4.0, 0.5,
               sz=13, col=GRAY_M, align=AC, rtl=False, font=F)

        # Description
        rect(sl, 0.4, 3.85, 12.53, 3.0, OFF_W)
        rect(sl, 0.4, 3.85, 0.07, 3.0, NAVY)
        desc_lines = {
            "ar": [
                "تهدف الفعالية إلى استقطاب النخبة من المهنيين والمختصين في مجال المراجعة الداخلية في بيئة احترافية تفاعلية.",
                "تشتمل على عشر جلسات حوارية متخصصة توزّع بالتساوي على يومين متتاليين تغطي أبرز محاور المهنة وتحدياتها.",
                "تُدار كل جلسة بمدير حوار متخصص من مكتب العوفي والحربي مع راعٍ مؤسسي ومتحدثَين مستقلَّين.",
                "تمنح الرعاة والمتحدثين فرصة تعزيز حضورهم المهني وتوسيع شبكة علاقاتهم في الوسط المهني.",
            ],
            "en": [
                "The event aims to bring together leading professionals and specialists in internal auditing in a dynamic, interactive setting.",
                "It comprises ten specialised dialogue sessions spread equally over two consecutive days, covering key professional topics.",
                "Each session is moderated by a specialist from Al-Awfi & Al-Harbi, alongside a corporate sponsor and two independent speakers.",
                "Sponsors and speakers gain a platform to strengthen their professional profile and expand their network.",
            ],
        }
        ML(sl, desc_lines[d["lang"]], 0.65, 4.05, 12.1, 2.65,
           sz=14, col=DARK, sp_a=9)

    add_overview()

    # ─────────────────────────────────────────────────────────────────
    # SLIDE 3 – STRUCTURE
    # ─────────────────────────────────────────────────────────────────
    def add_structure():
        sl = prs.slides.add_slide(BL)
        rect(sl, 0, 0, 13.33, 7.5, WHITE)
        rect(sl, 0, 0, 13.33, 1.3, NAVY)
        diag_accent(sl, 3.6, 0, 1.0, 1.3, NAVY2)
        foot(sl)
        T(sl, d["structure_title"], 0.45, 0.18, 9, 0.92,
          sz=30, bold=True, col=WHITE, align=AL)

        cols = [
            (0.35,  d["day_labels"][0], d["day1_list"], 0),
            (6.84,  d["day_labels"][1], d["day2_list"], 5),
        ]
        for cx, day_lbl, sess_list, offset in cols:
            cw = 6.3
            # Header card
            rect(sl, cx, 1.48, cw, 0.6, NAVY)
            tb(sl, day_lbl, cx, 1.48, cw, 0.6,
               sz=17, bold=True, col=WHITE, align=AC, rtl=False, font=F)
            for i, name in enumerate(sess_list):
                ry = 2.18 + i * 0.98
                bg = OFF_W if i % 2 == 0 else WHITE
                rect(sl, cx, ry, cw, 0.88, bg)
                # num badge
                nx = cx + cw - 0.6
                rect(sl, nx, ry+0.19, 0.48, 0.48, NAVY)
                tb(sl, str(i + 1 + offset), nx, ry+0.19, 0.48, 0.48,
                   sz=13, bold=True, col=WHITE, align=AC, rtl=False, font=F)
                tb(sl, name, cx+0.12, ry+0.16, cw-0.85, 0.56,
                   sz=11, col=DARK, align=AL, rtl=False, font=F)

    add_structure()

    # ─────────────────────────────────────────────────────────────────
    # SLIDES 4-13 – INDIVIDUAL SESSION SLIDES
    # ─────────────────────────────────────────────────────────────────
    def add_session(s):
        sl = prs.slides.add_slide(BL)
        rect(sl, 0, 0, 13.33, 7.5, WHITE)

        # Full-width thin top stripe
        rect(sl, 0, 0, 13.33, 0.07, NAVY)

        # Left navy sidebar (narrow, full height)
        rect(sl, 0, 0, 0.6, 7.5, NAVY2)

        # Session number block (on sidebar, centred vertically top area)
        rect(sl, 0.0, 0.6, 0.6, 0.6, NAVY_M)
        tb(sl, str(s["num"]), 0.0, 0.6, 0.6, 0.6,
           sz=20, bold=True, col=WHITE, align=AC, rtl=False, font=F)

        # Day tag (on sidebar below number)
        day_short = "D1" if s["day"] == 1 else "D2"
        tb(sl, day_short, 0.0, 1.28, 0.6, 0.38,
           sz=11, bold=True, col=GRAY_L, align=AC, rtl=False, font=F)

        # Bottom footer
        foot(sl)

        # Title area (below top stripe)
        rect(sl, 0.75, 0.12, 12.4, 1.28, OFF_W)
        # Session word tag
        sess_tag = d["session_word"] + "  " + str(s["num"])
        tb(sl, sess_tag, 0.85, 0.15, 3.5, 0.38,
           sz=11, bold=True, col=NAVY_M, align=AL, rtl=False, font=F)
        day_full = d["day_labels"][s["day"] - 1]
        tb(sl, day_full, 0.85 + (9.5 if not AR else 0), 0.15, 3.5, 0.38,
           sz=11, col=GRAY_M, align=AL, rtl=False, font=F)
        # Main title
        tb(sl, s["title"], 0.85, 0.5, 12.0, 0.9,
           sz=22, bold=True, col=NAVY2, align=AM, rtl=AR, font=F)

        # Thin navy divider under title
        rect(sl, 0.75, 1.45, 12.4, 0.04, NAVY)

        # ── LEFT/MAIN AREA: Axes ─────────────────────────────────────
        # (70% of content area)
        rect(sl, 0.75, 1.55, 8.55, 5.45, WHITE)
        # Axes header
        tb(sl, d["axes_title"], 0.9, 1.6, 4, 0.42,
           sz=13, bold=True, col=NAVY, align=AL, rtl=False, font=F)
        rect(sl, 0.75, 2.08, 8.55, 0.04, GRAY_L)

        num_labels_ar = ["أولاً", "ثانياً", "ثالثاً", "رابعاً"]
        num_labels_en = ["First", "Second", "Third", "Fourth"]
        num_labels = num_labels_ar if AR else num_labels_en

        for i, (axis, lbl) in enumerate(zip(s["axes"], num_labels)):
            ay = 2.18 + i * 1.27
            # Alternating row bg
            if i % 2 == 1:
                rect(sl, 0.75, ay-0.08, 8.55, 1.18, OFF_W)
            # Vertical navy accent
            rect(sl, 0.75, ay-0.04, 0.05, 1.1, NAVY_M)
            # Number pill
            rect(sl, 0.88, ay+0.2, 0.9, 0.38, NAVY)
            tb(sl, lbl, 0.88, ay+0.2, 0.9, 0.38,
               sz=11, bold=True, col=WHITE, align=AC, rtl=False, font=F)
            # Axis text
            tb(sl, axis, 1.88, ay+0.08, 7.2, 0.85,
               sz=13, col=DARK, align=AM, rtl=AR, font=F)

        # ── RIGHT PANEL: Speakers ─────────────────────────────────────
        rect(sl, 9.45, 1.55, 3.73, 5.45, OFF_W)
        rect(sl, 9.45, 1.55, 3.73, 0.06, NAVY)

        tb(sl, d["team_title"], 9.45, 1.62, 3.73, 0.45,
           sz=13, bold=True, col=NAVY2, align=AC, rtl=False, font=F)

        sections = [
            (d["moderator_lbl"], d["moderator_val"],  None),
            (d["sponsor_lbl"],   d["sponsor_val"],    d["sponsor_fee"]),
            (d["speakers_lbl"],  "\n".join(d["speaker_val"]), d["speaker_fee"]),
        ]
        sy = 2.2
        for lbl, val, fee in sections:
            rect(sl, 9.55, sy, 3.53, 0.32, GRAY_L)
            tb(sl, lbl, 9.55, sy, 3.53, 0.32,
               sz=11, bold=True, col=NAVY2, align=AC, rtl=False, font=F)
            sy += 0.36
            tb(sl, val, 9.55, sy, 3.53, 0.55 if '\n' not in val else 0.72,
               sz=11, col=DARK, align=AC, rtl=False, font=F)
            sy += 0.6 if '\n' not in val else 0.78
            if fee:
                tb(sl, fee, 9.55, sy, 3.53, 0.32,
                   sz=10, col=NAVY_M, align=AC, rtl=False, font=F)
                sy += 0.38
            sy += 0.12

    for s in d["sessions"]:
        add_session(s)

    # ─────────────────────────────────────────────────────────────────
    # SLIDE 14 – COSTS
    # ─────────────────────────────────────────────────────────────────
    def add_costs():
        sl = prs.slides.add_slide(BL)
        rect(sl, 0, 0, 13.33, 7.5, WHITE)
        rect(sl, 0, 0, 13.33, 1.3, NAVY)
        diag_accent(sl, 3.6, 0, 1.0, 1.3, NAVY2)
        foot(sl)
        T(sl, d["costs_title"], 0.45, 0.18, 9, 0.92,
          sz=30, bold=True, col=WHITE, align=AL)

        # Column headers
        rect(sl, 0.35, 1.45, 12.63, 0.5, NAVY2)
        col_x = [0.45, 4.6, 10.2]
        for i, hdr in enumerate(d["cost_hdr"]):
            tb(sl, hdr, col_x[i], 1.45, 5.5, 0.5,
               sz=13, bold=True, col=WHITE, align=AL, rtl=False, font=F)

        for i, (item, amount, kind) in enumerate(d["costs_rows"]):
            ry = 2.05 + i * 1.08
            bg = OFF_W if i % 2 == 0 else WHITE
            rect(sl, 0.35, ry, 12.63, 0.98, bg)
            rect(sl, 0.35, ry, 0.06, 0.98, NAVY_M)
            tb(sl, item,   0.55,  ry+0.22, 9.5, 0.55, sz=13, col=DARK,
               align=AL, rtl=False, font=F)
            tb(sl, amount, 10.25, ry+0.22, 2.5, 0.55, sz=15, bold=True,
               col=NAVY, align=AL, rtl=False, font=F)

        # Total
        rect(sl, 0.35, 6.35, 12.63, 0.65, NAVY)
        tb(sl, d["total_costs_lbl"], 0.55, 6.35, 8.0, 0.65,
           sz=15, bold=True, col=WHITE, align=AL, rtl=False, font=F)
        tb(sl, d["total_costs"],     9.5,  6.35, 3.3, 0.65,
           sz=15, bold=True, col=GRAY_L, align=AL, rtl=False, font=F)

        tb(sl, d["cost_note"], 0.45, 7.08, 12, 0.3,
           sz=10, col=GRAY_M, align=AL, rtl=False, font=F)

    add_costs()

    # ─────────────────────────────────────────────────────────────────
    # SLIDE 15 – REVENUE
    # ─────────────────────────────────────────────────────────────────
    def add_revenue():
        sl = prs.slides.add_slide(BL)
        rect(sl, 0, 0, 13.33, 7.5, WHITE)
        rect(sl, 0, 0, 13.33, 1.3, NAVY)
        diag_accent(sl, 3.6, 0, 1.0, 1.3, NAVY2)
        foot(sl)
        T(sl, d["revenue_title"], 0.45, 0.18, 9, 0.92,
          sz=30, bold=True, col=WHITE, align=AL)

        for i, (src, calc, total) in enumerate(d["rev_rows"]):
            ry = 1.5 + i * 1.85
            rect(sl, 0.35, ry, 12.63, 1.65, OFF_W if i%2==0 else WHITE)
            rect(sl, 0.35, ry, 0.07, 1.65, NAVY)
            tb(sl, src,   0.6, ry+0.15, 8.8, 0.58, sz=18, bold=True,
               col=NAVY2, align=AL, rtl=False, font=F)
            tb(sl, calc,  0.6, ry+0.75, 8.8, 0.45, sz=13, col=GRAY_M,
               align=AL, rtl=False, font=F)
            # Amount box
            rect(sl, 10.0, ry+0.35, 2.8, 0.8, NAVY)
            tb(sl, total, 10.0, ry+0.35, 2.8, 0.8,
               sz=18, bold=True, col=WHITE, align=AC, rtl=False, font=F)

        # Total row
        rect(sl, 0.35, 5.3, 12.63, 0.72, NAVY2)
        tb(sl, d["total_rev_lbl"], 0.55, 5.3, 8.5, 0.72,
           sz=16, bold=True, col=WHITE, align=AL, rtl=False, font=F)
        tb(sl, d["total_rev"],     9.6,  5.3, 3.0, 0.72,
           sz=18, bold=True, col=GRAY_L, align=AL, rtl=False, font=F)

        # Surplus box
        rect(sl, 0.35, 6.18, 12.63, 0.78, OFF_W)
        rect(sl, 0.35, 6.18, 0.07,  0.78, NAVY_M)
        tb(sl, d["surplus_lbl"],  0.6, 6.18, 6.0, 0.78,
           sz=15, bold=True, col=NAVY2, align=AL, rtl=False, font=F)
        tb(sl, d["surplus_val"], 7.0, 6.18, 5.8, 0.78,
           sz=15, bold=True, col=NAVY, align=AL, rtl=False, font=F)

        tb(sl, d["rev_note"], 0.45, 7.08, 12, 0.3,
           sz=10, col=GRAY_M, align=AL, rtl=False, font=F)

    add_revenue()

    # ─────────────────────────────────────────────────────────────────
    # SLIDE 16 – CLOSING
    # ─────────────────────────────────────────────────────────────────
    def add_closing():
        sl = prs.slides.add_slide(BL)
        rect(sl, 0, 0, 13.33, 7.5, NAVY2)
        rect(sl, 0, 0, 13.33, 0.07, NAVY_M)
        rect(sl, 0, 7.43, 13.33, 0.07, NAVY_M)

        # Right white panel
        rect(sl, 6.8, 0, 6.53, 7.5, WHITE)
        diag_accent(sl, 5.6, 0, 1.6, 7.5, WHITE)

        # Left panel text
        tb(sl, d["closing_line1"], 0.5, 2.3, 6.0, 0.85,
           sz=32, bold=True, col=WHITE, align=AL, rtl=AR, font=F)
        tb(sl, d["closing_line2"], 0.5, 3.1, 6.0, 0.85,
           sz=32, bold=True, col=GRAY_L, align=AL, rtl=AR, font=F)
        rect(sl, 0.5, 4.05, 4.5, 0.05, NAVY_M)
        tb(sl, d["org_name"], 0.5, 4.25, 6.0, 0.6,
           sz=18, bold=True, col=WHITE, align=AL, rtl=AR, font=F)
        tb(sl, d["org_sub"], 0.5, 4.82, 6.0, 0.45,
           sz=11, col=GRAY_M, align=AL, rtl=False, font=F)
        tb(sl, d["web"], 0.5, 5.55, 3.0, 0.4,
           sz=14, col=GRAY_L, align=AL, rtl=False, font=F)

        # Right panel: big session stats
        tb(sl, "10", 8.5, 1.5, 3.5, 1.5,
           sz=90, bold=True, col=NAVY2, align=AC, rtl=False, font=F)
        tb(sl, d["kpi_lbl"][1], 8.5, 2.9, 3.5, 0.55,
           sz=16, col=GRAY_M, align=AC, rtl=False, font=F)
        rect(sl, 9.2, 3.65, 2.1, 0.05, GRAY_L)
        tb(sl, "2", 8.5, 3.9, 3.5, 1.1,
           sz=70, bold=True, col=NAVY2, align=AC, rtl=False, font=F)
        tb(sl, d["kpi_lbl"][0], 8.5, 4.85, 3.5, 0.5,
           sz=16, col=GRAY_M, align=AC, rtl=False, font=F)

    add_closing()

    return prs

# ── Build both versions ───────────────────────────────────────────────
for lang, suffix in [("ar", "عرض_مقترح_فعالية_غداء_العمل_v2"),
                      ("en", "Working_Lunch_Event_Proposal_v2")]:
    p = build(DATA[lang])
    out = f"/home/user/Abdulsalam-/{suffix}.pptx"
    p.save(out)
    print(f"✅  {out}  ({len(p.slides)} slides)")
