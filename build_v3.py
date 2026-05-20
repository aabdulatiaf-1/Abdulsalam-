#!/usr/bin/env python3
"""
Big-Four-style presentation v3
• Arabic font : Traditional Arabic (formal)
• English font: Georgia (serif)
• No diagonal / parallelogram shapes
• Added TOC slide
• Added Introduction + Benefits slides
• 5 agenda items per session
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

# ── Colors: Navy + White ──────────────────────────────────────────────
NAVY   = RGBColor(0x00, 0x30, 0x87)
NAVY2  = RGBColor(0x00, 0x1A, 0x55)
NAVY_M = RGBColor(0x1A, 0x52, 0xA8)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
OFF_W  = RGBColor(0xF5, 0xF7, 0xFB)
GRAY_L = RGBColor(0xE8, 0xEC, 0xF4)
GRAY_M = RGBColor(0x8A, 0x94, 0xAE)
DARK   = RGBColor(0x0D, 0x1A, 0x33)

FONT_AR = "Traditional Arabic"
FONT_EN = "Georgia"

# ── Core helpers ──────────────────────────────────────────────────────
def _cs(run, font):
    rPr = run._r.get_or_add_rPr()
    for e in rPr.findall(qn('a:cs')): rPr.remove(e)
    cs = etree.SubElement(rPr, qn('a:cs'))
    cs.set('typeface', font)

def R(sl, l, t, w, h, c):
    s = sl.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = c
    s.line.fill.background()
    return s

def tb(sl, text, l, t, w, h, sz=16, bold=False, col=DARK,
       align=PP_ALIGN.LEFT, rtl=False, italic=False, font=FONT_EN, wrap=True):
    box = sl.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    if rtl:
        pPr = p._p.get_or_add_pPr(); pPr.set(qn('a:rtl'), '1')
    r = p.add_run()
    r.text = text; r.font.size = Pt(sz); r.font.bold = bold
    r.font.italic = italic; r.font.color.rgb = col; r.font.name = font
    _cs(r, font)
    return box

def mltb(sl, lines, l, t, w, h, sz=14, bold=False, col=DARK,
         align=PP_ALIGN.LEFT, sp_a=5, rtl=False, font=FONT_EN,
         hdr=None, hdr_sz=None, hdr_col=None):
    box = sl.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = True
    all_l = ([hdr] + list(lines)) if hdr else list(lines)
    all_s = ([hdr_sz or sz+2] + [sz]*len(lines)) if hdr else [sz]*len(lines)
    all_b = ([True] + [bold]*len(lines)) if hdr else [bold]*len(lines)
    all_c = ([(hdr_col or col)] + [col]*len(lines)) if hdr else [col]*len(lines)
    for i, (line, s, b, c) in enumerate(zip(all_l, all_s, all_b, all_c)):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_after = Pt(sp_a)
        if rtl:
            pPr = p._p.get_or_add_pPr(); pPr.set(qn('a:rtl'), '1')
        r = p.add_run()
        r.text = line; r.font.size = Pt(s); r.font.bold = b
        r.font.color.rgb = c; r.font.name = font; _cs(r, font)
    return box

# ══════════════════════════════════════════════════════════════════════
# CONTENT
# ══════════════════════════════════════════════════════════════════════
DATA = {
  # ── ARABIC ─────────────────────────────────────────────────────────
  "ar": {
    "lang":"ar", "font":FONT_AR, "rtl":True,
    "am":PP_ALIGN.RIGHT, "ac":PP_ALIGN.CENTER, "al":PP_ALIGN.LEFT,
    "event_title":   "فعالية غداء العمل",
    "event_sub":     "جلسات حوارية متخصصة في مجال المراجعة الداخلية",
    "proposal_tag":  "عرض مقترح",
    "org_name":      "شركة العوفي والحربي",
    "org_sub":       "محاسبون ومراجعون قانونيون  •  عضو CLA Global",
    "web":           "www.cla.sa",
    "kpi":           ["يومان", "10", "70–100"],
    "kpi_lbl":       ["مدة الفعالية", "جلسة حوارية", "مشارك / يوم"],
    "toc_title":     "الفهرس",
    "toc_items": [
      ("01", "مقدمة الفعالية"),
      ("02", "مزايا الاجتماع والفئة المستهدفة"),
      ("03", "هيكل الفعالية"),
      ("04", "الجلسة الأولى  –  المراجعة الداخلية والحوكمة"),
      ("05", "الجلسة الثانية  –  معايير IPPF 2024"),
      ("06", "الجلسة الثالثة  –  إدارة المخاطر المؤسسية"),
      ("07", "الجلسة الرابعة  –  التحول الرقمي"),
      ("08", "الجلسة الخامسة  –  الأخلاقيات والاستقلالية"),
      ("09", "الجلسة السادسة  –  المراجعة القائمة على المخاطر"),
      ("10", "الجلسة السابعة  –  مكافحة الاحتيال"),
      ("11", "الجلسة الثامنة  –  تقييم الرقابة الداخلية COSO"),
      ("12", "الجلسة التاسعة  –  مستقبل المهنة ورؤية 2030"),
      ("13", "الجلسة العاشرة  –  تكامل المراجعتين"),
      ("14", "التكاليف التقديرية"),
      ("15", "الإيرادات المتوقعة"),
    ],
    "intro_title":   "مقدمة عن الفعالية",
    "intro_body": [
      "فعالية غداء العمل في مجال المراجعة الداخلية هي تجمّع مهني راقٍ تُقيمه شركة العوفي والحربي "
      "لمحاسبون ومراجعون قانونيون بوصفها منظّماً ورعيةً رئيسيةً، في بيئة احترافية تجمع بين الطرح "
      "الأكاديمي والتطبيق العملي.",
      "تُقام الفعالية على مدى يومين متتاليين وتضمّ عشر جلسات حوارية متخصصة، كل جلسة تتناول "
      "موضوعاً محورياً في مهنة المراجعة الداخلية بأسلوب تفاعلي ومُثري.",
      "تُدار كل جلسة بمدير حوار متخصص من مكتب العوفي والحربي، ويُشارك فيها راعٍ مؤسسي واثنان "
      "من المتحدثين المستقلين ذوي الكفاءة والخبرة المعمّقة في المجال.",
    ],
    "benefits_title": "مزايا الاجتماع",
    "benefits_cats": [
      {
        "cat":   "مزايا للحضور والمشاركين",
        "items": [
          "اكتساب معرفة معمّقة بأحدث ممارسات ومعايير المراجعة الداخلية",
          "بناء شبكة علاقات مهنية مع النخبة من المختصين والمراجعين",
          "مناقشة التحديات العملية واستقاء الحلول من المتخصصين",
          "الحصول على ساعات التطوير المهني المستمر (CPE)",
          "الاطلاع على أفضل الممارسات المحلية والعالمية في المهنة",
        ],
      },
      {
        "cat":   "مزايا للرعاة والجهات الداعمة",
        "items": [
          "تعزيز الحضور العلامي أمام صانعي القرار وكبار المختصين",
          "استهداف نخبة من المهنيين في بيئة تفاعلية متميزة",
          "تقديم المنتجات والخدمات أمام جمهور مؤهّل وذي قرار",
          "بناء شراكات استراتيجية مع مكتب العوفي والحربي",
          "تعزيز القيادة الفكرية للمؤسسة في مجال المراجعة الداخلية",
        ],
      },
      {
        "cat":   "الفئات المستهدفة",
        "items": [
          "رؤساء لجان المراجعة وأعضاء مجالس الإدارة",
          "مديرو ومحترفو المراجعة الداخلية في القطاعين العام والخاص",
          "المراجعون الخارجيون والمستشارون الماليون",
          "مسؤولو الامتثال وإدارة المخاطر المؤسسية",
          "الأكاديميون والباحثون في مجال المحاسبة والمراجعة",
        ],
      },
    ],
    "structure_title": "هيكل الفعالية",
    "costs_title":     "التكاليف التقديرية",
    "revenue_title":   "الإيرادات المتوقعة",
    "closing_l1": "نتطلع إلى شراكة ناجحة",
    "closing_l2": "وفعالية متميزة",
    "day_labels":      ["اليوم الأول", "اليوم الثاني"],
    "session_word":    "الجلسة",
    "agenda_title":    "أجندة الجلسة",
    "team_title":      "فريق الجلسة",
    "mod_lbl":  "مدير الحوار",
    "mod_val":  "أ. [اسم مدير الحوار]\nمكتب العوفي والحربي",
    "spon_lbl": "راعي الجلسة",
    "spon_val": "[اسم الجهة الراعية]",
    "spon_fee": "رسوم الرعاية: 10,000 ريال",
    "spkr_lbl": "المتحدثون المستقلون",
    "spkr_val": ["م. [المتحدث الأول]", "م. [المتحدث الثاني]"],
    "spkr_fee": "4,000 ريال / متحدث",
    "day1_list": [
      "الجلسة 1  –  المراجعة الداخلية والحوكمة",
      "الجلسة 2  –  معايير IPPF 2024",
      "الجلسة 3  –  إدارة المخاطر المؤسسية",
      "الجلسة 4  –  التحول الرقمي والمراجعة",
      "الجلسة 5  –  الأخلاقيات والاستقلالية",
    ],
    "day2_list": [
      "الجلسة 6   –  المراجعة القائمة على المخاطر",
      "الجلسة 7   –  مكافحة الاحتيال والفساد",
      "الجلسة 8   –  تقييم الرقابة الداخلية COSO",
      "الجلسة 9   –  مستقبل المهنة ورؤية 2030",
      "الجلسة 10  –  تكامل المراجعتين",
    ],
    "cost_hdr": ["البند", "المبلغ (ريال)", "النوع"],
    "costs_rows": [
      ("رخصة تصريح إقامة الفعالية",        "5,000 ريال",  "ثابت"),
      ("تكاليف مكتب الفعاليات",             "8,000 ريال",  "ثابت"),
      ("استضافة – الحد الأدنى  (70×280×2)", "39,200 ريال", "متغير"),
      ("استضافة – الحد الأعلى  (100×350×2)","70,000 ريال", "متغير"),
    ],
    "total_costs_lbl": "إجمالي التكاليف المتوقعة",
    "total_costs":     "52,200  –  83,000 ريال",
    "cost_note":  "* التكاليف المتغيرة مبنية على توقع حضور 70–100 شخص/يوم × يومين",
    "rev_rows": [
      ("رعاة الجلسات الحوارية",    "10 جلسات × 10,000 ريال",           "100,000 ريال"),
      ("المتحدثون المستقلون",      "10 جلسات × 2 متحدث × 4,000 ريال", "80,000 ريال"),
    ],
    "total_rev_lbl": "إجمالي الإيرادات المتوقعة",
    "total_rev":     "180,000 ريال",
    "surplus_lbl":   "صافي الفائض المتوقع",
    "surplus_val":   "97,000  –  127,800 ريال",
    "rev_note": "* الأرقام تقديرية وقابلة للتعديل وفقاً لشروط التفاوض مع الرعاة",
    "num_labels": ["أولاً","ثانياً","ثالثاً","رابعاً","خامساً"],
    "sessions": [
      { "num":1,"day":1,
        "title":"دور المراجعة الداخلية في تعزيز حوكمة الشركات",
        "axes":[
          "مفهوم الحوكمة المؤسسية ومتطلباتها في الأنظمة السعودية",
          "دور المراجع الداخلي في دعم وتعزيز بيئة الحوكمة",
          "العلاقة التكاملية بين المراجعة الداخلية ومجلس الإدارة",
          "أفضل الممارسات والتجارب الدولية في حوكمة المراجعة",
          "نماذج تطبيقية من منشآت رائدة في القطاع السعودي",
        ]},
      { "num":2,"day":1,
        "title":"معايير المراجعة الداخلية الدولية (IPPF 2024) والبيئة السعودية",
        "axes":[
          "نظرة عامة على إطار IPPF 2024 وأبرز المستجدات الجوهرية",
          "متطلبات الامتثال للمعايير الدولية وآليات التطبيق المحلي",
          "المقارنة بين الإطار القديم والجديد وأهم التغييرات",
          "التحديات العملية في تطبيق المعايير بالمنشآت السعودية",
          "خارطة طريق مقترحة للانتقال إلى الإطار الجديد",
        ]},
      { "num":3,"day":1,
        "title":"إدارة المخاطر المؤسسية وأثرها على خطة المراجعة الداخلية",
        "axes":[
          "إطار إدارة المخاطر المؤسسية (ERM) وركائزه الأساسية",
          "تحديد وتقييم المخاطر الجوهرية في بيئة الأعمال السعودية",
          "بناء خطة المراجعة السنوية على أساس التحليل الدقيق للمخاطر",
          "الربط الفعّال بين خريطة المخاطر وأولويات المراجعة الداخلية",
          "قياس فاعلية الاستجابة للمخاطر وتحديث خطة المراجعة",
        ]},
      { "num":4,"day":1,
        "title":"التحول الرقمي وأثره على مهنة المراجعة الداخلية",
        "axes":[
          "الذكاء الاصطناعي وتحليل البيانات الضخمة في عمليات المراجعة",
          "أدوات المراجعة الرقمية الحديثة وتطبيقاتها العملية الميدانية",
          "مراجعة أمن المعلومات وحماية البنية التحتية للأنظمة",
          "متطلبات تطوير كفاءات المراجع الداخلي في العصر الرقمي",
          "الفرص والتحديات التي يطرحها التحول الرقمي على المهنة",
        ]},
      { "num":5,"day":1,
        "title":"الاستقلالية والموضوعية وأخلاقيات المراجعة الداخلية",
        "axes":[
          "ميثاق الأخلاقيات المهنية للمراجع الداخلي وفق معايير IIA",
          "الاستقلالية التنظيمية والموضوعية الفردية وآليات ضمانهما",
          "إدارة حالات تعارض المصالح ومعالجتها بمهنية عالية",
          "دور الثقافة المؤسسية في تعزيز الأخلاقيات والنزاهة المهنية",
          "تعزيز مكانة المراجع الداخلي وبناء الثقة مع أصحاب المصلحة",
        ]},
      { "num":6,"day":2,
        "title":"المراجعة الداخلية القائمة على المخاطر – المنهجية والتطبيق",
        "axes":[
          "منهجية المراجعة القائمة على المخاطر والإطار النظري المرجعي",
          "بناء خطة مهمة المراجعة انطلاقاً من التقييم الشامل للمخاطر",
          "تحديد حجم العينة والإجراءات اللازمة بناءً على درجة المخاطر",
          "صياغة تقارير المراجعة القائمة على المخاطر بأسلوب مهني",
          "متابعة تنفيذ التوصيات وقياس أثر برامج المراجعة",
        ]},
      { "num":7,"day":2,
        "title":"دور المراجعة الداخلية في مكافحة الاحتيال والفساد المالي",
        "axes":[
          "طبيعة الاحتيال المالي وأساليب الكشف المبكر في المنشآت",
          "إطار الوقاية من الاحتيال ودور المراجع الداخلي المحوري",
          "إجراءات التحقيق الداخلي والأدلة الجنائية عند الاشتباه",
          "الامتثال لأنظمة مكافحة الاحتيال وغسيل الأموال والرشوة",
          "حالات دراسية من بيئة الأعمال السعودية والعربية",
        ]},
      { "num":8,"day":2,
        "title":"تقييم فاعلية الرقابة الداخلية وفق إطار COSO",
        "axes":[
          "مكونات إطار COSO للرقابة الداخلية ومبادئه الخمسة",
          "منهجيات تقييم فاعلية بيئة الرقابة الداخلية بشكل متكامل",
          "تحديد نقاط الضعف الجوهرية والإخفاقات الرقابية الحرجة",
          "بناء خطط المعالجة الفعّالة والمتابعة الدقيقة لتنفيذها",
          "قياس أثر برامج الرقابة الداخلية على الأداء المؤسسي",
        ]},
      { "num":9,"day":2,
        "title":"مستقبل مهنة المراجعة الداخلية في ظل رؤية 2030",
        "axes":[
          "الاتجاهات العالمية الراهنة لمهنة المراجعة الداخلية وتطورها",
          "أثر رؤية 2030 والإصلاحات الاقتصادية الكبرى على مسيرة المهنة",
          "متطلبات سوق العمل السعودي للمراجع الداخلي الكفء المستقبلي",
          "مسارات تطوير الكفاءات المهنية والشهادات الدولية المعتمدة",
          "رؤية استشرافية لمستقبل المراجعة الداخلية في المملكة",
        ]},
      { "num":10,"day":2,
        "title":"التكامل بين المراجعة الداخلية والخارجية ودور لجنة المراجعة",
        "axes":[
          "الفوارق والتكامل المنهجي بين المراجعة الداخلية والخارجية",
          "آليات التنسيق والتعاون الفعّال بين الفريقين الميدانيين",
          "دور لجنة المراجعة في الإشراف والحوكمة وتفعيل المساءلة",
          "معالجة الازدواجية في الجهود وتعظيم القيمة المضافة المشتركة",
          "نماذج تطبيقية ناجحة في تعزيز التنسيق والعمل المشترك",
        ]},
    ],
  },

  # ── ENGLISH ────────────────────────────────────────────────────────
  "en": {
    "lang":"en", "font":FONT_EN, "rtl":False,
    "am":PP_ALIGN.LEFT, "ac":PP_ALIGN.CENTER, "al":PP_ALIGN.LEFT,
    "event_title":  "Working Lunch Event",
    "event_sub":    "Specialised Internal Audit Dialogue Sessions",
    "proposal_tag": "PROPOSAL",
    "org_name":     "Al-Awfi & Al-Harbi",
    "org_sub":      "Certified Public Accountants & Legal Auditors  •  CLA Global Member",
    "web":          "www.cla.sa",
    "kpi":          ["2 Days", "10", "70–100"],
    "kpi_lbl":      ["Event Duration", "Dialogue Sessions", "Attendees / Day"],
    "toc_title":    "Table of Contents",
    "toc_items": [
      ("01", "Event Introduction"),
      ("02", "Meeting Benefits & Target Audience"),
      ("03", "Event Structure"),
      ("04", "Session 1  –  Internal Audit & Governance"),
      ("05", "Session 2  –  IPPF 2024 Standards"),
      ("06", "Session 3  –  Enterprise Risk Management"),
      ("07", "Session 4  –  Digital Transformation"),
      ("08", "Session 5  –  Ethics & Independence"),
      ("09", "Session 6  –  Risk-Based Auditing"),
      ("10", "Session 7  –  Fraud & Corruption"),
      ("11", "Session 8  –  Internal Control Evaluation"),
      ("12", "Session 9  –  Future of the Profession"),
      ("13", "Session 10 –  Internal & External Audit Integration"),
      ("14", "Estimated Costs"),
      ("15", "Projected Revenue"),
    ],
    "intro_title":  "About the Event",
    "intro_body": [
      "The Working Lunch Event in Internal Auditing is a premium professional gathering "
      "organised and sponsored by Al-Awfi & Al-Harbi – Certified Public Accountants & "
      "Legal Auditors (CLA Global Member) – combining academic insight with practical "
      "application in an interactive setting.",
      "The event spans two consecutive days and features ten specialised dialogue sessions, "
      "each addressing a pivotal theme in the internal audit profession through an "
      "engaging and enriching format.",
      "Every session is led by a specialist moderator from Al-Awfi & Al-Harbi alongside a "
      "corporate sponsor and two independent speakers of recognised expertise and "
      "deep experience in their respective fields.",
    ],
    "benefits_title": "Meeting Benefits",
    "benefits_cats": [
      {
        "cat":   "Benefits for Attendees & Participants",
        "items": [
          "Gain in-depth knowledge of the latest internal audit practices and standards",
          "Build a professional network with the elite of practitioners and auditors",
          "Discuss real-world challenges and draw solutions from domain experts",
          "Earn Continuing Professional Education (CPE) hours",
          "Exposure to leading local and global best practices in the profession",
        ],
      },
      {
        "cat":   "Benefits for Sponsors & Supporting Organisations",
        "items": [
          "Enhance brand visibility among decision-makers and senior professionals",
          "Target a high-calibre audience in a premium interactive environment",
          "Showcase products and services to a qualified, decision-making audience",
          "Build strategic partnerships with Al-Awfi & Al-Harbi",
          "Establish thought leadership in the internal audit landscape",
        ],
      },
      {
        "cat":   "Target Audience",
        "items": [
          "Audit committee chairs, members, and board directors",
          "Internal audit managers and professionals (public & private sectors)",
          "External auditors and financial consultants",
          "Chief Compliance Officers and enterprise risk managers",
          "Academics and researchers in accounting and auditing",
        ],
      },
    ],
    "structure_title": "Event Structure",
    "costs_title":     "Estimated Costs",
    "revenue_title":   "Projected Revenue",
    "closing_l1": "We look forward to a",
    "closing_l2": "successful partnership",
    "day_labels":   ["Day One", "Day Two"],
    "session_word": "Session",
    "agenda_title": "Session Agenda",
    "team_title":   "Session Team",
    "mod_lbl":  "Dialogue Moderator",
    "mod_val":  "[Moderator Name]\nAl-Awfi & Al-Harbi Office",
    "spon_lbl": "Session Sponsor",
    "spon_val": "[Sponsor Organisation]",
    "spon_fee": "Sponsorship Fee: SAR 10,000",
    "spkr_lbl": "Independent Speakers",
    "spkr_val": ["[Speaker 1]", "[Speaker 2]"],
    "spkr_fee": "SAR 4,000 / speaker",
    "day1_list": [
      "Session 1  –  Internal Audit & Governance",
      "Session 2  –  IPPF 2024 Standards",
      "Session 3  –  Enterprise Risk Management",
      "Session 4  –  Digital Transformation & Audit",
      "Session 5  –  Ethics & Independence",
    ],
    "day2_list": [
      "Session 6   –  Risk-Based Internal Auditing",
      "Session 7   –  Fraud & Financial Corruption",
      "Session 8   –  Internal Control (COSO)",
      "Session 9   –  Future of the Profession",
      "Session 10  –  Internal & External Audit Integration",
    ],
    "cost_hdr": ["Item", "Amount (SAR)", "Type"],
    "costs_rows": [
      ("Event Permit License",                        "SAR 5,000",  "Fixed"),
      ("Event Office Fees",                           "SAR 8,000",  "Fixed"),
      ("Hospitality – Minimum  (70 × 280 × 2 days)", "SAR 39,200", "Variable"),
      ("Hospitality – Maximum  (100 × 350 × 2 days)","SAR 70,000", "Variable"),
    ],
    "total_costs_lbl": "Total Estimated Costs",
    "total_costs":     "SAR 52,200  –  83,000",
    "cost_note":  "* Variable costs based on expected attendance of 70–100 per day × 2 days",
    "rev_rows": [
      ("Session Sponsors",      "10 sessions × SAR 10,000",           "SAR 100,000"),
      ("Independent Speakers",  "10 sessions × 2 speakers × SAR 4,000","SAR 80,000"),
    ],
    "total_rev_lbl": "Total Projected Revenue",
    "total_rev":     "SAR 180,000",
    "surplus_lbl":   "Expected Net Surplus",
    "surplus_val":   "SAR 97,000  –  127,800",
    "rev_note": "* Figures are estimates subject to negotiation with sponsors",
    "num_labels": ["First","Second","Third","Fourth","Fifth"],
    "sessions": [
      { "num":1,"day":1,
        "title":"The Role of Internal Audit in Enhancing Corporate Governance",
        "axes":[
          "Concept of corporate governance and Saudi regulatory requirements",
          "The internal auditor's role in supporting governance frameworks",
          "The complementary relationship between internal audit and the board",
          "Global best practices and success stories in audit governance",
          "Applied models from leading Saudi sector organisations",
        ]},
      { "num":2,"day":1,
        "title":"International Internal Auditing Standards (IPPF 2024) in the Saudi Context",
        "axes":[
          "Overview of the updated IPPF 2024 framework and key changes",
          "Compliance requirements and local implementation approaches",
          "Comparison between the old and new frameworks",
          "Practical challenges in applying standards within Saudi organisations",
          "A proposed roadmap for transitioning to the new framework",
        ]},
      { "num":3,"day":1,
        "title":"Enterprise Risk Management and Its Impact on the Internal Audit Plan",
        "axes":[
          "Enterprise Risk Management (ERM) framework and core pillars",
          "Identifying and assessing material risks in the Saudi business environment",
          "Building the annual audit plan on a comprehensive risk-assessment basis",
          "Aligning the risk map effectively with internal audit priorities",
          "Measuring the effectiveness of risk responses and updating the audit plan",
        ]},
      { "num":4,"day":1,
        "title":"Digital Transformation and Its Effect on the Internal Audit Profession",
        "axes":[
          "Artificial intelligence and big data analytics in auditing operations",
          "Modern digital audit tools and their practical field applications",
          "Information security auditing and infrastructure protection",
          "Competency development requirements for auditors in the digital age",
          "Opportunities and challenges of digital transformation for the profession",
        ]},
      { "num":5,"day":1,
        "title":"Independence, Objectivity, and Ethics in Internal Auditing",
        "axes":[
          "IIA Code of Ethics and professional conduct standards",
          "Organisational independence and individual objectivity – assurance mechanisms",
          "Managing and resolving conflicts of interest professionally",
          "The role of organisational culture in promoting integrity",
          "Strengthening the auditor's standing and building stakeholder trust",
        ]},
      { "num":6,"day":2,
        "title":"Risk-Based Internal Auditing – Methodology and Application",
        "axes":[
          "Risk-based internal auditing methodology and its theoretical reference framework",
          "Building the engagement plan based on a comprehensive risk assessment",
          "Determining sample size and procedures according to risk level",
          "Writing risk-based audit reports in a professional manner",
          "Following up on recommendations and measuring the impact of audit programmes",
        ]},
      { "num":7,"day":2,
        "title":"The Role of Internal Audit in Combating Fraud and Financial Corruption",
        "axes":[
          "Nature of financial fraud and early detection techniques in organisations",
          "Fraud prevention framework and the internal auditor's pivotal role",
          "Internal investigation procedures and forensic evidence when fraud is suspected",
          "Compliance with anti-fraud, anti-money-laundering and bribery regulations",
          "Case studies from the Saudi and Arab business environment",
        ]},
      { "num":8,"day":2,
        "title":"Evaluating Internal Control Effectiveness Based on the COSO Framework",
        "axes":[
          "COSO internal control framework: five components and principles",
          "Methodologies for assessing the effectiveness of control environments",
          "Identifying material weaknesses and critical control deficiencies",
          "Developing effective remediation plans and rigorously tracking implementation",
          "Measuring the impact of internal control programmes on institutional performance",
        ]},
      { "num":9,"day":2,
        "title":"The Future of the Internal Audit Profession Under Vision 2030",
        "axes":[
          "Current global trends shaping the internal audit profession",
          "Impact of Vision 2030 and major economic reforms on the profession",
          "Saudi labour market requirements for the competent future internal auditor",
          "Professional development pathways and internationally recognised certifications",
          "A forward-looking vision for internal auditing in the Kingdom",
        ]},
      { "num":10,"day":2,
        "title":"Integration Between Internal & External Audit and the Audit Committee",
        "axes":[
          "Differences and methodological synergies between internal and external audit",
          "Mechanisms for effective coordination between both field teams",
          "The audit committee's role in oversight, governance, and accountability",
          "Eliminating duplication and maximising shared added value",
          "Successful models of collaboration and joint working",
        ]},
    ],
  },
}

# ══════════════════════════════════════════════════════════════════════
# SLIDE BUILDER
# ══════════════════════════════════════════════════════════════════════
def build(d):
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    BL = prs.slide_layouts[6]
    F  = d["font"]; AR = d["rtl"]
    AM = d["am"]; AC = d["ac"]; AL = d["al"]

    def T(sl, text, l, t, w, h, sz=16, bold=False, col=DARK, align=None,
          italic=False, ltr=False):
        a = align if align is not None else AM
        r = not AR if ltr else AR
        return tb(sl, text, l, t, w, h, sz, bold, col, a, r, italic, F)

    def ML(sl, lines, l, t, w, h, sz=14, bold=False, col=DARK, align=None,
           sp_a=6, hdr=None, hdr_sz=None, hdr_col=None, ltr=False):
        a = align if align is not None else AM
        r = not AR if ltr else AR
        return mltb(sl, lines, l, t, w, h, sz, bold, col, a, sp_a, r, F,
                    hdr, hdr_sz, True, hdr_col)

    def foot(sl):
        R(sl, 0, 7.2, 13.33, 0.3, NAVY)
        tb(sl, d["org_name"] + "   |   " + d["web"],
           0.4, 7.22, 12.53, 0.26, sz=11, col=WHITE, align=AC, rtl=False, font=F)

    def std_nav_bar(sl, title):
        """Clean navy top bar – no diagonal shapes."""
        R(sl, 0, 0, 13.33, 7.5, WHITE)
        R(sl, 0, 0, 13.33, 1.3, NAVY)
        # Small square accent left of title
        R(sl, 0.4, 0.35, 0.06, 0.6, WHITE)
        T(sl, title, 0.6, 0.18, 11, 0.95, sz=30, bold=True, col=WHITE, align=AL)
        foot(sl)

    # ── COVER ─────────────────────────────────────────────────────────
    def add_cover():
        sl = prs.slides.add_slide(BL)
        R(sl, 0, 0, 13.33, 7.5, OFF_W)
        # Left navy panel – clean rectangle, no diagonal
        R(sl, 0, 0, 6.8, 7.5, NAVY2)
        # Thin accent stripe inside left panel
        R(sl, 0, 0, 0.06, 7.5, NAVY_M)
        R(sl, 0, 0, 6.8, 0.06, NAVY_M)
        R(sl, 0, 7.44, 6.8, 0.06, NAVY_M)

        # Proposal tag
        R(sl, 0.45, 0.65, 2.0, 0.38, NAVY_M)
        tb(sl, d["proposal_tag"], 0.45, 0.65, 2.0, 0.38,
           sz=12, bold=True, col=WHITE, align=AC, rtl=False, font=F)

        # Main title
        tb(sl, d["event_title"], 0.45, 1.25, 6.15, 1.4,
           sz=44, bold=True, col=WHITE, align=AL, rtl=AR, font=F)
        # Subtitle
        tb(sl, d["event_sub"], 0.45, 2.72, 6.15, 0.65,
           sz=17, col=GRAY_L, align=AL, rtl=AR, font=F)

        R(sl, 0.45, 3.6, 5.2, 0.05, WHITE)

        tb(sl, d["org_name"], 0.45, 3.82, 6.15, 0.68,
           sz=22, bold=True, col=WHITE, align=AL, rtl=AR, font=F)
        tb(sl, d["org_sub"], 0.45, 4.48, 6.15, 0.45,
           sz=11, col=GRAY_M, align=AL, rtl=False, font=F)
        tb(sl, d["web"], 0.45, 6.72, 2.8, 0.4,
           sz=13, col=GRAY_L, align=AL, rtl=False, font=F)

        # Right side KPI tiles
        for i, (val, lbl) in enumerate(zip(d["kpi"], d["kpi_lbl"])):
            ky = 1.0 + i * 1.85
            R(sl, 7.35, ky, 5.5, 1.55, WHITE)
            R(sl, 7.35, ky, 0.07, 1.55, NAVY)
            tb(sl, val, 7.55, ky+0.1, 5.2, 0.78,
               sz=38, bold=True, col=NAVY2, align=AL, rtl=False, font=F)
            tb(sl, lbl, 7.55, ky+0.9, 5.2, 0.45,
               sz=14, col=GRAY_M, align=AL, rtl=False, font=F)

    add_cover()

    # ── TABLE OF CONTENTS ─────────────────────────────────────────────
    def add_toc():
        sl = prs.slides.add_slide(BL)
        R(sl, 0, 0, 13.33, 7.5, WHITE)
        # Narrow left navy column
        R(sl, 0, 0, 0.55, 7.5, NAVY2)
        R(sl, 0.55, 0, 0.04, 7.5, GRAY_L)
        # Top header strip
        R(sl, 0, 0, 13.33, 0.07, NAVY)
        foot(sl)

        T(sl, d["toc_title"], 0.75, 0.12, 5, 0.55,
          sz=26, bold=True, col=NAVY2, align=AL)
        R(sl, 0.75, 0.72, 11.8, 0.04, NAVY)

        items = d["toc_items"]
        half  = len(items) // 2 + (len(items) % 2)
        col_x = [0.75, 7.1]

        for col_i, start in enumerate([0, half]):
            chunk = items[start: start + half]
            for row_i, (num, label) in enumerate(chunk):
                ry = 0.88 + row_i * 0.42
                # Number badge
                R(sl, col_x[col_i], ry, 0.42, 0.34, NAVY)
                tb(sl, num, col_x[col_i], ry, 0.42, 0.34,
                   sz=11, bold=True, col=WHITE, align=AC, rtl=False, font=F)
                # Label
                tb(sl, label, col_x[col_i]+0.5, ry+0.01, 5.8, 0.34,
                   sz=12, col=DARK, align=AL, rtl=False, font=F)
                # Thin row divider
                R(sl, col_x[col_i], ry+0.36, 6.2, 0.02, GRAY_L)

    add_toc()

    # ── INTRODUCTION ──────────────────────────────────────────────────
    def add_intro():
        sl = prs.slides.add_slide(BL)
        std_nav_bar(sl, d["intro_title"])

        R(sl, 0.4, 1.45, 12.53, 5.6, OFF_W)
        R(sl, 0.4, 1.45, 0.07, 5.6, NAVY)

        for i, para in enumerate(d["intro_body"]):
            py = 1.65 + i * 1.72
            R(sl, 0.6, py, 12.13, 1.5, WHITE)
            R(sl, 0.6, py, 12.13, 0.06, NAVY_M)
            # Paragraph number
            R(sl, 0.72, py+0.22, 0.46, 0.46, NAVY)
            tb(sl, str(i+1), 0.72, py+0.22, 0.46, 0.46,
               sz=14, bold=True, col=WHITE, align=AC, rtl=False, font=F)
            tb(sl, para, 1.3, py+0.14, 11.2, 1.22,
               sz=13, col=DARK, align=AM, rtl=AR, font=F)

    add_intro()

    # ── BENEFITS ──────────────────────────────────────────────────────
    def add_benefits():
        sl = prs.slides.add_slide(BL)
        std_nav_bar(sl, d["benefits_title"])

        cats = d["benefits_cats"]
        cw = 3.96; gap = 0.22
        for i, cat in enumerate(cats):
            cx = 0.4 + i * (cw + gap)
            # Card
            R(sl, cx, 1.45, cw, 5.6, OFF_W)
            R(sl, cx, 1.45, cw, 0.55, NAVY)
            tb(sl, cat["cat"], cx, 1.45, cw, 0.55,
               sz=12, bold=True, col=WHITE, align=AC, rtl=AR, font=F)
            for j, item in enumerate(cat["items"]):
                iy = 2.1 + j * 1.0
                R(sl, cx+0.12, iy, cw-0.24, 0.85, WHITE)
                R(sl, cx+0.12, iy, 0.05, 0.85, NAVY_M)
                # bullet number
                R(sl, cx+0.22, iy+0.2, 0.36, 0.36, NAVY)
                tb(sl, str(j+1), cx+0.22, iy+0.2, 0.36, 0.36,
                   sz=11, bold=True, col=WHITE, align=AC, rtl=False, font=F)
                tb(sl, item, cx+0.65, iy+0.06, cw-0.82, 0.72,
                   sz=11, col=DARK, align=AM, rtl=AR, font=F)

    add_benefits()

    # ── STRUCTURE ─────────────────────────────────────────────────────
    def add_structure():
        sl = prs.slides.add_slide(BL)
        std_nav_bar(sl, d["structure_title"])

        for d_i, (day_lbl, sess_list, offset, cx) in enumerate([
            (d["day_labels"][0], d["day1_list"], 0, 0.4),
            (d["day_labels"][1], d["day2_list"], 5, 6.87),
        ]):
            cw = 6.26
            R(sl, cx, 1.45, cw, 5.65, OFF_W)
            R(sl, cx, 1.45, cw, 0.55, NAVY)
            tb(sl, day_lbl, cx, 1.45, cw, 0.55,
               sz=16, bold=True, col=WHITE, align=AC, rtl=False, font=F)
            for i, name in enumerate(sess_list):
                ry = 2.1 + i * 1.0
                bg = WHITE if i % 2 == 0 else GRAY_L
                R(sl, cx+0.05, ry, cw-0.1, 0.9, bg)
                R(sl, cx+0.05, ry, 0.05, 0.9, NAVY_M)
                # num badge
                R(sl, cx+cw-0.6, ry+0.22, 0.44, 0.44, NAVY)
                tb(sl, str(i+1+offset), cx+cw-0.6, ry+0.22, 0.44, 0.44,
                   sz=12, bold=True, col=WHITE, align=AC, rtl=False, font=F)
                tb(sl, name, cx+0.18, ry+0.2, cw-0.9, 0.52,
                   sz=11, col=DARK, align=AL, rtl=False, font=F)

    add_structure()

    # ── SESSION SLIDES (5 agenda items each) ──────────────────────────
    def add_session(s):
        sl = prs.slides.add_slide(BL)
        R(sl, 0, 0, 13.33, 7.5, WHITE)
        R(sl, 0, 0, 13.33, 0.06, NAVY)
        # Left sidebar
        R(sl, 0, 0, 0.55, 7.5, NAVY2)
        # Session number on sidebar
        R(sl, 0, 0.55, 0.55, 0.65, NAVY_M)
        tb(sl, str(s["num"]), 0, 0.55, 0.55, 0.65,
           sz=22, bold=True, col=WHITE, align=AC, rtl=False, font=F)
        # Day on sidebar
        day_s = "D1" if s["day"] == 1 else "D2"
        tb(sl, day_s, 0, 1.28, 0.55, 0.35,
           sz=11, bold=True, col=GRAY_L, align=AC, rtl=False, font=F)
        foot(sl)

        # Title bar
        R(sl, 0.7, 0.1, 12.45, 1.22, OFF_W)
        R(sl, 0.7, 0.1, 12.45, 0.04, NAVY_M)
        tag_txt = d["session_word"] + "  " + str(s["num"]) + "   |   " + d["day_labels"][s["day"]-1]
        tb(sl, tag_txt, 0.82, 0.14, 10, 0.35,
           sz=11, bold=True, col=NAVY_M, align=AL, rtl=False, font=F)
        tb(sl, s["title"], 0.82, 0.46, 12.1, 0.82,
           sz=20, bold=True, col=NAVY2, align=AM, rtl=AR, font=F)
        R(sl, 0.7, 1.36, 12.45, 0.04, NAVY)

        # Agenda card (main 68% width)
        R(sl, 0.7, 1.48, 8.6, 5.82, OFF_W)
        R(sl, 0.7, 1.48, 8.6, 0.44, NAVY)
        tb(sl, d["agenda_title"], 0.84, 1.48, 8.3, 0.44,
           sz=13, bold=True, col=WHITE, align=AL, rtl=False, font=F)

        # 5 agenda items – compact row height
        row_h = 1.04
        for i, (axis, lbl) in enumerate(zip(s["axes"], d["num_labels"])):
            ay = 2.0 + i * row_h
            if i % 2 == 1:
                R(sl, 0.7, ay-0.04, 8.6, row_h, GRAY_L)
            R(sl, 0.7, ay, 0.05, row_h-0.04, NAVY_M)
            R(sl, 0.82, ay+0.27, 1.02, 0.36, NAVY)
            tb(sl, lbl, 0.82, ay+0.27, 1.02, 0.36,
               sz=11, bold=True, col=WHITE, align=AC, rtl=False, font=F)
            tb(sl, axis, 1.92, ay+0.12, 7.2, 0.78,
               sz=12, col=DARK, align=AM, rtl=AR, font=F)

        # Speakers panel (right 28%)
        R(sl, 9.45, 1.48, 3.73, 5.82, NAVY2)
        R(sl, 9.45, 1.48, 3.73, 0.06, NAVY_M)
        tb(sl, d["team_title"], 9.45, 1.56, 3.73, 0.44,
           sz=13, bold=True, col=WHITE, align=AC, rtl=False, font=F)

        sections = [
            (d["mod_lbl"],  d["mod_val"],               None),
            (d["spon_lbl"], d["spon_val"],               d["spon_fee"]),
            (d["spkr_lbl"], "\n".join(d["spkr_val"]),   d["spkr_fee"]),
        ]
        sy = 2.1
        for lbl, val, fee in sections:
            R(sl, 9.55, sy, 3.53, 0.32, NAVY_M)
            tb(sl, lbl, 9.55, sy, 3.53, 0.32,
               sz=11, bold=True, col=WHITE, align=AC, rtl=False, font=F)
            sy += 0.34
            multi = '\n' in val
            vh = 0.72 if multi else 0.48
            tb(sl, val, 9.55, sy, 3.53, vh,
               sz=11, col=GRAY_L, align=AC, rtl=False, font=F)
            sy += vh + 0.04
            if fee:
                tb(sl, fee, 9.55, sy, 3.53, 0.3,
                   sz=10, col=NAVY_M, align=AC, rtl=False, font=F)
                sy += 0.32
            sy += 0.14

    for s in d["sessions"]:
        add_session(s)

    # ── COSTS ─────────────────────────────────────────────────────────
    def add_costs():
        sl = prs.slides.add_slide(BL)
        std_nav_bar(sl, d["costs_title"])

        R(sl, 0.4, 1.45, 12.53, 0.48, NAVY2)
        for i, hdr in enumerate(d["cost_hdr"]):
            xs = [0.55, 10.3, 11.85]
            tb(sl, hdr, xs[i], 1.45, 5.0, 0.48,
               sz=13, bold=True, col=WHITE, align=AL, rtl=False, font=F)

        for i, (item, amount, kind) in enumerate(d["costs_rows"]):
            ry = 2.03 + i * 1.12
            bg = OFF_W if i % 2 == 0 else WHITE
            R(sl, 0.4, ry, 12.53, 1.0, bg)
            R(sl, 0.4, ry, 0.06, 1.0, NAVY_M)
            tb(sl, item,   0.6,  ry+0.25, 9.5, 0.55, sz=13, col=DARK,
               align=AL, rtl=False, font=F)
            tb(sl, amount, 10.3, ry+0.25, 1.5, 0.55, sz=14, bold=True,
               col=NAVY, align=AL, rtl=False, font=F)
            # type pill
            R(sl, 11.83, ry+0.28, 0.85, 0.38, NAVY)
            tb(sl, kind, 11.83, ry+0.28, 0.85, 0.38,
               sz=11, bold=True, col=WHITE, align=AC, rtl=False, font=F)

        R(sl, 0.4, 6.55, 12.53, 0.55, NAVY)
        tb(sl, d["total_costs_lbl"], 0.6,  6.55, 9.0, 0.55,
           sz=15, bold=True, col=WHITE, align=AL, rtl=False, font=F)
        tb(sl, d["total_costs"],     9.8,  6.55, 3.0, 0.55,
           sz=15, bold=True, col=GRAY_L, align=AL, rtl=False, font=F)

        tb(sl, d["cost_note"], 0.45, 7.18, 12, 0.28,
           sz=10, col=GRAY_M, align=AL, rtl=False, font=F)

    add_costs()

    # ── REVENUE ───────────────────────────────────────────────────────
    def add_revenue():
        sl = prs.slides.add_slide(BL)
        std_nav_bar(sl, d["revenue_title"])

        for i, (src, calc, total) in enumerate(d["rev_rows"]):
            ry = 1.5 + i * 1.9
            R(sl, 0.4, ry, 12.53, 1.72, OFF_W if i%2==0 else WHITE)
            R(sl, 0.4, ry, 0.07, 1.72, NAVY)
            tb(sl, src,  0.65, ry+0.15, 9.0, 0.6,  sz=19, bold=True,
               col=NAVY2, align=AL, rtl=False, font=F)
            tb(sl, calc, 0.65, ry+0.8,  9.0, 0.45, sz=13, col=GRAY_M,
               align=AL, rtl=False, font=F)
            R(sl, 10.1, ry+0.42, 2.7, 0.75, NAVY)
            tb(sl, total, 10.1, ry+0.42, 2.7, 0.75,
               sz=18, bold=True, col=WHITE, align=AC, rtl=False, font=F)

        R(sl, 0.4, 5.45, 12.53, 0.68, NAVY2)
        tb(sl, d["total_rev_lbl"], 0.6,  5.45, 8.5, 0.68,
           sz=16, bold=True, col=WHITE, align=AL, rtl=False, font=F)
        tb(sl, d["total_rev"],     9.8,  5.45, 3.0, 0.68,
           sz=18, bold=True, col=GRAY_L, align=AL, rtl=False, font=F)

        R(sl, 0.4, 6.28, 12.53, 0.75, OFF_W)
        R(sl, 0.4, 6.28, 0.07,  0.75, NAVY_M)
        tb(sl, d["surplus_lbl"],  0.65, 6.28, 6.0, 0.75,
           sz=14, bold=True, col=NAVY2, align=AL, rtl=False, font=F)
        tb(sl, d["surplus_val"], 7.1, 6.28, 5.5, 0.75,
           sz=14, bold=True, col=NAVY, align=AL, rtl=False, font=F)

        tb(sl, d["rev_note"], 0.45, 7.18, 12, 0.28,
           sz=10, col=GRAY_M, align=AL, rtl=False, font=F)

    add_revenue()

    # ── CLOSING ───────────────────────────────────────────────────────
    def add_closing():
        sl = prs.slides.add_slide(BL)
        R(sl, 0, 0, 13.33, 7.5, NAVY2)
        R(sl, 0, 0, 13.33, 0.06, NAVY_M)
        R(sl, 0, 7.44, 13.33, 0.06, NAVY_M)
        # Right white panel – clean rectangle
        R(sl, 7.2, 0, 6.13, 7.5, WHITE)
        # Thin border between panels
        R(sl, 7.2, 0, 0.05, 7.5, NAVY_M)

        tb(sl, d["closing_l1"], 0.55, 2.3, 6.4, 0.85,
           sz=32, bold=True, col=WHITE, align=AL, rtl=AR, font=F)
        tb(sl, d["closing_l2"], 0.55, 3.12, 6.4, 0.85,
           sz=32, bold=True, col=GRAY_L, align=AL, rtl=AR, font=F)
        R(sl, 0.55, 4.1, 5.0, 0.05, NAVY_M)
        tb(sl, d["org_name"], 0.55, 4.3, 6.4, 0.6,
           sz=20, bold=True, col=WHITE, align=AL, rtl=AR, font=F)
        tb(sl, d["org_sub"], 0.55, 4.88, 6.4, 0.45,
           sz=11, col=GRAY_M, align=AL, rtl=False, font=F)
        tb(sl, d["web"], 0.55, 5.58, 3.5, 0.42,
           sz=14, col=GRAY_L, align=AL, rtl=False, font=F)

        # Right panel – bold stats
        tb(sl, "10", 8.3, 1.3, 4.0, 1.5,
           sz=88, bold=True, col=NAVY2, align=AC, rtl=False, font=F)
        tb(sl, d["kpi_lbl"][1], 8.3, 2.72, 4.0, 0.5,
           sz=15, col=GRAY_M, align=AC, rtl=False, font=F)
        R(sl, 9.0, 3.42, 2.5, 0.05, GRAY_L)
        tb(sl, "2",  8.3, 3.65, 4.0, 1.1,
           sz=68, bold=True, col=NAVY2, align=AC, rtl=False, font=F)
        tb(sl, d["kpi_lbl"][0], 8.3, 4.65, 4.0, 0.5,
           sz=15, col=GRAY_M, align=AC, rtl=False, font=F)

    add_closing()

    return prs

# ── Build both ────────────────────────────────────────────────────────
for lang, suffix in [
    ("ar", "عرض_مقترح_فعالية_غداء_العمل_v3"),
    ("en", "Working_Lunch_Event_Proposal_v3"),
]:
    p = build(DATA[lang])
    out = f"/home/user/Abdulsalam-/{suffix}.pptx"
    p.save(out)
    print(f"✅  {out}  ({len(p.slides)} slides)")
