#!/usr/bin/env python3
"""Generate LOT Atlanta x SAP deck PPTX -- 12 slides."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

ORANGE = RGBColor(0xD9, 0x77, 0x06)  # NBMBAA amber — readable on light bg
SAP    = RGBColor(0x00, 0x55, 0xCC)  # deep SAP blue
GREEN  = RGBColor(0x16, 0xA3, 0x4A)
BG     = RGBColor(0xED, 0xF1, 0xFB)  # light blue-white
SURF   = RGBColor(0xFF, 0xFF, 0xFF)  # white cards
SURF2  = RGBColor(0xE5, 0xEC, 0xFA)  # light blue tint
BORDER = RGBColor(0xC4, 0xD0, 0xEC)
TEXT   = RGBColor(0x1C, 0x22, 0x40)  # dark navy
MUTED  = RGBColor(0x4E, 0x5B, 0x78)
DIM    = RGBColor(0x6B, 0x74, 0x90)
BRIGHT = RGBColor(0x0A, 0x0E, 0x20)  # darkest navy — highest contrast on light bg
BLACK  = RGBColor(0x00, 0x00, 0x00)

W = Inches(13.33)
H = Inches(7.5)
FD = "Georgia"
FU = "Calibri"
FB = "Calibri"
FM = "Courier New"


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])

def bg(sl, c=BG):
    f = sl.background.fill; f.solid(); f.fore_color.rgb = c

def rect(sl, l, t, w, h, fill=None, line=None, lw=Pt(0.75)):
    sh = sl.shapes.add_shape(1, l, t, w, h)
    if fill: sh.fill.solid(); sh.fill.fore_color.rgb = fill
    else: sh.fill.background()
    if line: sh.line.color.rgb = line; sh.line.width = lw
    else: sh.line.fill.background()
    return sh

def tb(sl, l, t, w, h):
    return sl.shapes.add_textbox(l, t, w, h).text_frame

def para(tf, text, font=FB, size=14, bold=False, color=TEXT,
         align=PP_ALIGN.LEFT, italic=False):
    p = tf.add_paragraph(); p.alignment = align
    r = p.add_run(); r.text = text
    r.font.name = font; r.font.size = Pt(size)
    r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color
    return p

def kicker(sl, text, top=Inches(0.55)):
    tf = tb(sl, Inches(0.7), top, Inches(12), Pt(18))
    para(tf, text.upper(), font=FU, size=9, bold=True, color=SAP)

def stripe(sl):
    """Split co-brand stripe: SAP blue left, NBMBAA orange right."""
    half = W / 2
    r1 = sl.shapes.add_shape(1, 0, 0, half, Pt(4))
    r1.fill.solid(); r1.fill.fore_color.rgb = SAP; r1.line.fill.background()
    r2 = sl.shapes.add_shape(1, half, 0, half, Pt(4))
    r2.fill.solid(); r2.fill.fore_color.rgb = ORANGE; r2.line.fill.background()

def title(sl, text, top=Inches(0.85), size=44):
    tf = tb(sl, Inches(0.7), top, Inches(10), Inches(2.0))
    tf.word_wrap = True; para(tf, text, font=FD, size=size, color=BRIGHT)

def orange_rule(sl, top=Inches(2.0)):
    r = sl.shapes.add_shape(1, Inches(0.7), top, Inches(0.72), Pt(3))
    r.fill.solid(); r.fill.fore_color.rgb = ORANGE; r.line.fill.background()

def card(sl, l, t, w, h, accent=None):
    rect(sl, l, t, w, h, fill=SURF, line=BORDER)
    if accent:
        rect(sl, l, t, w, Pt(2.5), fill=accent)

def notes(sl, text):
    sl.notes_slide.notes_text_frame.text = text


# ---- Slide functions ----

def s1_title(prs):
    sl = blank(prs); bg(sl)
    stripe(sl)
    rect(sl, 0, H-Pt(4), W, Pt(4), fill=ORANGE)  # NBMBAA orange bottom
    # SAP co-brand box top-right
    cb_w = Inches(2.6); cb_h = Pt(28)
    cb_l = W - cb_w - Inches(0.5)
    r_sap = sl.shapes.add_shape(1, cb_l, Inches(0.55), Inches(1.1), cb_h)
    r_sap.fill.solid(); r_sap.fill.fore_color.rgb = SAP; r_sap.line.fill.background()
    tf_sap = tb(sl, cb_l, Inches(0.55), Inches(1.1), cb_h)
    para(tf_sap, "SAP", font=FU, size=13, bold=True, color=RGBColor(0xFF,0xFF,0xFF), align=PP_ALIGN.CENTER)
    r_nbm = sl.shapes.add_shape(1, cb_l+Inches(1.1), Inches(0.55), Inches(1.5), cb_h)
    r_nbm.fill.solid(); r_nbm.fill.fore_color.rgb = ORANGE; r_nbm.line.fill.background()
    tf_nbm = tb(sl, cb_l+Inches(1.1), Inches(0.55), Inches(1.5), cb_h)
    para(tf_nbm, "NBMBAA LOT", font=FU, size=8, bold=True, color=RGBColor(0xFF,0xFF,0xFF), align=PP_ALIGN.CENTER)
    tf = tb(sl, Inches(0.7), Inches(0.65), Inches(5), Pt(22))
    para(tf, "NBMBAA(R) LOT x SAP", font=FU, size=9, bold=True, color=MUTED)
    tf2 = tb(sl, Inches(0.7), Inches(1.4), Inches(9), Inches(3.2))
    tf2.word_wrap = True
    para(tf2, "Build Your", font=FD, size=80, color=BRIGHT)
    para(tf2, "AI-Powered", font=FD, size=80, color=ORANGE)
    para(tf2, "Career Portfolio", font=FD, size=80, color=BRIGHT)
    tf3 = tb(sl, Inches(0.7), Inches(6.3), Inches(9), Pt(22))
    para(tf3, "October 17, 2026  .  Atlanta  .  Powered by Google Gemini & GitHub Pages",
         font=FU, size=11, color=MUTED)
    notes(sl, (
        "[NEIL] Hey everyone — we're Neil and India from SAP. We have 30 minutes and we're "
        "going to use all of it. By the time we're done, every one of you will have a real "
        "AI-generated portfolio page — the kind you can link to on a college application or LinkedIn. "
        "No coding experience, no software to install, nothing to pay for. "
        "Two free tools. Let's go."
    ))

def s2_big_picture(prs):
    sl = blank(prs); bg(sl); stripe(sl)
    kicker(sl, "Full Event -- Today's Schedule")
    title(sl, "The Big Picture")
    orange_rule(sl, Inches(1.95))
    blocks = [
        ("30 MIN", "AI Exploration\nGamefied",    None),
        ("10 MIN", "Intro to SAP",                 None),
        ("30 MIN", "Hands-On:\nBuild with AI",     ORANGE),
        ("10 MIN", "Responsible AI",                None),
        ("10 MIN", "Closing &\nDiscussion",         None),
    ]
    cw = Inches(2.28); gap = Inches(0.16); top = Inches(2.1); h = Inches(3.8)
    for i, (tm, ttl, acc) in enumerate(blocks):
        l = Inches(0.7) + i*(cw+gap)
        card(sl, l, top, cw, h, accent=acc or DIM)
        tf_t = tb(sl, l+Pt(10), top+Pt(10), cw-Pt(20), Pt(18))
        para(tf_t, tm, font=FU, size=9, bold=True, color=ORANGE if acc else MUTED)
        tf_h = tb(sl, l+Pt(10), top+Pt(32), cw-Pt(20), Pt(50))
        tf_h.word_wrap = True
        para(tf_h, ttl, font=FU, size=13, bold=True, color=BRIGHT)
        if acc:
            yb = rect(sl, l+Pt(8), top+h-Pt(28), cw-Pt(16), Pt(20), fill=ORANGE)
            tf_y = tb(sl, l+Pt(8), top+h-Pt(28), cw-Pt(16), Pt(20))
            para(tf_y, "<- YOU ARE HERE", font=FU, size=8, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
    tf_tot = tb(sl, Inches(10.5), Inches(6.1), Inches(2.5), Pt(36))
    para(tf_tot, "90 MIN", font=FD, size=32, color=ORANGE, align=PP_ALIGN.RIGHT)
    notes(sl, (
        "[NEIL or INDIA] Quick map of the morning. You've already done the first two blocks. "
        "We're the orange one — the hands-on build. After us: a responsible AI segment, then closing. "
        "Our 30 minutes. Everything we're doing is built on two free tools: "
        "Google Gemini for the AI, and GitHub to host the website. "
        "You've probably heard of both. You're about to use them together to build something real."
    ))

def s3_session(prs):
    sl = blank(prs); bg(sl); stripe(sl)
    kicker(sl, "Session Overview")
    title(sl, "Today's Session")
    orange_rule(sl, Inches(1.95))
    rect(sl, Inches(0.7), Inches(2.1), Inches(12), Inches(0.75),
         fill=RGBColor(0xFF,0xF0,0xD8), line=RGBColor(0xD9,0x77,0x06))
    tf_b = tb(sl, Inches(0.85), Inches(2.18), Inches(11.7), Pt(44))
    tf_b.word_wrap = True
    para(tf_b, "Employers don't just want diplomas -- they want to see what you can DO. "
              "In 30 minutes you'll build a real AI-generated portfolio for college apps, internships, and LinkedIn.",
         font=FB, size=12, color=TEXT)
    beats = [
        ("5 MIN",  "Audience Warmup",               "Quick poll -- careers? AI usage?",               GREEN),
        ("5 MIN",  "AI & Your Career",              "How AI reshapes every field.",                    SAP),
        ("15 MIN", "Build: Gemini -> Portfolio -> Live", "Demo first. Then you run it. Gemini writes your full webpage.", ORANGE),
        ("5 MIN",  "Share & Wrap",                  "Volunteers share. Take-home challenge.",          MUTED),
    ]
    cw = Inches(2.85); gap = Inches(0.13); top = Inches(3.05); h = Inches(3.6)
    for i, (tm, head, desc, ac) in enumerate(beats):
        l = Inches(0.7) + i*(cw+gap)
        card(sl, l, top, cw, h, accent=ac)
        tf_t = tb(sl, l+Pt(10), top+Pt(10), cw-Pt(20), Pt(16))
        para(tf_t, tm, font=FU, size=9, bold=True, color=MUTED)
        tf_h = tb(sl, l+Pt(10), top+Pt(28), cw-Pt(20), Pt(34))
        tf_h.word_wrap = True
        para(tf_h, head, font=FU, size=12, bold=True, color=BRIGHT)
        tf_d = tb(sl, l+Pt(10), top+Pt(68), cw-Pt(20), Inches(1.5))
        tf_d.word_wrap = True
        para(tf_d, desc, font=FB, size=11, color=MUTED)
    notes(sl, (
        "[INDIA] Here's our 30 minutes. Five minutes warming up, five on AI and careers, "
        "fifteen minutes building — that's the main event — and five minutes to share. "
        "The key idea on this slide: employers don't care that you went to school. "
        "They want to see what you built. A portfolio page with your actual name, skills, "
        "and career thinking? That stands out from every other applicant who just submits a transcript."
    ))

def s4_ai_careers(prs):
    sl = blank(prs); bg(sl); stripe(sl)
    kicker(sl, "AI & Your Future")
    title(sl, "AI Is Already in the Room")
    orange_rule(sl, Inches(1.95))
    industries = [
        ("Healthcare",  "AI reads scans, flags risks, and surfaces diagnoses faster than any one doctor."),
        ("Business",    "AI drafts contracts, forecasts revenue, and analyzes deals in seconds."),
        ("Creative",    "AI generates images, music, and video. Humans direct and judge it."),
        ("Sports",      "AI tracks athlete data in real time -- predicting injuries before they happen."),
        ("Engineering", "AI designs structures, models climate, and simulates systems at scale."),
    ]
    cw = Inches(2.28); gap = Inches(0.16); top = Inches(2.1); h = Inches(2.9)
    for i, (name, desc) in enumerate(industries):
        l = Inches(0.7) + i*(cw+gap)
        card(sl, l, top, cw, h)
        tf_n = tb(sl, l+Pt(10), top+Pt(12), cw-Pt(20), Pt(22))
        para(tf_n, name, font=FU, size=13, bold=True, color=ORANGE)
        tf_d = tb(sl, l+Pt(10), top+Pt(38), cw-Pt(20), Inches(1.8))
        tf_d.word_wrap = True
        para(tf_d, desc, font=FB, size=11, color=MUTED)
    rect(sl, Inches(0.7), Inches(5.2), Inches(12), Inches(0.85),
         fill=RGBColor(0xFF,0xF0,0xD8), line=RGBColor(0xD9,0x77,0x06))
    tf_c = tb(sl, Inches(0.85), Inches(5.3), Inches(11.7), Pt(44))
    tf_c.word_wrap = True
    para(tf_c, "The real skill in every field: knowing how to ask AI the right question -- and knowing when to trust the answer.",
         font=FB, size=13, color=TEXT)
    notes(sl, (
        "[NEIL] Show of hands — who's thinking about healthcare? Business? Creative work? "
        "Sports? Engineering? [pause] Here's what's true for every one of those fields: "
        "AI is already in them. AI reads medical scans faster than any one doctor. "
        "AI writes contracts and runs financial models in seconds. AI generates music, images, video. "
        "AI tracks athlete performance in real time and predicts injuries before they happen. "
        "AI designs physical structures and simulates climate systems. This isn't a future thing — "
        "it's happening now. The students who learn to work WITH AI, not just use it, will have a real edge. "
        "That's what we're showing you today."
    ))

def s5_architecture(prs):
    sl = blank(prs); bg(sl); stripe(sl)
    kicker(sl, "How It Works -- System Overview")
    title(sl, "The Architecture")
    orange_rule(sl, Inches(1.95))
    nodes = [
        ("You",           "Fill in 4 fields",       "Grade . School\nActivity . Career",   MUTED),
        ("Google Gemini", "AI generates code",       "Complete HTML + CSS\nin seconds",     ORANGE),
        ("GitHub",        "Free file hosting",       "Web UI only --\nno code, no tokens",  SAP),
        ("Live Website",  "Public URL, free forever","yourname\n.github.io/portfolio",       GREEN),
    ]
    conn_labels = ["Your Prompt", "Copy & Paste", "Pages ON"]
    cw = Inches(2.4); gap_conn = Inches(0.55); top = Inches(2.2); h = Inches(2.8)
    total_w = 4*cw + 3*gap_conn
    start_l = (W - total_w) / 2
    for i, (ntitle, nhead, ndetail, ac) in enumerate(nodes):
        l = start_l + i*(cw+gap_conn)
        card(sl, l, top, cw, h, accent=ac)
        tf_t = tb(sl, l+Pt(10), top+Pt(14), cw-Pt(20), Pt(22))
        para(tf_t, ntitle, font=FU, size=14, bold=True, color=BRIGHT, align=PP_ALIGN.CENTER)
        tf_h = tb(sl, l+Pt(10), top+Pt(40), cw-Pt(20), Pt(20))
        para(tf_h, nhead, font=FB, size=11, color=MUTED, align=PP_ALIGN.CENTER)
        tf_d = tb(sl, l+Pt(10), top+Pt(60), cw-Pt(20), Pt(44))
        tf_d.word_wrap = True
        para(tf_d, ndetail, font=FM, size=10, color=DIM, align=PP_ALIGN.CENTER)
        if i < 3:
            ax = l + cw + Pt(4)
            tf_a = tb(sl, ax, top+h/2-Pt(14), gap_conn-Pt(8), Pt(26))
            para(tf_a, "->", font=FU, size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)
            tf_cl = tb(sl, ax, top+h/2+Pt(14), gap_conn-Pt(8), Pt(26))
            para(tf_cl, conn_labels[i], font=FU, size=8, bold=True, color=DIM, align=PP_ALIGN.CENTER)
    rect(sl, Inches(0.7), Inches(5.25), Inches(12), Inches(0.85),
         fill=RGBColor(0xFF,0xF0,0xD8), line=RGBColor(0xD9,0x77,0x06))
    tf_note = tb(sl, Inches(0.85), Inches(5.32), Inches(11.7), Pt(54))
    tf_note.word_wrap = True
    para(tf_note, "No coding. No API tokens. No command line. "
                  "Describe yourself -> Gemini writes the site -> paste one file into GitHub -> flip one switch. Live on the real internet.",
         font=FB, size=12, color=TEXT)
    notes(sl, (
        "[NEIL] Before we touch anything, here's exactly what's about to happen. Four steps. "
        "You fill in four things about yourself — your grade, school, one activity, and a career you're curious about. "
        "That goes into Google Gemini as a prompt. Gemini reads it and writes a complete website — "
        "not just text, the actual code for a real webpage. Then you paste that code into GitHub, "
        "which is a free file-hosting service used by millions of developers. "
        "Flip one switch to turn on GitHub Pages, and your portfolio is live at a real URL. "
        "No coding. No command line. No credit card. That's the whole system."
    ))

def s6_build(prs):
    sl = blank(prs); bg(sl); stripe(sl)
    kicker(sl, "Today's Activity")
    title(sl, "What You're Building")
    orange_rule(sl, Inches(1.95))
    steps = [
        ("01", "Tell Gemini about yourself",
         "Fill in 4 fields -- grade, school, one thing you do, and a career you're curious about. No writing experience needed."),
        ("02", "Gemini writes your entire portfolio site",
         "Not just your bio -- the complete working webpage. HTML, CSS, layout, design, all of it. Ready to publish the second you copy it."),
        ("03", "Paste one file into GitHub",
         "Go to github.com, create a free account, paste your HTML into one file, click two buttons. Your portfolio is live at a real URL."),
    ]
    cw = Inches(3.75); gap = Inches(0.18); top = Inches(2.1); h = Inches(4.6)
    for i, (num, head, desc) in enumerate(steps):
        l = Inches(0.7) + i*(cw+gap)
        card(sl, l, top, cw, h, accent=ORANGE if i==1 else BORDER)
        tf_n = tb(sl, l+Pt(12), top+Pt(12), cw-Pt(24), Pt(52))
        para(tf_n, num, font=FD, size=48, color=ORANGE)
        tf_h = tb(sl, l+Pt(12), top+Pt(70), cw-Pt(24), Pt(42))
        tf_h.word_wrap = True
        para(tf_h, head, font=FU, size=14, bold=True, color=BRIGHT)
        tf_d = tb(sl, l+Pt(12), top+Pt(118), cw-Pt(24), Inches(2.2))
        tf_d.word_wrap = True
        para(tf_d, desc, font=FB, size=12, color=MUTED)
    notes(sl, (
        "[INDIA] Here's what you're walking away with. Step one: four fields, 30 seconds. "
        "Step two — this is the part people usually don't believe until they see it — "
        "Gemini writes the ENTIRE website for you. Not just your bio. "
        "The HTML, the CSS, the layout, the design — a fully working webpage, ready to publish. "
        "Step three: tonight, paste one file into GitHub and your portfolio is live at a real URL. "
        "Free. No expiration. Yours to share anywhere."
    ))

def s7_every_step(prs):
    sl = blank(prs); bg(sl); stripe(sl)
    kicker(sl, "The Exercise -- Start to Finish")
    title(sl, "Every Step, In Order")
    orange_rule(sl, Inches(1.95))
    col_w = Inches(5.7); top = Inches(2.1)
    l1 = Inches(0.7); l2 = l1+col_w+Inches(0.75)
    rect(sl, l1+col_w+Inches(0.37), top, Pt(1), Inches(5.0), fill=BORDER)
    tf_h1 = tb(sl, l1, top, col_w, Pt(20))
    para(tf_h1, "IN THE ROOM  .  Today ~15 min", font=FU, size=10, bold=True, color=GREEN)
    tf_h2 = tb(sl, l2, top, col_w, Pt(20))
    para(tf_h2, "TAKE HOME  .  Tonight ~5 min", font=FU, size=10, bold=True, color=SAP)
    in_steps = [
        ("1", "Open gemini.google.com",        "Works on any phone or laptop"),
        ("2", "Go to the prompt slide",        "Slide 8 -- read it before you type"),
        ("3", "Fill in your 4 fields",         "Grade . School . Activity . Career"),
        ("4", "Run it -- scroll through output","You'll see a complete HTML webpage"),
        ("5", "Select all -> Copy",             "Ctrl+A then Ctrl+C (Cmd on Mac)"),
    ]
    take_steps = [
        ("6",  "Create a free GitHub account",    "github.com -- any email, 2 minutes"),
        ("7",  "New repo -> Public -> Create",    "Click + at top right to start"),
        ("8",  "Add file -> Create new -> index.html","Name it exactly 'index.html'"),
        ("9",  "Paste HTML -> Commit changes",    "Paste everything, click Commit"),
        ("10", "Settings -> Pages -> main -> Save","Live in 60s at yourname.github.io"),
    ]
    row_h = Pt(56)
    for steps, lx, ac in [(in_steps, l1, ORANGE), (take_steps, l2, SAP)]:
        for k, (num, head, detail) in enumerate(steps):
            t = top+Pt(28)+k*row_h
            tf_n = tb(sl, lx, t+Pt(4), Pt(26), Pt(28))
            para(tf_n, num, font=FD, size=22, color=ac, align=PP_ALIGN.CENTER)
            tf_h = tb(sl, lx+Pt(30), t+Pt(3), col_w-Pt(38), Pt(20))
            para(tf_h, head, font=FU, size=12, bold=True, color=BRIGHT)
            tf_d = tb(sl, lx+Pt(30), t+Pt(22), col_w-Pt(38), Pt(18))
            para(tf_d, detail, font=FB, size=10.5, color=MUTED)
            if k < 4:
                rect(sl, lx+Pt(30), t+row_h-Pt(3), col_w-Pt(38), Pt(1), fill=BORDER)
    notes(sl, (
        "[NEIL] Let's walk through all ten steps before we do them so nothing surprises you. "
        "Left column is what we're doing right now — five steps, about 15 minutes. "
        "Open Gemini, use the prompt on slide 8, fill in your four fields, run it, copy the output. "
        "Right column is tonight — five more steps, about 5 minutes. "
        "Create a free GitHub account, make a new public repo, create a file called index.html, "
        "paste your HTML, and turn on Pages. "
        "We'll walk through the right column again at the end. For now — let's start on the left."
    ))

def s8_prompt(prs):
    sl = blank(prs); bg(sl); stripe(sl)
    kicker(sl, "Step 1 -- Open Gemini")
    title(sl, "Copy This Prompt")
    orange_rule(sl, Inches(1.95))
    pb = sl.shapes.add_shape(1, Inches(0.7), Inches(2.08), Inches(2.8), Pt(18))
    pb.fill.solid(); pb.fill.fore_color.rgb = ORANGE; pb.line.fill.background()
    tf_pl = tb(sl, Inches(0.7), Inches(2.08), Inches(2.8), Pt(18))
    para(tf_pl, "PASTE INTO GEMINI.GOOGLE.COM", font=FU, size=7, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
    rect(sl, Inches(0.7), Inches(2.3), Inches(12), Inches(3.0), fill=RGBColor(0xF0,0xF4,0xFF), line=BORDER)
    lines = [
        [("I'm a ", MUTED), ("[GRADE]", ORANGE), (" student at ", MUTED), ("[SCHOOL NAME]", ORANGE), (".", MUTED)],
        [("Outside of school I ", MUTED), ("[one activity -- job, sport, hobby, or club]", ORANGE), (".", MUTED)],
        [("I'm curious about careers in ", MUTED), ("[CAREER FIELD]", ORANGE), (".", MUTED)],
        [],
        [("Write me a complete, self-contained HTML portfolio with: header, bio, skills tags, career+AI section", MUTED)],
        [],
        [("// Rules: All CSS in <style> tag. Dark background. Modern. Return ONLY the HTML code.", DIM)],
    ]
    line_t = Inches(2.38)
    for line in lines:
        if not line: line_t += Pt(5); continue
        tf_l = tb(sl, Inches(0.9), line_t, Inches(11.6), Pt(16))
        p = tf_l.add_paragraph()
        for text, color in line:
            r = p.add_run(); r.text = text
            r.font.name = FM; r.font.size = Pt(11.5); r.font.color.rgb = color
            r.font.bold = (color == ORANGE); r.font.italic = (color == DIM)
        line_t += Pt(17)
    actions = [
        ("1", "Open gemini.google.com",       "Sign in with any Google account"),
        ("2", "Paste & fill orange parts",    "Replace each bracket with YOUR real info"),
        ("3", "Copy ALL the HTML output",     "Ctrl+A then Ctrl+C -- save for tonight"),
    ]
    aw = Inches(3.85); at = Inches(5.5)
    for i, (num, head, detail) in enumerate(actions):
        al = Inches(0.7)+i*(aw+Inches(0.12))
        card(sl, al, at, aw, Inches(0.8))
        tf_n = tb(sl, al+Pt(10), at+Pt(8), Pt(22), Pt(24))
        para(tf_n, num, font=FD, size=18, color=ORANGE)
        tf_h = tb(sl, al+Pt(34), at+Pt(7), aw-Pt(44), Pt(20))
        para(tf_h, head, font=FU, size=11, bold=True, color=BRIGHT)
        tf_d = tb(sl, al+Pt(34), at+Pt(26), aw-Pt(44), Pt(20))
        para(tf_d, detail, font=FB, size=10, color=MUTED)
    notes(sl, (
        "[INDIA] This is the prompt. See the orange brackets? Those are the four things you fill in. "
        "Your grade — like '10th grade'. Your school name. One activity — a job, a sport, a club, "
        "anything you do outside school. And a career field you're curious about. That's it. "
        "Fill in the four orange parts, paste the whole thing into gemini.google.com, hit Enter. "
        "The last line is key — it tells Gemini to return ONLY the HTML code, so you get a clean file "
        "you can paste directly into GitHub with nothing to edit. "
        "[Point at slide] Read it once, then fill in your own info. Don't overthink it."
    ))

def s9_demo(prs):
    sl = blank(prs); bg(sl); stripe(sl)
    kicker(sl, "Facilitator Demo -- Watch the Full Flow")
    title(sl, "Watch Once,\nDo It Yourself")
    orange_rule(sl, Inches(2.5))
    items = [
        ("Open gemini.google.com",                                                        ORANGE, False),
        ("Paste the prompt -- fill in grade, school, activity, career",                   ORANGE, False),
        ("Show the HTML output -- scroll so students see what code looks like",            ORANGE, False),
        ("Select all -> Copy (this is what gets pasted into GitHub)",                     ORANGE, False),
        ("Go to github.com -> New repo -> Public -> Add file -> Create new file -> index.html", SAP, True),
        ("Paste HTML -> Commit -> Settings -> Pages -> Source: main -> Save -> Live in 60s",  SAP, True),
    ]
    row_h = Inches(0.62); t = Inches(2.65)
    for text, ac, hi in items:
        bg_c = SURF2 if hi else SURF
        b_c  = SAP if hi else BORDER
        rect(sl, Inches(0.7), t, Inches(12), row_h-Pt(5), fill=bg_c, line=b_c)
        tf_a = tb(sl, Inches(0.85), t+Pt(8), Pt(22), Pt(22))
        para(tf_a, "->", font=FU, size=14, bold=True, color=ac)
        tf_t = tb(sl, Inches(1.15), t+Pt(9), Inches(11.4), Pt(22))
        para(tf_t, text, font=FB, size=13, bold=hi, color=BRIGHT if hi else TEXT)
        t += row_h
    rect(sl, Inches(0.7), t+Pt(4), Inches(12), Pt(28),
         fill=SURF2, line=SAP)
    tf_note = tb(sl, Inches(0.85), t+Pt(6), Inches(11.7), Pt(22))
    para(tf_note, "No API tokens. No command line. Just github.com in a browser -- anyone can do this in 5 minutes.",
         font=FU, size=11, bold=True, color=SAP)
    notes(sl, (
        "[NEIL — LIVE DEMO] Watch this once — I'm going to run through the full flow live "
        "so you can see every step before you try it yourself.\n\n"
        "GEMINI STEPS:\n"
        "- Open gemini.google.com in the browser\n"
        "- Paste the prompt with real info filled in (use your own grade/school/activity/career)\n"
        "- Show the HTML output scrolling — say: 'See this? This is a complete webpage. "
        "Gemini just wrote it in about 10 seconds.'\n"
        "- Select all (Ctrl+A), Copy (Ctrl+C)\n\n"
        "GITHUB STEPS:\n"
        "- Go to github.com — say: 'This is just a website, no software to install'\n"
        "- Click + > New repository > name it 'portfolio' > Public > Create repository\n"
        "- Click Add file > Create new file > type 'index.html'\n"
        "- Paste the HTML > scroll down > click Commit changes\n"
        "- Settings > Pages > Source: main > Save\n"
        "- Wait ~60 seconds, show the live URL\n\n"
        "Say: 'That's a real website. Anyone with that URL can see it. Now you do it.'"
    ))

def s10_your_turn(prs):
    sl = blank(prs); bg(sl); stripe(sl)
    kicker(sl, "Hands-On Activity")
    title(sl, "Your Turn", size=52)
    orange_rule(sl, Inches(2.1))
    tf_timer = tb(sl, Inches(0.7), Inches(2.25), Inches(5), Inches(1.4))
    para(tf_timer, "5 Minutes", font=FD, size=72, color=ORANGE)
    tf_sub = tb(sl, Inches(0.7), Inches(3.55), Inches(5), Pt(18))
    para(tf_sub, "YOU'VE GOT THIS", font=FU, size=10, bold=True, color=MUTED)
    steps = [
        ("1", "Open Gemini",         "gemini.google.com on your phone or laptop"),
        ("2", "Go to the prompt",    "Slide 8 -- read it once before you start"),
        ("3", "Fill in 4 fields",    "Replace each orange bracket with something real about you"),
        ("4", "Copy the HTML output","Select all, copy -- that's your complete portfolio site"),
    ]
    sw = Inches(5.6); gap = Inches(0.14); top = Inches(3.85)
    for i, (num, head, detail) in enumerate(steps):
        col = i % 2; row = i // 2
        l = Inches(0.7)+col*(sw+gap)
        t = top+row*Inches(1.45)
        card(sl, l, t, sw, Inches(1.3))
        tf_n = tb(sl, l+Pt(10), t+Pt(8), Pt(30), Pt(34))
        para(tf_n, num, font=FD, size=28, color=ORANGE)
        tf_h = tb(sl, l+Pt(44), t+Pt(8), sw-Pt(54), Pt(20))
        para(tf_h, head, font=FU, size=13, bold=True, color=BRIGHT)
        tf_d = tb(sl, l+Pt(44), t+Pt(30), sw-Pt(54), Pt(32))
        tf_d.word_wrap = True
        para(tf_d, detail, font=FB, size=11, color=MUTED)
    tf_help = tb(sl, Inches(0.7), Inches(6.95), Inches(8), Pt(20))
    para(tf_help, "Need help? Raise your hand -- SAP team members are here.", font=FU, size=11, bold=True, color=SAP)
    notes(sl, (
        "[INDIA] Five minutes. Open Gemini on your phone or laptop — gemini.google.com. "
        "Go to slide 8 for the prompt, fill in your four fields with real info, run it, "
        "and copy all the HTML output. That's everything you need to do right now. "
        "If you get stuck at any point, raise your hand — we have SAP team members in the room. "
        "Don't worry about it being perfect. You can always regenerate with different info later.\n\n"
        "[NEIL — while students work] If you finish early, go back into Gemini and try a different "
        "career field. See how the portfolio changes. Or try making the prompt more specific — "
        "add your GPA, a specific club role, a specific company you want to work for."
    ))

def s11_take_home(prs):
    sl = blank(prs); bg(sl); stripe(sl)
    kicker(sl, "Take-Home Challenge")
    title(sl, "Your Portfolio Goes\nLive Tonight", size=40)
    orange_rule(sl, Inches(2.35))
    steps = [
        ("1", "Create a free GitHub account",              "github.com -- takes 2 min, any email works",        None),
        ("2", "New repository -> Public -> Create",       "Click + at top right to start",                     None),
        ("3", "Add file -> Create new file -> index.html","Name the file exactly 'index.html'",                None),
        ("4", "Paste your HTML from Gemini -> Commit",    "Paste everything, scroll down, click Commit changes", None),
        ("5", "Settings -> Pages -> Source: main -> Save","Wait ~60 seconds -- then visit your URL",            GREEN),
    ]
    row_h = Inches(0.72); t = Inches(2.5)
    for num, head, detail, ac in steps:
        bg_c = RGBColor(0xEC,0xFF,0xF2) if ac else SURF
        b_c  = GREEN if ac else BORDER
        rect(sl, Inches(0.7), t, Inches(12), row_h-Pt(5), fill=bg_c, line=b_c)
        tf_n = tb(sl, Inches(0.76), t+Pt(8), Pt(30), Pt(26))
        para(tf_n, num, font=FD, size=22, color=GREEN if ac else ORANGE)
        tf_h = tb(sl, Inches(1.2), t+Pt(8), Inches(5.2), Pt(20))
        para(tf_h, head, font=FU, size=12, bold=True, color=BRIGHT)
        tf_d = tb(sl, Inches(6.5), t+Pt(9), Inches(6.1), Pt(20))
        para(tf_d, detail, font=FB, size=11, color=MUTED)
        t += row_h
    rect(sl, Inches(0.7), t+Pt(4), Inches(12), Pt(30),
         fill=RGBColor(0xEC,0xFF,0xF2), line=GREEN)
    tf_url_l = tb(sl, Inches(0.85), t+Pt(7), Inches(1.2), Pt(22))
    para(tf_url_l, "YOUR URL", font=FU, size=9, bold=True, color=GREEN)
    tf_url_v = tb(sl, Inches(2.1), t+Pt(6), Inches(10), Pt(24))
    para(tf_url_v, "yourname.github.io / your-repo-name", font=FM, size=14, color=BRIGHT)
    notes(sl, (
        "[NEIL] You've got the HTML — the hard part is done. Tonight, five steps to get it live. "
        "One: create a free GitHub account at github.com. Any email, takes two minutes. "
        "Two: click the plus sign at top right, create a new repository, make it Public, hit Create. "
        "Three: click Add file, Create new file, and name it exactly 'index.html' — "
        "lowercase, spelling matters, that exact filename. "
        "Four: paste your entire HTML from Gemini into the file editor, scroll to the bottom, "
        "click Commit changes. "
        "Five: go to Settings, then Pages, set Source to 'main', click Save. "
        "Wait about 60 seconds. Your URL appears — yourname.github.io/your-repo-name. "
        "Screenshot it. Put it in your college applications, your LinkedIn, your email signature. "
        "It's yours and it's free forever."
    ))

def s12_close(prs):
    sl = blank(prs); bg(sl); stripe(sl)
    rect(sl, 0, H-Pt(4), W, Pt(4), fill=ORANGE)
    tf_q = tb(sl, Inches(0.7), Inches(1.6), Inches(10), Inches(3.2))
    tf_q.word_wrap = True
    para(tf_q, "You just used AI", font=FD, size=76, color=BRIGHT)
    para(tf_q, "like a professional.", font=FD, size=76, color=ORANGE)
    rule = sl.shapes.add_shape(1, Inches(0.7), Inches(5.1), Inches(0.72), Pt(3))
    rule.fill.solid(); rule.fill.fore_color.rgb = ORANGE; rule.line.fill.background()
    tf_b = tb(sl, Inches(0.7), Inches(5.35), Inches(9), Inches(0.9))
    tf_b.word_wrap = True
    para(tf_b, "The real skill isn't coding. It's knowing what to ask, how to judge the output, "
              "and when to trust it -- and when not to. That judgment is what every employer is looking for.",
         font=FB, size=16, color=MUTED)
    tf_link = tb(sl, Inches(0.7), Inches(6.5), Inches(6), Pt(22))
    para(tf_link, "-> github.com/neilgbp/ai-portfolio", font=FM, size=13, color=ORANGE)
    notes(sl, (
        "[INDIA] What you just did isn't a party trick. You gave an AI the right instructions, "
        "it produced professional output, and you can deploy it to the real internet. "
        "That is the same workflow a product manager, a marketing director, a healthcare analyst, "
        "or a startup founder uses today. The skill employers are paying for isn't knowing how to code. "
        "It's knowing what to ask, knowing how to judge what comes back, and knowing when to trust it "
        "and when to question it. You just practiced that.\n\n"
        "[NEIL] Everything from today is at github.com/neilgbp/ai-portfolio — "
        "the prompts, the deployment steps, all of it. Free. Grab it. "
        "And if you want to connect on LinkedIn or have questions later, find us. "
        "You've got this. Go build something."
    ))


def main():
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    s1_title(prs)
    s2_big_picture(prs)
    s3_session(prs)
    s4_ai_careers(prs)
    s5_architecture(prs)
    s6_build(prs)
    s7_every_step(prs)
    s8_prompt(prs)
    s9_demo(prs)
    s10_your_turn(prs)
    s11_take_home(prs)
    s12_close(prs)
    out = "/Users/I751845/Desktop/ai-portfolio/LOT_Atlanta_SAP_Portfolio_Session.pptx"
    prs.save(out)
    print(f"Saved: {out}  ({len(prs.slides)} slides)")

if __name__ == "__main__":
    main()
