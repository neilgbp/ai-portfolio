#!/usr/bin/env python3
"""Build the LOT Atlanta x SAP PPTX from presentation.html.

Single source of truth = presentation.html. This renders each slide of the real
deck at 1280x720 (2x for crispness), lays each in full-bleed on a 16:9 slide, and
attaches the presenter talk track as speaker notes. Re-run any time the HTML changes.
"""
import functools
import http.server
import socketserver
import tempfile
import threading
import time
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
HTML = "presentation.html"
OUT = HERE / "LOT_Atlanta_SAP_Portfolio_Session.pptx"
N = 12

# Presenter talk track — [NEIL] / [INDIA] tag each line. Attached as speaker notes.
NOTES = [
    # 1 — Title
    "[NEIL] Hey everyone - we're Neil and India from SAP. We have 30 minutes and we're "
    "going to use all of it. By the time we're done, every one of you will have a real "
    "AI-generated portfolio page - the kind you can link to on a college application or "
    "LinkedIn. No coding experience, no software to install, nothing to pay for. "
    "Two free tools. Let's go.",
    # 2 — Big Picture
    "[NEIL or INDIA] Quick map of the morning. You've already done the first two blocks. "
    "We're the orange one - the hands-on build. After us: a responsible AI segment, then "
    "closing. Our 30 minutes. Everything we're doing is built on two free tools: Google "
    "Gemini for the AI, and GitHub to host the website. You've probably heard of both. "
    "You're about to use them together to build something real.",
    # 3 — Session
    "[INDIA] Here's our 30 minutes. Five minutes warming up, five on AI and careers, "
    "fifteen minutes building - that's the main event - and five minutes to share. "
    "The key idea on this slide: employers don't care that you went to school. They want "
    "to see what you built. A portfolio page with your actual name, skills, and career "
    "thinking stands out from every other applicant who just submits a transcript.",
    # 4 — AI careers
    "[NEIL] Show of hands - who's thinking about healthcare? Business? Creative work? "
    "Sports? Engineering? [pause] Here's what's true for every one of those fields: AI is "
    "already in them. AI reads medical scans faster than any one doctor. AI writes "
    "contracts and runs financial models in seconds. AI generates music, images, video. "
    "AI tracks athlete performance in real time and predicts injuries before they happen. "
    "AI designs physical structures and simulates climate systems. This isn't a future "
    "thing - it's happening now. The students who learn to work WITH AI, not just use it, "
    "will have a real edge. That's what we're showing you today.",
    # 5 — Architecture
    "[NEIL] Before we touch anything, here's exactly what's about to happen. Four steps. "
    "You fill in four things about yourself - your grade, school, one activity, and a "
    "career you're curious about. That goes into Google Gemini as a prompt. Gemini reads "
    "it and writes a complete website - not just text, the actual code for a real webpage. "
    "Then you paste that code into GitHub, a free file-hosting service used by millions of "
    "developers. Flip one switch to turn on GitHub Pages, and your portfolio is live at a "
    "real URL. No coding. No command line. No credit card. That's the whole system.",
    # 6 — Build
    "[INDIA] Here's what you're walking away with. Step one: four fields, 30 seconds. "
    "Step two - this is the part people usually don't believe until they see it - Gemini "
    "writes the ENTIRE website for you. Not just your bio. The HTML, the CSS, the layout, "
    "the design - a fully working webpage, ready to publish. Step three: tonight, paste "
    "one file into GitHub and your portfolio is live at a real URL. Free. No expiration. "
    "Yours to share anywhere.",
    # 7 — Every step
    "[NEIL] Let's walk through all ten steps before we do them so nothing surprises you. "
    "Left column is what we're doing right now - five steps, about 15 minutes. Open "
    "Gemini, use the prompt on slide 8, fill in your four fields, run it, copy the output. "
    "Right column is tonight - five more steps, about 5 minutes. Create a free GitHub "
    "account, make a new public repo, create a file called index.html, paste your HTML, "
    "and turn on Pages. We'll walk through the right column again at the end. For now - "
    "let's start on the left.",
    # 8 — Prompt
    "[INDIA] This is the prompt. See the orange brackets? Those are the four things you "
    "fill in. Your grade - like '10th grade'. Your school name. One activity - a job, a "
    "sport, a club, anything you do outside school. And a career field you're curious "
    "about. That's it. Fill in the four orange parts, paste the whole thing into "
    "gemini.google.com, hit Enter. The last line is key - it tells Gemini to return ONLY "
    "the HTML code, so you get a clean file you can paste straight into GitHub with "
    "nothing to edit. [Point at slide] Read it once, then fill in your own info. Don't "
    "overthink it.",
    # 9 — Demo
    "[NEIL - LIVE DEMO] Watch this once - I'll run the full flow live so you see every "
    "step before you try it.\n\n"
    "GEMINI STEPS:\n"
    "- Open gemini.google.com in the browser\n"
    "- Paste the prompt with real info filled in (your grade/school/activity/career)\n"
    "- Scroll the HTML output - say: 'See this? A complete webpage. Gemini wrote it in "
    "about 10 seconds.'\n"
    "- Select all (Ctrl+A), Copy (Ctrl+C)\n\n"
    "GITHUB STEPS:\n"
    "- Go to github.com - say: 'Just a website, no software to install'\n"
    "- Click + > New repository > name it 'portfolio' > Public > Create repository\n"
    "- Add file > Create new file > type 'index.html'\n"
    "- Paste the HTML > scroll down > Commit changes\n"
    "- Settings > Pages > Source: main > Save\n"
    "- Wait ~60 seconds, show the live URL\n\n"
    "Say: 'That's a real website. Anyone with that URL can see it. Now you do it.'",
    # 10 — Your turn
    "[INDIA] Five minutes. Open Gemini on your phone or laptop - gemini.google.com. Go to "
    "slide 8 for the prompt, fill in your four fields with real info, run it, and copy all "
    "the HTML output. That's everything you need to do right now. Stuck anywhere? Raise "
    "your hand - we have SAP team members in the room. Don't worry about it being perfect; "
    "you can always regenerate with different info later.\n\n"
    "[NEIL - while students work] Finished early? Go back into Gemini and try a different "
    "career field - see how the portfolio changes. Or make the prompt more specific: add "
    "your GPA, a specific club role, a company you want to work for.",
    # 11 — Take home
    "[NEIL] You've got the HTML - the hard part is done. Tonight, five steps to get it "
    "live. One: create a free GitHub account at github.com. Any email, two minutes. Two: "
    "click the plus sign at top right, new repository, make it Public, Create. Three: Add "
    "file, Create new file, name it exactly 'index.html' - lowercase, spelling matters. "
    "Four: paste your entire HTML from Gemini, scroll to the bottom, Commit changes. Five: "
    "Settings, then Pages, set Source to 'main', Save. Wait about 60 seconds. Your URL "
    "appears - yourname.github.io/your-repo-name. Screenshot it. Put it on your college "
    "apps, your LinkedIn, your email signature. It's yours and it's free forever.",
    # 12 — Close
    "[INDIA] What you just did isn't a party trick. You gave an AI the right instructions, "
    "it produced professional output, and you can deploy it to the real internet. That's "
    "the same workflow a product manager, a marketing director, a healthcare analyst, or a "
    "startup founder uses today. The skill employers pay for isn't knowing how to code - "
    "it's knowing what to ask, how to judge what comes back, and when to trust it versus "
    "question it. You just practiced that.\n\n"
    "[NEIL] Everything from today is at github.com/neilgbp/ai-portfolio - the prompts, the "
    "deployment steps, all of it. Free. Grab it. Want to connect on LinkedIn or have "
    "questions later? Find us. You've got this. Go build something.",
]


def serve(directory):
    handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=str(directory))
    httpd = socketserver.TCPServer(("127.0.0.1", 0), handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    return httpd, httpd.server_address[1]


def render_slides(url, outdir):
    shots = []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1280, "height": 720}, device_scale_factor=2)
        page.goto(url, wait_until="load")
        page.evaluate("document.body.classList.add('export'); if (typeof fit==='function') fit();")
        page.evaluate("async () => { try { await document.fonts.ready; } catch(e){} }")
        time.sleep(0.8)
        deck = page.locator(".deck")
        for i in range(N):
            page.evaluate(f"go({i})")
            time.sleep(0.25)
            fp = outdir / f"slide_{i+1:02d}.png"
            deck.screenshot(path=str(fp))
            shots.append(fp)
        browser.close()
    return shots


def main():
    httpd, port = serve(HERE)
    url = f"http://127.0.0.1:{port}/{HTML}"
    with tempfile.TemporaryDirectory() as tmp:
        shots = render_slides(url, Path(tmp))
        httpd.shutdown()

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank = prs.slide_layouts[6]
        for i, fp in enumerate(shots):
            sl = prs.slides.add_slide(blank)
            sl.shapes.add_picture(str(fp), 0, 0, width=prs.slide_width, height=prs.slide_height)
            sl.notes_slide.notes_text_frame.text = NOTES[i]
        prs.save(str(OUT))
    print(f"Saved: {OUT}  ({len(prs.slides)} slides, with speaker notes)")


if __name__ == "__main__":
    main()
