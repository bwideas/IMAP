"""Data Centre Trade Media Report, July 2026 - Saudi Projects (TPG Media) + Gulf Construction.

Run from anywhere:  python build_report.py
Output: Grundfos_Data_Centre_Trade_Media_Report_July2026.pptx next to this file.
Needs python-pptx and the brand kit in ../../brand.
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1] / "brand" / "tools"))
from grundfos_pptx import *  # noqa: E402

OUT = HERE / "Grundfos_Data_Centre_Trade_Media_Report_July2026.pptx"
ARTICLE_IMG = HERE / "assets" / "saudi_projects_spread.jpg"

# ------------------------------------------------------------------ data (from the publishers' reports)
SP = dict(contents=7992, issuu=15831, web_views=12026, web_unique=10854, web_clicks=827,
          carousel=172317, nl_recipients=212886, nl_unique=102866, followers=195700,
          likes=31, comments=2, reposts=14)
SP["total"] = SP["contents"] + SP["issuu"] + SP["web_views"] + SP["carousel"] + SP["nl_recipients"] + SP["followers"]
SP["engagements"] = SP["likes"] + SP["comments"] + SP["reposts"]
SP["ctr"] = SP["web_clicks"] / SP["web_views"]
SP["article_views"] = SP["issuu"] + SP["web_views"]

GC = dict(edm=37112, open_rate=0.1611, views=188, unique=173, ctr=0.0051, social=1488)
GC["opens"] = round(GC["edm"] * GC["open_rate"])
GC["total"] = GC["edm"] + GC["views"] + GC["social"]
GEO = [("Saudi Arabia", 20), ("UAE", 15), ("MENA (other)", 13), ("Qatar", 10),
       ("Kuwait", 5), ("Oman", 4), ("Bahrain", 3), ("Rest of world", 30)]
GCC = 20 + 15 + 10 + 5 + 4 + 3
MENA = GCC + 13

TOTAL = SP["total"] + GC["total"]
EMAIL = SP["nl_recipients"] + GC["edm"]
VIEWS = SP["article_views"] + GC["views"]
UNIQUE = SP["web_unique"] + GC["unique"]
SP_SHARE = SP["total"] / TOTAL

assert SP["total"] == 616752 and GC["total"] == 38788 and TOTAL == 655540
n = fmt

prs = load_template()

# ================================================================== 1. Cover
new_cover(prs, "Data Centre Trade Media Report",
          ["Saudi Projects and Gulf Construction articles", "NAS Marketing  |  July 2026"])

# ================================================================== 2. Summary
s = new_content_slide(prs, "Summary – two data centre articles, July 2026")
textbox(s, 0.58, 1.3, 12.17, 0.75, [
    "In July we ran two features on data centre cooling: one with Saudi Projects (TPG Media) and one with "
    "Gulf Construction. Both publishers have now sent their results. The main numbers are below; the "
    "following two slides show each publisher in more detail.",
], size=12, line_spacing=1.1)

rows = [
    ["", "Saudi Projects (TPG Media)", "Gulf Construction", "Both articles"],
    ["Article", "Advancing Data Centre Efficiency in Saudi Arabia", "Smart cooling key to data centre expansion: Grundfos", ""],
    ["Email sent to", n(SP["nl_recipients"]), n(GC["edm"]), n(EMAIL)],
    ["Email open rate", "48%", "16.11%", "–"],
    ["Article page views", f"{n(SP['article_views'])}  (website {n(SP['web_views'])} + digital issue {n(SP['issuu'])})", n(GC["views"]), n(VIEWS)],
    ["Unique readers", n(SP["web_unique"]), n(GC["unique"]), n(UNIQUE)],
    ["Click-throughs", f"{n(SP['web_clicks'])}  ({SP['ctr']:.1%} of page views)", "0.51% click-through rate", "–"],
    ["Social media", f"{n(SP['followers'])}+ LinkedIn followers, {SP['engagements']} engagements", f"{n(GC['social'])} views", "–"],
    ["Total reach", n(SP["total"]), n(GC["total"]), n(TOTAL)],
]
table(s, 0.58, 2.15, 12.17, [1.9, 4.2, 4.0, 2.07], rows, size=10.5, row_h=0.47, header_h=0.45,
      align={3: PP_ALIGN.RIGHT}, bold_first_col=True)
textbox(s, 0.58, 6.45, 12.17, 0.5, [
    [("Note: ", {"bold": True, "color": MUTED}),
     ("total reach is the publishers' gross figure – impressions, email recipients and followers added together, "
      f"not de-duplicated. Saudi Projects accounts for {SP_SHARE:.0%} of it.", {"color": MUTED})],
], size=9.5)

# ================================================================== 3. Saudi Projects
s = new_content_slide(prs, "Saudi Projects (TPG Media) – results")
textbox(s, 0.58, 1.3, 7.9, 0.55, [
    "“Advancing Data Centre Efficiency in Saudi Arabia” – interview with Waleed Okasha, Grundfos Saudi Arabia. "
    "Published in digital issue 67 (July/August 2026), pages 82–85, and promoted on the website, newsletter and LinkedIn.",
], size=11, color=MUTED, line_spacing=1.1)
rows = [
    ["Channel", "Result", "Time on page"],
    ["Digital issue – contents page", f"{n(SP['contents'])} page impressions", "48 sec"],
    ["Digital issue – article (2 spreads)", f"{n(SP['issuu'])} page impressions", "23 sec / 1 min 12 sec"],
    ["Website – article page", f"{n(SP['web_views'])} page views, {n(SP['web_unique'])} unique visitors, {n(SP['web_clicks'])} click-throughs", "56 sec"],
    ["Website – home page carousel", f"{n(SP['carousel'])} page impressions", "19 sec"],
    ["Email newsletter – content banner", f"{n(SP['nl_recipients'])} recipients, 48% open rate ({n(SP['nl_unique'])} unique impressions)", "–"],
    ["LinkedIn – Saudi Projects and Emirates Projects", f"{n(SP['followers'])}+ followers, {SP['engagements']} engagements ({SP['likes']} likes, {SP['comments']} comments, {SP['reposts']} reposts)", "–"],
    ["Total reach", n(SP["total"]), ""],
]
table(s, 0.58, 2.0, 7.9, [2.75, 3.75, 1.4], rows, size=10, row_h=0.56, header_h=0.4,
      align={2: PP_ALIGN.CENTER})

img_x, img_y, img_w = 8.85, 1.3, 3.9
pic = add_picture(s, ARTICLE_IMG, img_x, img_y, w=img_w, name="Article image")
img_h = pic.height / 914400
textbox(s, img_x, img_y + img_h + 0.05, img_w, 0.3,
        ["Opening spread of the article, pages 82–83"], size=9, color=MUTED)
py = img_y + img_h + 0.5
table_bottom = 2.0 + 0.4 + 0.56 * 7
panel(s, img_x, py, img_w, table_bottom - py)
label(s, img_x + 0.15, py + 0.12, img_w - 0.3, "What stood out", size=12)
bullets(s, img_x + 0.15, py + 0.5, img_w - 0.3, table_bottom - py - 0.6, [
    f"{n(SP['web_clicks'])} readers clicked through from the article page – {SP['ctr']:.1%} of page views.",
    "Readers spent over a minute (1 min 12 sec) on the main article spread.",
    f"The newsletter banner reached {n(SP['nl_recipients'])} recipients with a 48% open rate.",
], size=10.5, space_after=6)

# ================================================================== 4. Gulf Construction
s = new_content_slide(prs, "Gulf Construction – results")
textbox(s, 0.58, 1.3, 6.3, 0.55, [
    "“Smart cooling key to data centre expansion: Grundfos” – sent by email to Gulf Construction's B2B readers "
    "and shared on their website and social channels.",
], size=11, color=MUTED, line_spacing=1.1)
rows = [
    ["Channel", "Result"],
    ["Email (EDM)", f"{n(GC['edm'])} recipients, 16.11% open rate – about {n(GC['opens'])} opens"],
    ["Website – article page", f"{n(GC['views'])} page views, {GC['unique']} unique users, 0.51% click-through rate"],
    ["Social media", f"{n(GC['social'])} views across Gulf Construction channels"],
    ["Total reach", n(GC["total"])],
]
table(s, 0.58, 2.0, 6.3, [2.1, 4.2], rows, size=10.5, row_h=0.56, header_h=0.4)
py = 2.0 + 0.4 + 0.56 * 4 + 0.3
BOTTOM = 6.6
panel(s, 0.58, py, 6.3, BOTTOM - py)
label(s, 0.73, py + 0.12, 6.0, "In short", size=12)
bullets(s, 0.73, py + 0.5, 6.0, BOTTOM - py - 0.6, [
    f"Small in numbers – {n(GC['total'])} touchpoints, about {1 - SP_SHARE:.0%} of the combined reach.",
    f"But well targeted: {MENA}% of readers are in the MENA region and {GCC}% in the GCC, with Saudi Arabia the largest market at 20%.",
    "Open rate (16.11%) is roughly a third of Saudi Projects' newsletter (48%).",
], size=10.5, space_after=6)

cx, cy, cw, ch = 7.3, 1.3, 5.45, BOTTOM - 1.3
label(s, cx, cy, cw, "Where the readers are (share of online audience)", size=12)
bar_chart(s, cx, cy + 0.35, cw, ch - 0.35, [g[0] for g in GEO], [g[1] / 100 for g in GEO],
          number_format="0%", max_scale=0.36)

# ================================================================== 5. Takeaways
s = new_content_slide(prs, "What we learned and next steps")
label(s, 0.58, 1.5, 6.0, "What we learned", size=13)
bullets(s, 0.58, 1.95, 6.0, 4.5, [
    [("Saudi Projects gave us the scale. ", {"bold": True}),
     f"It delivered {SP_SHARE:.0%} of the total reach and all {SP['web_clicks']} measured click-throughs. The newsletter "
     "(48% open rate) and the home page carousel did most of the work."],
    [("Gulf Construction is small but well targeted. ", {"bold": True}),
     f"Only {n(GC['total'])} touchpoints, but {MENA}% of its readers are in MENA, so the article reached the right people."],
    [("Email was the strongest channel for both. ", {"bold": True}),
     f"Around {n(SP['nl_unique'] + GC['opens'])} opens from {n(EMAIL)} recipients. Social media added reach but little measurable engagement."],
    [("Readers stayed with the content. ", {"bold": True}),
     "Over a minute on the article spread and close to a minute on the web page – good for a technical topic."],
], size=13, space_after=14)

panel(s, 7.0, 1.35, 5.75, 4.2)
label(s, 7.25, 1.5, 5.3, "Next steps", size=13)
bullets(s, 7.25, 1.95, 5.3, 3.5, [
    "Share both articles on Grundfos LinkedIn and with the sales teams in Saudi Arabia and the Gulf.",
    "Ask both publishers for tracked links (UTM) on future features so we can follow clicks through to leads.",
    "Keep Saudi Projects as the lead title for data centre topics; use Gulf Construction when we need GCC-wide coverage.",
    "Plan the next feature around the same themes – cooling efficiency, water and energy savings, and Saudi data centre growth.",
], size=13, space_after=14)

prs.save(OUT)
print(f"saved {OUT} | SP {SP['total']:,} | GC {GC['total']:,} | total {TOTAL:,}")
