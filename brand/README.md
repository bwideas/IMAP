# Grundfos brand kit for reports

Everything needed to build a Grundfos-branded PowerPoint report that looks like the ones our
team already uses: the template, the logo files, the colours, the fonts to ask for, and a small
Python helper that puts it all together. Read this page once; after that, copy an existing
report folder and change the numbers.

**Rule of thumb:** one template, two blues, one font family, plain words, key numbers only.

---

## 1. Who we are (the short version)

Grundfos is a Danish pump and water-technology company – the world's largest pump manufacturer –
founded in 1945 by Poul Due Jensen in Bjerringbro, Denmark. It is majority-owned (about 86%) by
the Poul Due Jensen Foundation, employs around 20,800 people in more than 60 countries, produces
more than 15 million pump units a year and had revenue of about EUR 4.5 bn in 2024. The CEO is
Poul Due Jensen, appointed in 2020.

**Purpose:** pioneering solutions to the world's water and climate challenges and improving
quality of life for people.

**Brand promise (2022 refresh):** to *respect, protect and advance the flow of water* – for
people, for buildings, for industry and for the planet.

**Endline:** **Possibility in every drop** – introduced with the brand refresh on 30 Sept 2022
and part of the logo lock-up on our covers.

**Our team:** NAS Marketing (Grundfos Saudi Arabia – GSA, and Grundfos Egypt – GEG) for
Commercial Building Services (CBS). We work with consultants, contractors and end users on
HVAC, pressure boosting, water, wastewater and aftermarket. Priority verticals are data centres,
hospitality and healthcare. Typical activities: CPD-certified webinars and workshops, lunch and
learns, customer factory trips, HVAC Day and Engineering Day events, Ramadan iftars, product
launches, Marketo always-on emailers, LinkedIn campaigns, the AOC programme and trade-media
features. Reports are read by the marketing manager, the sales leaders and regional marketing –
busy people who want the result, not the process.

---

## 2. Logo

Files in `assets/`:

| File | Use |
|---|---|
| `grundfos-logo-endline.svg` / `.png` | Full logo with "Possibility in every drop". Covers only, bottom-right, white on the dark photo. |
| `grundfos-mark.svg` / `.png` | The "X" mark alone. Top-right of every content slide (template position 12.43 in from the left, 0.58 in from the top, 0.33 × 0.30 in). |

Rules:

- Keep the clear space around the logo free of text and pictures – at least the height of the
  "X" on every side. Never crop, stretch, rotate, recolour or add effects.
- One logo per slide. Content slides carry the mark only; the full logo with endline belongs on
  the cover (and a closing slide if you use one).
- Partner or publisher logos (TPG Media, Gulf Construction …) go in the body of the slide, never
  next to our mark in the header.
- In our reports the logos are **ordinary slide objects**, not locked into the master, so they
  can be moved, resized or deleted on any slide. Keep it that way – see section 8.

---

## 3. Colours

From the "Grundfos 2022" theme (`assets/theme-colours.json`). Use the first three for almost
everything.

| Role | Hex | Where |
|---|---|---|
| Dark blue (accent 1) | `#11497B` | Table headers, chart bars, small headings, bullet dots |
| Bright blue (accent 2) | `#126AF3` | Slide titles (set by the layout – do not restyle) |
| Pale blue (dk2) | `#E9F1F7` | Commentary panels, total rows |
| Light blue (accent 3) | `#47A6FF` | Only if a second series is unavoidable |
| Navy (accent 4) | `#092844` | Dark-background layouts only |
| Steel blue (lt2) | `#7497B2` | Secondary text on dark slides |
| Rule grey | `#CCCFD7` | Horizontal table lines, chart axis |
| Muted text | `#595959` | Footnotes, captions, source lines |
| Black / white | `#000000` / `#FFFFFF` | Body text / backgrounds |

The theme also has two sand colours (`#E3C181`, `#F0DDBA`). We do not use them in reports.
No red/green "traffic lights", no gradients, no shadows.

---

## 4. Typography

The corporate typeface is **Grundfos TheSans**. The files are licensed to Grundfos, so they are
**not** in this repository – install them from the internal brand portal (they usually live in
`%LOCALAPPDATA%\Microsoft\Windows\Fonts\GrundfosTheSans-*.ttf` on a company laptop).

| Use | Font | Size |
|---|---|---|
| Slide title | Grundfos TheSans **Extd Black** (from the layout) | 28 pt, bright blue |
| Cover title | Grundfos TheSans Extd Black (from the layout) | 40 pt, white |
| Body, tables, chart labels | Grundfos TheSans **SemiLight** | 10–12 pt |
| Emphasis, table headers, totals, small headings | Grundfos TheSans **SemiBold** | same size as the text around it |
| Footnotes, captions | Grundfos TheSans SemiLight, muted grey | 9–9.5 pt |

Emphasis comes from switching to the SemiBold family, never from the synthetic **B** button.
If the fonts are missing PowerPoint falls back to Arial – the layout still holds, but check line
breaks before sending.

---

## 5. Layouts and page geometry

Slides are 16:9, 13.333 × 7.5 in. The template (`templates/Grundfos_2022_Report_Template.pptx`)
keeps all 39 layouts from the corporate master; for reports we use three:

| Layout | Index | Use |
|---|---|---|
| Title Slide | 0 | Cover: dark aerial water photo, white title (0.58, 1.45 in), subtitle (0.58, 3.62 in), logo + endline bottom-right |
| Title Only | 2 | Every content slide: blue title at (0.58, 0.55 in), mark top-right, white background |
| Blank | 3 | Full-bleed pictures or a closing slide |

Content area on a content slide: **x 0.58 → 12.75 in, y 1.3 → 6.6 in**. Keep a 0.58 in margin
on the left and right and leave the bottom 0.9 in empty.

Dark and blue "Title Slide" variants exist in the template for other occasions; reports use the
standard water cover so they all look like one family.

---

## 6. Imagery

- Cover: the aerial water photograph from the template (`assets/cover-water.jpg`). It can be
  replaced with another dark, calm water or infrastructure photo – the title must stay readable
  in white.
- Inside the report: real artefacts only – the published article spread, a screenshot of the
  landing page, the event photo. No stock icons, no clip art, no decorative shapes.
- Give every picture a short grey caption (9 pt) saying what it is.

---

## 7. Report style

Our reports are written for a manager who has five minutes. Checklist:

1. **Five slides is normal**: cover, summary, one slide per channel/publisher/event, "what we
   learned and next steps".
2. **Plain sentences.** "In July we ran two features…" not "This document presents a
   comprehensive multi-channel performance analysis." Write it the way you would say it.
3. **Numbers in tables, words beside them.** One table per slide, dark-blue header, grey rules,
   pale-blue total row. Thousands separators always (`212,886`). Round percentages to whole
   numbers unless the publisher quoted decimals.
4. **One commentary panel per slide** ("What stood out", "In short", "Next steps") on a pale-blue
   rectangle, three bullets at most, each one a fact or a decision.
5. **Charts** are single-colour dark-blue bars with value labels and no gridlines, legend or
   title – the heading above the chart does that job.
6. **Say what the number means.** "Total reach" is the publisher's gross figure; say so in a
   footnote. Never present impressions as readers.
7. **Finish with next steps**, not with a thank-you slide.
8. Two blues and grey. If you are reaching for a third colour, simplify the slide instead.

---

## 8. Editable logos and pictures

In the corporate master the "X" mark, the cover photo and the cover logo are locked (they live
in the slide master / layout, so they cannot be selected on a slide). Our helper hides those
master graphics on each slide (`showMasterSp="0"`) and places identical copies as normal pictures
named **Grundfos mark**, **Cover photo** and **Grundfos logo**. Click, move, resize, swap
(right-click → Change Picture) or delete them like anything else. Do not "fix" this by editing the
master – keep the master untouched so the deck still matches the corporate template.

---

## 9. Building a report

Requirements: Python 3.10+, `pip install python-pptx pillow`. PowerPoint (Windows) is only needed
for the PNG preview.

```python
import sys; sys.path.insert(0, "brand/tools")
from grundfos_pptx import *

prs = load_template()                                   # empty deck on the Grundfos template
new_cover(prs, "Event Report – HVAC Day 2026", ["Riyadh, 9 February", "NAS Marketing  |  February 2026"])

s = new_content_slide(prs, "Summary")
textbox(s, 0.58, 1.3, 12.17, 0.6, ["One or two plain sentences about what happened."], size=12)
table(s, 0.58, 2.1, 7.6, [2.6, 3.0, 2.0], [
    ["Channel", "Result", "Note"],
    ["Email", "37,112 recipients", "16% open rate"],
    ["Total reach", "38,788", ""],
])
panel(s, 8.6, 2.1, 4.15, 2.9); label(s, 8.75, 2.22, 3.85, "What stood out")
bullets(s, 8.75, 2.6, 3.85, 2.3, ["Fact one.", "Fact two."], size=10.5)
bar_chart(s, 0.58, 4.6, 7.6, 1.9, ["Saudi Arabia", "UAE"], [0.20, 0.15])

prs.save("my_report.pptx")
export_png("my_report.pptx", "preview")                 # optional: slide01.png, slide02.png …
```

Helpers: `load_template`, `new_cover`, `new_content_slide`, `textbox`, `label`, `bullets`,
`panel`, `table`, `bar_chart`, `add_picture`, `export_png`, `fmt`. Positions and sizes are in
inches. Runs inside a paragraph can be `("text", {"bold": True, "color": DARK_BLUE})`.

The complete worked example is `reports/2026-07-data-centre-trade-media/build_report.py`.

---

## 10. What is in this folder

```
brand/
  README.md                      this page
  assets/
    grundfos-logo-endline.svg/png  full logo with endline (cover)
    grundfos-mark.svg/png          "X" mark (content slides)
    cover-water.jpg                cover photograph
    theme-colours.json             theme colours, roles and font names
  templates/
    Grundfos_2022_Report_Template.pptx   corporate master + layouts, two sample slides
  tools/
    grundfos_pptx.py               python-pptx helpers used by every report
```

The template was derived from the team's own "NAS Marketing – Q1 Plans FY2026" deck: all plan
content removed, master, layouts and theme kept as they are.

---

## Sources

- Grundfos, "Grundfos launches new brand identity" (news, 30 Sept 2022) – endline and promise.
- Grundfos press materials page – logo usage and clear-space rule.
- Grundfos "About us" / annual report – company facts (check the current figures before quoting
  them externally).
- Team decks (Q1 Plans FY2026, Proposed Plans FY2026, Monthly Marketing Review) – theme, fonts,
  layouts and house style.
