# CBS Project Stories — customer-facing news site

A simple news website that shows our Commercial Building Services (CBS) projects for 2025/2026.
Send the link to customers, consultants and contractors to build trust with proven references.

- `index.html`: home page with headline stats, a featured story, search, and filters (industry, application, country, year)
- `story.html?id=<project-id>`: one page per project with key numbers, challenge → solution → results, a customer quote, a gallery, and share buttons (WhatsApp, LinkedIn, Email, Copy link)
- `data/projects.js`: **the only file you edit** to add or update projects
- `images/`: project photos

There's no build step and nothing to install. Open `index.html` in a browser to preview it.

## Add a new project (about 10 minutes)

1. Open `data/projects.js` and copy one `{ ... }` block to the top of the list.
2. Set a unique `id`, for example `riyadh-data-center-2026`. Your share link is `story.html?id=riyadh-data-center-2026`.
3. Fill in the title, customer, date (`YYYY-MM`), city, country, industry, application, summary, challenge, solution, metrics, results, products, and quote.
4. Add photos to `images/` and set `"image": "images/your-photo.jpg"` (landscape, about 1600px wide, under 400 KB).
5. Set `featured: true` on the one story you want at the top of the home page.
6. Set `published: true` **only after the customer has approved external publication** (written approval by email is enough). Drafts with `published: false` stay hidden.

Before you share the link, delete the 5 `sample: true` blocks.

## Publish it (pick one)

| Option | Steps | Best for |
|---|---|---|
| **GitHub Pages** | Repo → Settings → Pages → Source: `main` branch, `/ (root)`. The site then lives at `https://<org>.github.io/IMAP/projects/` | The fastest free option |
| **Company website / CMS** | Give the `projects/` folder to the web team, or rebuild the stories as case-study pages on the corporate site | The most trusted domain for customers. Check brand guidelines first |
| **SharePoint / internal hosting** | Upload the folder to a site that allows HTML | Sharing internally before you go external |

Ask Marketing/Brand to confirm logo use, colours, and the domain before you send the site to customers.

## Story checklist (what builds trust)

- [ ] Customer name and logo approved for external use
- [ ] At least 2 hard numbers (energy saved %, capacity, pumps, floors, uptime)
- [ ] One real customer quote with name and title
- [ ] 3 or more real site photos (no stock images)
- [ ] Clear challenge → solution → result
- [ ] Contact email set in `SITE.contactEmail`
- [ ] Optional: an Arabic summary for Saudi customers

## Driving traffic

- Add UTM tags when you share the link, for example
  `story.html?id=xxx&utm_source=aoc_emailer&utm_medium=email&utm_campaign=cbs_references_2026`
- Put the link in AOC emailers, sales email signatures, LinkedIn posts, WhatsApp follow-ups after meetings, and QR codes on event roll-ups.
- Filtered links work too, for example `index.html?industry=Data%20Center` sends a consultant only the data-centre references.
