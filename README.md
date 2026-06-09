# Me Mother Earth


Next.js app for [Me Mother Earth](https://memotherearth.com) product landing pages,
plus brand assets and tooling.

## Getting started

```bash
npm install
npm run dev      # http://localhost:3000
npm run build    # production build
npm run start    # serve the production build
```

## Landing pages

Each product landing page is its own route under `app/<slug>/`. The markup was
converted from the Shopify PageFly HTML exports and is **fully self-contained** —
the original design (inline styles) is preserved verbatim and rendered via the
shared [`components/HtmlPage.tsx`](components/HtmlPage.tsx) component. Pages with
interactive JavaScript (the quiz and the bath-mat) have their scripts extracted
and re-run on the client.

Every page gets a shared, responsive [`SiteHeader`](components/SiteHeader.tsx)
(MeMotherEarth logo on the left + section links, collapsing to a hamburger menu
on mobile) and [`SiteFooter`](components/SiteFooter.tsx). The section links
smooth-scroll to anchors that the generator injects into each page's section
blocks. The link set per page is curated (e.g. Overview · How It Works ·
Ingredients · FAQ · Buy Now) and built automatically from the page's headings.

| Route | Product |
|-------|---------|
| `/bamboo-toothbrush-heads` | Bamboo electric toothbrush heads |
| `/laundry-sheets` | Laundry detergent sheets |
| `/laundry-scent-boosters` | Laundry scent boosters |
| `/mouthwash-tablets` | Mouthwash tablets |
| `/toothpaste-tablets` | Toothpaste tablets |
| `/tongue-scraper` | Stainless steel tongue scraper |
| `/toilet-bowl-cleaning-sheets` | Toilet bowl cleaning sheets |
| `/bath-mat` | Diatomaceous bath & dish mat (interactive) |
| `/quiz` | Eco-product finder quiz (interactive) |

These are standalone landing pages and are intentionally **not** linked from any
header or footer navigation. Product images load from the live
`memotherearthbrand.com` CDN, so no local image assets are required.

### Regenerating from source HTML

The pages were generated from `landing-pages/*.html` by a one-time tool:

```bash
npm run generate:pages   # reads landing-pages/, writes app/<slug>/content.json + page.tsx
```

Once generated, the routes no longer reference the `landing-pages/` folder, so it
can be deleted. (The generator only works while that folder still exists.)
Duplicate exports were collapsed: `Bamboo-Toothbrush-Heads-LP.html` /
`.print.html` map to `bamboo-toothbrush-heads.html`, and `Laundry.html` to the
newer `new-laundry.html`.

## Other contents

- `brand-knowledge.md` — Brand identity, compliance guardrails, marketing angles
- `bulkedit/` — Google Apps Script bulk-edit tools (unrelated to the web app)

## Brand at a glance

- **Style:** Minimal, clean, simple
- **Primary palette:** Sage Green (#A1AB97), Off-White (#EDEDED), Cream (#ECDFC9)
- **Key rule:** Never use the word "Natural" in copy — use plant-derived / mineral-based language instead
