// One-time build tool: converts the unique PageFly landing-page HTML files in
// ../landing-pages into self-contained Next.js routes under ../app/<slug>.
//
// For each page it writes:
//   app/<slug>/content.json  -> { html, scripts, sections }
//   app/<slug>/page.tsx      -> a route rendering <SiteHeader> + <HtmlPage> + <SiteFooter>
//
// "sections" is the curated list of anchor links shown in the header/footer.
// Anchor ids (and a .mme-scroll-target class) are injected into the page markup
// so the links smooth-scroll to each section.
//
// After running this once, the generated routes no longer reference the
// landing-pages folder, so that folder can be deleted. Re-run with:
//   npm run generate:pages
//
// NOTE: This script reads landing-pages/, so it only works while that folder
// still exists. The generated output does not depend on it.

import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { parse } from "node-html-parser";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, "..");
const SRC_DIR = join(ROOT, "landing-pages");
const APP_DIR = join(ROOT, "app");

// Deduplicated mapping of source file -> route slug.
// Skipped duplicates:
//   - Bamboo-Toothbrush-Heads-LP.html       (identical to bamboo-toothbrush-heads.html except a BOM)
//   - Bamboo-Toothbrush-Heads-LP.print.html (redundant print variant)
//   - Laundry.html                          (same as new-laundry.html, older hero image)
const PAGES = [
  { src: "bamboo-toothbrush-heads.html", slug: "bamboo-toothbrush-heads" },
  { src: "new-laundry.html", slug: "laundry-sheets" },
  { src: "laundry-scent-boosters.html", slug: "laundry-scent-boosters" },
  { src: "mouthwash-tablets.html", slug: "mouthwash-tablets" },
  { src: "toothpaste-tablets.html", slug: "toothpaste-tablets" },
  { src: "stainless-steel-tongue-scraper-amazon.html", slug: "tongue-scraper" },
  { src: "toilet-bowl-cleaning-sheets.html", slug: "toilet-bowl-cleaning-sheets" },
  {
    src: "quick-dry-diatomaceous-earth-floor-stone-bath-mat-and-dish-mat.html",
    slug: "bath-mat",
  },
  { src: "quiz.html", slug: "quiz" },
];

const PARSE_OPTS = {
  comment: true,
  blockTextElements: { script: true, style: true, pre: true, textarea: true },
};

const isFullDoc = (s) => /<!doctype/i.test(s) || /<html[\s>]/i.test(s);

function extractScripts(src) {
  const scripts = [];
  const re = /<script\b([^>]*)>([\s\S]*?)<\/script>/gi;
  let m;
  while ((m = re.exec(src)) !== null) {
    const attrs = m[1];
    const code = m[2];
    const srcMatch = attrs.match(/\bsrc\s*=\s*["']([^"']+)["']/i);
    scripts.push({
      src: srcMatch ? srcMatch[1] : null,
      code: srcMatch ? null : code,
      async: /\basync\b/i.test(attrs),
    });
  }
  return scripts;
}

// Anchors pointing at the Shopify storefront. Tagged with data-gtm-cta so a GTM
// Link Click trigger can target a[data-gtm-cta] instead of guessing by URL.
const SHOPIFY_HOST = /https?:\/\/(?:www\.)?memotherearthbrand\.com/i;

// Clean up the extracted scripts:
//  - drop the hard-coded Google Ads tag (now handled centrally via GTM),
//  - tag dynamically-generated Shopify CTAs (e.g. the quiz result links).
function processScripts(scripts) {
  return scripts
    .filter((s) => !(s.src && /googletagmanager\.com\/gtag\/js\?id=AW-/i.test(s.src)))
    .filter((s) => !(s.code && /gtag\(\s*['"]config['"]/.test(s.code)))
    .map((s) => {
      if (!s.code) return s;
      const code = s.code
        // Remove the Google Ads conversion function...
        .replace(/function\s+firetag\s*\(\)\s*\{\s*gtag\([\s\S]*?\}\);\s*\}/g, "")
        // ...and retarget the size-selector that found its buttons via it.
        .replace(/a\[onclick=["']firetag\(\)["']\]/g, 'a[data-gtm-cta="shopify"]')
        // Tag Shopify CTAs built inside template strings.
        .replace(
          /<a\b(?=[^>]*href=["']https?:\/\/(?:www\.)?memotherearthbrand\.com)(?![^>]*data-gtm-cta)([^>]*?)>/gi,
          '<a data-gtm-cta="shopify"$1>'
        );
      return { ...s, code };
    });
}

function extractStyles(src) {
  return (src.match(/<style\b[^>]*>[\s\S]*?<\/style>/gi) || []).join("\n");
}

function buildHtml(src) {
  let html;
  if (isFullDoc(src)) {
    const bodyMatch = src.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    const body = bodyMatch ? bodyMatch[1] : src;
    const headMatch = src.match(/<head[^>]*>([\s\S]*?)<\/head>/i);
    const headStyles = headMatch ? extractStyles(headMatch[1]) : "";
    html = `${headStyles}\n${body}`;
  } else {
    html = src;
  }
  return html
    .replace(/^﻿/, "")
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/^\s*<!--[\s\S]*?-->\s*/, "")
    .trim();
}

const decode = (s) =>
  s
    .replace(/&mdash;/g, "—")
    .replace(/&ndash;/g, "–")
    .replace(/&amp;/g, "&")
    .replace(/&nbsp;/g, " ")
    .replace(/&#39;|&rsquo;|&apos;/g, "'")
    .replace(/&quot;/g, '"');

const clean = (s) => decode(s).replace(/\s+/g, " ").trim();

const stripTags = (s) => clean(s.replace(/<[^>]+>/g, " "));

function firstMatch(src, re) {
  const m = src.match(re);
  return m ? m[1] : "";
}

// ------------------------------------------------------------------ //
// Section detection + curated navigation                              //
// ------------------------------------------------------------------ //

// Display order for the nav. Only categories in this list appear as links.
const ORDER = [
  "Overview",
  "Why It Matters",
  "How It Works",
  "Take the Quiz",
  "Ingredients",
  "Reviews",
  "FAQ",
  "Buy Now",
];
const SHOWN = new Set(ORDER);

function categorize(heading, index) {
  // The hero (first block) is never a nav target — the logo scrolls to top, and
  // hero headlines often contain stray keywords (e.g. "ditch the", "step onto").
  if (index === 0) return null;

  const t = heading.toLowerCase();
  if (/why it matters/.test(t)) return "Why It Matters";
  if (/find your|perfect match|take the quiz/.test(t)) return "Take the Quiz";
  if (/\bsteps?\b|how it works|how to use/.test(t)) return "How It Works";
  if (/question|faq|everything you want to know/.test(t)) return "FAQ";
  if (/review|customer|testimonial|what (people|customers) say/.test(t))
    return "Reviews";
  if (
    /ingredient|what'?s inside|materials|what it'?s made|inside the|transparen|formula/.test(
      t
    )
  )
    return "Ingredients";
  if (
    /make the switch|ready (to|for)|add to cart|\bshop\b|\bbuy\b|order now|ditch the|make a difference|get yours/.test(
      t
    )
  )
    return "Buy Now";
  // First content block after the hero, when nothing more specific matched.
  if (index === 1) return "Overview";
  return null;
}

const elementChildren = (node) =>
  node.childNodes.filter(
    (n) =>
      n.nodeType === 1 &&
      !["style", "script"].includes((n.rawTagName || "").toLowerCase())
  );

// Find the level whose direct children are the page's sections, then return the
// ones that contain a heading.
function findSectionBlocks(root) {
  let candidates = elementChildren(root);
  while (true) {
    const holders = candidates.filter((e) => e.querySelector("h1, h2"));
    if (holders.length === 1 && elementChildren(holders[0]).length > 0) {
      candidates = elementChildren(holders[0]);
    } else {
      break;
    }
  }
  return candidates.filter((e) => e.querySelector("h1, h2"));
}

function processSections(html) {
  const root = parse(html, PARSE_OPTS);

  // Tag every static CTA that links to the Shopify storefront, and drop the
  // old Google Ads click hook (conversions now run through GTM).
  root.querySelectorAll("a").forEach((a) => {
    const href = a.getAttribute("href") || "";
    if (SHOPIFY_HOST.test(href)) a.setAttribute("data-gtm-cta", "shopify");
    if ((a.getAttribute("onclick") || "").replace(/\s+/g, "") === "firetag()") {
      a.removeAttribute("onclick");
    }
  });

  const blocks = findSectionBlocks(root);

  const sections = [];
  const seen = new Set();

  blocks.forEach((el, i) => {
    let id = el.getAttribute("id");
    if (!id) {
      id = `mme-sec-${i}`;
      el.setAttribute("id", id);
    }
    el.classList.add("mme-scroll-target");

    const headingEl = el.querySelector("h1, h2");
    const heading = clean(headingEl ? headingEl.text : "");
    const cat = categorize(heading, i);
    if (cat && SHOWN.has(cat) && !seen.has(cat)) {
      seen.add(cat);
      sections.push({ id, label: cat, order: ORDER.indexOf(cat) });
    }
  });

  sections.sort((a, b) => a.order - b.order);
  return {
    html: root.toString(),
    sections: sections.map(({ id, label }) => ({ id, label })),
  };
}

function deriveMeta(src) {
  let title = "";
  let description = "";

  if (isFullDoc(src)) {
    title = stripTags(firstMatch(src, /<title[^>]*>([\s\S]*?)<\/title>/i));
    description = decode(
      firstMatch(
        src,
        /<meta\s+name=["']description["']\s+content=["']([^"']*)["']/i
      )
    );
  }

  if (!title) {
    const h1 = stripTags(firstMatch(src, /<h1[^>]*>([\s\S]*?)<\/h1>/i));
    title = h1 ? `${h1} — Me Mother Earth` : "Me Mother Earth";
  }

  if (!description) {
    const p = stripTags(firstMatch(src, /<p[^>]*>([\s\S]*?)<\/p>/i));
    description = p.slice(0, 200);
  }

  return {
    title: title || "Me Mother Earth",
    description:
      description ||
      "Sustainable, plastic-free home essentials from Me Mother Earth.",
  };
}

const pageTemplate = (meta) => `import type { Metadata } from "next";
import SiteHeader from "@/components/SiteHeader";
import SiteFooter from "@/components/SiteFooter";
import HtmlPage from "@/components/HtmlPage";
import content from "./content.json";

export const metadata: Metadata = {
  title: ${JSON.stringify(meta.title)},
  description: ${JSON.stringify(meta.description)},
};

export default function Page() {
  return (
    <>
      <SiteHeader sections={content.sections} />
      <main>
        <HtmlPage html={content.html} scripts={content.scripts} />
      </main>
      <SiteFooter sections={content.sections} />
    </>
  );
}
`;

for (const { src, slug } of PAGES) {
  const raw = readFileSync(join(SRC_DIR, src), "utf8");
  const scripts = processScripts(extractScripts(raw));
  const meta = deriveMeta(raw);
  const { html, sections } = processSections(buildHtml(raw));

  const outDir = join(APP_DIR, slug);
  mkdirSync(outDir, { recursive: true });
  writeFileSync(
    join(outDir, "content.json"),
    JSON.stringify({ html, scripts, sections }, null, 2)
  );
  writeFileSync(join(outDir, "page.tsx"), pageTemplate(meta));

  const navLabels = sections.map((s) => s.label).join(" · ") || "(none)";
  console.log(`  /${slug.padEnd(28)} nav: ${navLabels}`);
}

console.log(`\nGenerated ${PAGES.length} routes into app/.`);
