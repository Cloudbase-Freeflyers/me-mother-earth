import type { Section } from "./SiteHeader";

const LOGO =
  "https://memotherearthbrand.com/cdn/shop/files/oie_bjcQ6vDjHNyg_150x.png?v=1652469127";

/**
 * Responsive site footer: brand mark plus links that smooth-scroll to each
 * section (native smooth scroll via `scroll-behavior` in globals.css).
 */
export default function SiteFooter({ sections = [] }: { sections?: Section[] }) {
  const year = new Date().getFullYear();

  return (
    <footer className="site-footer">
      <div className="site-footer__inner">
        <div className="site-footer__brand">
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img src={LOGO} alt="Me Mother Earth" />
          <p>Sustainable, plastic-free home essentials — plant-derived and mineral-based.</p>
        </div>

        {sections.length > 0 && (
          <nav className="site-footer__nav" aria-label="Page sections">
            {sections.map((s) => (
              <a key={s.id} href={`#${s.id}`}>
                {s.label}
              </a>
            ))}
          </nav>
        )}
      </div>

      <div className="site-footer__bar">
        © {year} Me Mother Earth. All rights reserved.
      </div>
    </footer>
  );
}
