"use client";

import { useState } from "react";

export type Section = { id: string; label: string };

const LOGO =
  "https://memotherearthbrand.com/cdn/shop/files/oie_bjcQ6vDjHNyg_150x.png?v=1652469127";

/**
 * Sticky, responsive site header: MeMotherEarth logo on the left and section
 * links on the right that smooth-scroll to the matching section. Collapses to a
 * hamburger menu on small screens.
 */
export default function SiteHeader({ sections = [] }: { sections?: Section[] }) {
  const [open, setOpen] = useState(false);

  const toTop = (e: React.MouseEvent) => {
    e.preventDefault();
    setOpen(false);
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  return (
    <header className="site-header">
      <div className="site-header__inner">
        <a
          href="#"
          className="site-header__logo"
          aria-label="Me Mother Earth — back to top"
          onClick={toTop}
        >
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img src={LOGO} alt="Me Mother Earth" />
        </a>

        {sections.length > 0 && (
          <>
            <button
              type="button"
              className="site-header__toggle"
              aria-label="Toggle navigation menu"
              aria-expanded={open}
              onClick={() => setOpen((o) => !o)}
            >
              <span />
              <span />
              <span />
            </button>

            <nav className={`site-nav${open ? " site-nav--open" : ""}`}>
              {sections.map((s) => (
                <a key={s.id} href={`#${s.id}`} onClick={() => setOpen(false)}>
                  {s.label}
                </a>
              ))}
            </nav>
          </>
        )}
      </div>
    </header>
  );
}
