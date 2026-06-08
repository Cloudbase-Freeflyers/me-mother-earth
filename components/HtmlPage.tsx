"use client";

import { useEffect, useRef } from "react";

export type ScriptDef = {
  src: string | null;
  code: string | null;
  async: boolean;
};

/**
 * Renders a self-contained block of landing-page HTML (originally authored in
 * Shopify's PageFly builder) and runs any scripts it shipped with.
 *
 * The HTML is injected verbatim so the original inline styling is preserved
 * exactly. Inline <script> tags cannot execute when set via innerHTML, so they
 * are extracted at build time and re-attached here as real <script> elements.
 * Several original scripts register their initializers on `DOMContentLoaded`,
 * which has already fired by the time we run, so we re-dispatch it afterwards.
 */
export default function HtmlPage({
  html,
  scripts = [],
}: {
  html: string;
  scripts?: ScriptDef[];
}) {
  const ran = useRef(false);

  useEffect(() => {
    if (ran.current || scripts.length === 0) return;
    ran.current = true;

    for (const def of scripts) {
      const el = document.createElement("script");
      if (def.src) {
        el.src = def.src;
        if (def.async) el.async = true;
      } else if (def.code) {
        el.text = def.code;
      } else {
        continue;
      }
      el.setAttribute("data-landing-script", "");
      document.body.appendChild(el);
    }

    // Re-fire so handlers registered on DOMContentLoaded still initialise.
    document.dispatchEvent(new Event("DOMContentLoaded"));
  }, [scripts]);

  return <div dangerouslySetInnerHTML={{ __html: html }} />;
}
