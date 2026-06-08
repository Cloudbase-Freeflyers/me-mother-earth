import type { Metadata } from "next";
import SiteHeader from "@/components/SiteHeader";
import SiteFooter from "@/components/SiteFooter";
import HtmlPage from "@/components/HtmlPage";
import content from "./content.json";

export const metadata: Metadata = {
  title: "A cleaner mouth in seconds — just a few gentle strokes each morning — Me Mother Earth",
  description: "Scrape, rinse, done. This surgical-grade stainless steel tongue scraper removes odor-causing bacteria and buildup more effectively than brushing alone — built to last a lifetime with zero plastic wast",
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
