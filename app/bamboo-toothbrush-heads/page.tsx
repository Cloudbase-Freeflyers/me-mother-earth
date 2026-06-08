import type { Metadata } from "next";
import SiteHeader from "@/components/SiteHeader";
import SiteFooter from "@/components/SiteFooter";
import HtmlPage from "@/components/HtmlPage";
import content from "./content.json";

export const metadata: Metadata = {
  title: "Keep your electric routine — ditch the plastic-heavy habit — Me Mother Earth",
  description: "Swap, click, brush. These bamboo electric toothbrush heads are designed to fit most Philips Sonicare click-on handles, so you can keep what works while reducing everyday oral care waste.",
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
