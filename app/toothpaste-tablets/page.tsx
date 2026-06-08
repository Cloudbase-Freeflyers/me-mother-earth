import type { Metadata } from "next";
import SiteHeader from "@/components/SiteHeader";
import SiteFooter from "@/components/SiteFooter";
import HtmlPage from "@/components/HtmlPage";
import content from "./content.json";

export const metadata: Metadata = {
  title: "Fresh foam, zero tube — in a refillable jar — Me Mother Earth",
  description: "Chew, brush, breathe out. Foamy, minty, plastic-free toothpaste tablets in a refillable amber glass jar. Same clean you love — no tube, no waste. Need more? Get our 3-month refill pouch and pour into ",
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
