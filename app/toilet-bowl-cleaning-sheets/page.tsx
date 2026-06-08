import type { Metadata } from "next";
import SiteHeader from "@/components/SiteHeader";
import SiteFooter from "@/components/SiteFooter";
import HtmlPage from "@/components/HtmlPage";
import content from "./content.json";

export const metadata: Metadata = {
  title: "Clean in seconds — without the plastic bottle — Me Mother Earth",
  description: "Tear, toss, and scrub. These toilet bowl cleaning sheets offer a smarter alternative to traditional liquid cleaners — designed to help lift stains and reduce odors without bulky jugs, spills, or waste",
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
