import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Me Mother Earth",
  robots: { index: false, follow: false },
};

// Minimal homepage. The product landing pages are standalone routes and are
// intentionally not linked from here (no header/footer navigation to them).
export default function Home() {
  return (
    <main
      style={{
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        fontFamily: "'Montserrat', system-ui, sans-serif",
        color: "#5E704C",
        background: "#EDEDED",
      }}
    >
      <h1
        style={{
          fontFamily: "'Abril Fatface', Georgia, serif",
          fontWeight: 400,
          fontSize: "clamp(1.5rem, 5vw, 2.5rem)",
        }}
      >
        Me Mother Earth
      </h1>
    </main>
  );
}
