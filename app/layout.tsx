import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Me Mother Earth",
  description:
    "Me Mother Earth — sustainable, plastic-free home essentials with plant-derived and mineral-based formulas.",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <head>
        {/*
          The PageFly landing pages reference "Abril Fatface" and "Montserrat"
          by their literal family names in inline styles, so they are loaded
          globally here rather than via next/font (which generates hashed names).
        */}
        <link rel="preconnect" href="https://fonts.googleapis.com" />
        <link
          rel="preconnect"
          href="https://fonts.gstatic.com"
          crossOrigin="anonymous"
        />
        <link
          href="https://fonts.googleapis.com/css2?family=Abril+Fatface&family=Montserrat:wght@400;500;600;700;800;900&display=swap"
          rel="stylesheet"
        />
      </head>
      <body>{children}</body>
    </html>
  );
}
