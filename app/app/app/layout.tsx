import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "The Morning Tape — A Pre-Market Brief",
  description: "A personalized pre-market briefing delivered before the bell. Built for traders who don't have time to read 40 tabs.",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
