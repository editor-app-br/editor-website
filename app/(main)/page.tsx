import type { Metadata } from "next";
import { OpenView } from "@/components/main/open-view";
import { getRecommendedTemplates } from "@/utils/templates";

export const metadata: Metadata = {
  title: "Free Online Office Editor — Open & Edit Word, Excel, PowerPoint",
  description:
    "Open, view, and edit DOCX, XLSX, PPTX files directly in your browser for free. Nothing is sent to any server — 100% local, 100% secure. No upload, no login.",
  keywords: [
    "online office editor",
    "free Word editor online",
    "free Excel editor online",
    "free PowerPoint editor online",
    "open docx online",
    "edit xlsx in browser",
    "edit pptx in browser",
    "no upload document editor",
    "privacy first office",
    "office gratuito",
    "editor.app.br",
  ],
  alternates: {
    canonical: "https://editor.app.br",
  },
  openGraph: {
    title: "Free Online Office Editor — Word, Excel, PowerPoint | EDITOR GRATUITO",
    description:
      "Edit Office documents in your browser. Nothing is sent to any server — 100% local, 100% secure.",
    url: "https://editor.app.br",
    type: "website",
  },
  twitter: {
    card: "summary_large_image",
    title: "Free Online Office Editor — Word, Excel, PowerPoint | EDITOR GRATUITO",
    description:
      "Edit Office documents in your browser. Nothing is sent to any server — 100% local, 100% secure.",
  },
};

export default function HomePage() {
  const templates = getRecommendedTemplates();
  return <OpenView recommendedTemplates={templates} />;
}
