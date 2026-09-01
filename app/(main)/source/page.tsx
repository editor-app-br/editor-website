import type { Metadata } from "next";
import { SourceView } from "@/components/main/source-view";

export const metadata: Metadata = {
  title: "Source code",
  description:
    "Corresponding source code for editor.app.br under GNU AGPL v3 Section 13.",
  robots: { index: true, follow: true },
};

export default function SourcePage() {
  return <SourceView />;
}
