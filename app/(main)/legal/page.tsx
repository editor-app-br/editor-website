import type { Metadata } from "next";
import { LegalView } from "@/components/main/legal-view";

export const metadata: Metadata = {
  title: "Legal notice",
  description:
    "Legal notice, AGPL v3 compliance, and upstream attributions for editor.app.br.",
  robots: { index: true, follow: true },
};

export default function LegalPage() {
  return <LegalView />;
}
