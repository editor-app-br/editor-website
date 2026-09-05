import type { Metadata } from "next";
import { IntegrateView } from "@/components/main/integrate-view";

export const metadata: Metadata = {
  title: "Integrar via iframe",
  description:
    "Incorpore o EDITOR GRATUITO (editor.app.br) via iframe + postMessage e o plugin Agente. Software livre GNU AGPL v3.",
  robots: { index: true, follow: true },
};

export default function IntegratePage() {
  return <IntegrateView />;
}
