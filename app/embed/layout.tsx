import type { Metadata } from "next";
import { PropsWithChildren } from "react";

export const metadata: Metadata = {
  title: "Thin Embed",
  description: "Thin embed AGPL v3 document editor endpoint for web applications.",
  robots: {
    index: false,
    follow: false,
  },
};

export default function EmbedLayout({ children }: PropsWithChildren<unknown>) {
  return (
    <div className="w-screen h-screen overflow-hidden bg-background m-0 p-0">
      {children}
    </div>
  );
}
