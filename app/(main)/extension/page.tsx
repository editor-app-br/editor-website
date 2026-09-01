import type { Metadata } from "next";
import Image from "next/image";
import Link from "next/link";

export const metadata: Metadata = {
  title: "Browser extension",
  description:
    "Install the editor.app.br Chrome / Edge extension to open local DOCX, XLSX, PPTX and PDF files without uploading them.",
};

export default function ExtensionPage() {
  return (
    <main className="mx-auto max-w-2xl px-6 py-12">
      <div className="flex items-center gap-3 mb-8">
        <Image
          src="/editor_cube_transparent.png"
          alt=""
          width={48}
          height={48}
        />
        <div>
          <h1 className="text-2xl font-semibold tracking-tight text-foreground">
            EDITOR GRATUITO
          </h1>
          <p className="text-sm text-muted-foreground">
            Browser extension — open local Office files in the browser
          </p>
        </div>
      </div>

      <p className="text-sm text-muted-foreground leading-relaxed mb-8">
        The extension lets the site read <code className="text-foreground">file://</code>{" "}
        paths and cross-origin downloads without uploading documents to a server.
        Source lives next to this website as{" "}
        <code className="text-foreground">editor-google-app</code> (AGPL-3.0).
      </p>

      <section className="space-y-4 mb-10">
        <h2 className="text-base font-semibold text-foreground">
          Install (unpacked)
        </h2>
        <ol className="list-decimal list-inside space-y-2 text-sm text-muted-foreground leading-relaxed">
          <li>
            Clone or copy the{" "}
            <code className="text-foreground">editor-google-app</code> folder
            from the editor.app.br project.
          </li>
          <li>
            Open{" "}
            <code className="text-foreground">chrome://extensions</code> (or{" "}
            <code className="text-foreground">edge://extensions</code>).
          </li>
          <li>Enable Developer mode → Load unpacked → select that folder.</li>
          <li>
            Optional: enable{" "}
            <strong className="text-foreground">Allow access to file URLs</strong>{" "}
            so drag-and-drop of local files works.
          </li>
          <li>
            Return here and{" "}
            <strong className="text-foreground">reload</strong> the tab.
          </li>
        </ol>
      </section>

      <section className="space-y-3 mb-10">
        <h2 className="text-base font-semibold text-foreground">Chrome Web Store</h2>
        <p className="text-sm text-muted-foreground leading-relaxed">
          A public store listing is not published yet. When it is, this page will
          link to it. Until then, use Load unpacked from the source tree.
        </p>
      </section>

      <p className="text-sm">
        <Link
          href="/"
          className="text-primary font-medium hover:underline underline-offset-4"
        >
          ← Back to EDITOR GRATUITO
        </Link>
      </p>
    </main>
  );
}
