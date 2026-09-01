import { useEffect } from "react";

const TITLE_BRAND = "EDITOR GRATUITO";

/**
 * Client-side document.title override for i18n.
 * SSG metadata stays in English for SEO; this updates the title
 * for users who switch languages on the client.
 */
export function usePageTitle(title: string) {
  useEffect(() => {
    const next =
      !title || title.includes(TITLE_BRAND)
        ? title
        : `${title} | ${TITLE_BRAND}`;
    document.title = next;
  }, [title]);
}
