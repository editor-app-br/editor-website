// Kill-switch for the old site-wide SW. Keep editor-static-* and OnlyOffice
// document_editor_static_* caches used by /embed?warmup=1.
self.addEventListener("install", () => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {
      const keys = await caches.keys();
      await Promise.all(
        keys
          .filter(
            (key) =>
              !key.startsWith("editor-static-") &&
              !key.startsWith("document_editor_static_"),
          )
          .map((key) => caches.delete(key)),
      );
      await self.registration.unregister();
    })(),
  );
});
