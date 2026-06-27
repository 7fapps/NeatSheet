// service-worker.js
// A minimal service worker for NeatSheet.
//
// PWAs require a service worker to be considered "installable" by
// browsers (especially Android Chrome). This one is intentionally
// simple: it doesn't cache anything fancy, since NeatSheet needs a
// live connection to the Flask server to actually clean files anyway
// (it's not the kind of app that needs to work fully offline).

self.addEventListener("install", (event) => {
  // Nothing special to pre-cache - NeatSheet always needs a live
  // server connection to process files.
  console.log("NeatSheet service worker installed.");
});

self.addEventListener("fetch", (event) => {
  // Just let all requests pass through normally to the server.
  event.respondWith(fetch(event.request));
});
