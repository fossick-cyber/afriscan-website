// Host redirects (architecture.FINAL §7.2). The deploy token cannot create Cloudflare Bulk Redirects or
// Redirect Rules, so www and the production pages.dev host are sent to the apex here, keeping path and
// query. Preview hosts (<hash>. and <branch>.afriscan-website.pages.dev) are left alone for review.
// dist/_routes.json keeps /assets/* out of Functions, so those requests never count against the quota.
const APEX = "afri-scan.com";
const REDIRECT_HOSTS = new Set(["www.afri-scan.com", "afriscan-website.pages.dev"]);

export function onRequest({ request, next }) {
  const url = new URL(request.url);
  if (!REDIRECT_HOSTS.has(url.hostname)) return next();
  url.protocol = "https:";
  url.hostname = APEX;
  url.port = "";
  return new Response(null, {
    status: 301,
    headers: {
      Location: url.toString(),
      "Cache-Control": "public, max-age=86400",
      "Strict-Transport-Security": "max-age=31536000; includeSubDomains",
    },
  });
}
