// GET /geo -> {"country":"ZA"}. Only this route invokes a Function (auto-generated _routes.json).
// _headers does not apply to Function responses, so headers are set here.
export function onRequestGet({ request }) {
  const country = (request.cf && request.cf.country) || request.headers.get("CF-IPCountry") || "XX";
  return new Response(JSON.stringify({ country }), {
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "private, no-store",
      "X-Robots-Tag": "noindex",
    },
  });
}
