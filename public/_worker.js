const EDITOR_PATH_PREFIX = "/v9.3.0.24-1/";
const EDITOR_UPSTREAM_ORIGIN = "https://office-editor.ziziyi.com";

const copyRequestHeaders = (request) => {
  const headers = new Headers();
  [
    "accept",
    "accept-language",
    "if-modified-since",
    "if-none-match",
    "range",
    "user-agent",
  ].forEach((name) => {
    const value = request.headers.get(name);
    if (value) headers.set(name, value);
  });
  return headers;
};

const proxyEditorAsset = async (request) => {
  const requestUrl = new URL(request.url);
  const upstreamUrl = new URL(`${requestUrl.pathname}${requestUrl.search}`, EDITOR_UPSTREAM_ORIGIN);
  const upstreamResponse = await fetch(
    new Request(upstreamUrl, {
      method: request.method,
      headers: copyRequestHeaders(request),
      redirect: "follow",
    })
  );

  const headers = new Headers(upstreamResponse.headers);
  headers.delete("set-cookie");
  headers.set("x-content-type-options", "nosniff");
  if (upstreamResponse.ok && !headers.has("cache-control")) {
    headers.set("cache-control", "public, max-age=31536000, immutable");
  }

  const contentType = headers.get("content-type")?.toLowerCase() ?? "";
  if (contentType.includes("text/html")) {
    const html = (await upstreamResponse.text()).replaceAll(
      EDITOR_UPSTREAM_ORIGIN,
      requestUrl.origin
    );
    headers.delete("content-length");
    return new Response(html, {
      status: upstreamResponse.status,
      statusText: upstreamResponse.statusText,
      headers,
    });
  }

  return new Response(upstreamResponse.body, {
    status: upstreamResponse.status,
    statusText: upstreamResponse.statusText,
    headers,
  });
};

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const localResponse = await env.ASSETS.fetch(request);

    if (
      localResponse.status !== 404 ||
      !url.pathname.startsWith(EDITOR_PATH_PREFIX) ||
      (request.method !== "GET" && request.method !== "HEAD")
    ) {
      return localResponse;
    }

    try {
      return await proxyEditorAsset(request);
    } catch (error) {
      console.error("ONLYOFFICE asset proxy failed", url.pathname, error);
      return new Response("Office editor asset is temporarily unavailable.", {
        status: 502,
        headers: { "content-type": "text/plain; charset=utf-8" },
      });
    }
  },
};
