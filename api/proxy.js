/**
 * CVP 本量利儀表板 — Vercel Function CORS 代理（Fluid Compute / Node.js）
 *
 * 支援 8 個 AI 供應商（OpenAI / Anthropic / DeepSeek / MiniMax /
 *                      硅基流動 / OpenRouter / Ollama / 自訂）
 *
 * 用法：
 *   GET/POST  https://your-app.vercel.app/api/proxy?url=<encoded-target-url>
 */

/* ── CORS 標頭 ── */
const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': [
    'Authorization',
    'Content-Type',
    'x-api-key',
    'anthropic-version',
    'anthropic-dangerous-direct-browser-access',
    'OpenAI-Organization', 'OpenAI-Beta',
    'HTTP-Referer', 'X-Title',
    'api-key',
    'X-Goog-Api-Key',
    'X-MiniMax-Group-Id',
    'X-Proxy-Target',
  ].join(', '),
  'Access-Control-Max-Age': '86400',
};

/* ── 不轉發給上游的 request headers ── */
const SKIP_REQ_HEADERS = new Set([
  'host', 'content-length', 'connection', 'transfer-encoding',
  'te', 'trailer', 'upgrade', 'origin', 'referer',
  'x-vercel-id', 'x-vercel-deployment-url', 'x-real-ip',
  'x-forwarded-for', 'x-forwarded-host', 'x-forwarded-proto',
  'x-vercel-forwarded-for', 'x-vercel-ip-country',
  'x-vercel-ip-country-region', 'x-vercel-ip-city',
  'x-vercel-ip-latitude', 'x-vercel-ip-longitude',
  'x-vercel-ip-timezone', 'x-vercel-internal-ingress-bucket',
  'cdn-loop', 'forwarded',
]);

/* ── 不轉發給 client 的 response headers ── */
const SKIP_RESP_HEADERS = new Set([
  'access-control-allow-origin',
  'access-control-allow-headers',
  'access-control-allow-methods',
  'access-control-max-age',
  'content-encoding',
  'content-length',
  'transfer-encoding',
  'connection',
]);

function sendJson(res, status, body) {
  const payload = JSON.stringify(body, null, 2);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    ...CORS_HEADERS,
  });
  res.end(payload);
}

/* ── 主 handler（Node.js IncomingMessage / ServerResponse） ── */
export default async function handler(req, res) {
  // CORS preflight
  if (req.method === 'OPTIONS') {
    res.writeHead(200, CORS_HEADERS);
    res.end();
    return;
  }

  // Node.js runtime 的 req.url 只有路徑，需補 host 才能用 URL API 解析
  const baseUrl = `https://${req.headers.host}`;
  const parsedUrl = new URL(req.url, baseUrl);
  // 支援兩種傳 target 方式：?url= query param 或 X-Proxy-Target header
  const target = parsedUrl.searchParams.get('url') || req.headers['x-proxy-target'];

  if (!target) {
    sendJson(res, 400, {
      error: '缺少目標 URL：請用 ?url= 參數或 X-Proxy-Target header 指定',
      hint: '用法 A（query）：/api/proxy?url=https%3A%2F%2Fapi.openai.com%2Fv1%2Fchat%2Fcompletions\n用法 B（header）：X-Proxy-Target: https://api.openai.com/v1/chat/completions',
    });
    return;
  }

  // 驗證目標 URL
  let targetUrl;
  try {
    targetUrl = new URL(target);
    if (!['http:', 'https:'].includes(targetUrl.protocol)) {
      sendJson(res, 400, { error: `不支援的協議：${targetUrl.protocol}` });
      return;
    }
  } catch {
    sendJson(res, 400, { error: `無效的目標 URL：${target.slice(0, 100)}` });
    return;
  }

  // 組裝轉發 headers
  const fwdHeaders = {};
  for (const [key, value] of Object.entries(req.headers)) {
    if (!SKIP_REQ_HEADERS.has(key.toLowerCase())) {
      fwdHeaders[key] = value;
    }
  }

  // 收集 request body（POST / PUT / PATCH）
  let body = undefined;
  if (req.method !== 'GET' && req.method !== 'HEAD') {
    const chunks = [];
    for await (const chunk of req) {
      chunks.push(chunk);
    }
    body = Buffer.concat(chunks);
  }

  // 轉發請求
  try {
    const upstream = await fetch(target, {
      method: req.method,
      headers: fwdHeaders,
      body,
      redirect: 'manual',
    });

    // 設定回應 status + headers
    const respHeaders = { ...CORS_HEADERS };
    upstream.headers.forEach((value, key) => {
      if (!SKIP_RESP_HEADERS.has(key.toLowerCase())) {
        respHeaders[key] = value;
      }
    });
    res.writeHead(upstream.status, respHeaders);

    // 串流回傳（SSE / chunked 直接穿透）
    if (upstream.body) {
      const reader = upstream.body.getReader();
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        res.write(value);
      }
    }
    res.end();
  } catch (e) {
    if (!res.headersSent) {
      sendJson(res, 502, {
        error: `代理錯誤：${e.message || e}`,
        target: target.slice(0, 200),
      });
    }
  }
}
