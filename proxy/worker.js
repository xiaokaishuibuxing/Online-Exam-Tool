// 可选：Cloudflare Worker 中转（仅当浏览器直连 api.deepseek.com 被跨域拦截时才需要）
// 用法：在 Cloudflare Workers 新建一个 Worker，把本文件内容粘进去并部署，
// 然后在页面"AI 接口设置"里，把"接口地址"改成你的 Worker 地址（形如 https://xxx.workers.dev）。
// 说明：API Key 仍由页面通过 Authorization 头传来，这里只做转发，不保存任何内容。
export default {
  async fetch(request) {
    const cors = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Headers': 'Authorization, Content-Type',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    };
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors });
    const url = new URL(request.url);
    const target = 'https://api.deepseek.com' + url.pathname + url.search;
    const resp = await fetch(target, {
      method: request.method,
      headers: {
        'Content-Type': 'application/json',
        'Authorization': request.headers.get('Authorization') || '',
      },
      body: request.method === 'POST' ? await request.text() : undefined,
    });
    const headers = new Headers(resp.headers);
    Object.entries(cors).forEach(([k, v]) => headers.set(k, v));
    return new Response(resp.body, { status: resp.status, headers });
  },
};
