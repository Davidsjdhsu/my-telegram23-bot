/* Telegram menu launches use signed HTTP requests, never sendData fallback. */
(function (root) {
  function createMiniAppSession({ getTg, apiBase, fetchFn = fetch }) {
    let pending = false;
    let retryRequest = null;
    async function request(path, body) {
      const initData = getTg()?.initData;
      if (!initData) throw new Error('telegram_auth_required');
      const response = await fetchFn(apiBase + path, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ...body, initData }), cache: 'no-store',
      });
      const result = await response.json();
      if (!response.ok || !result.ok) throw new Error(result.error || 'request_failed');
      return result;
    }
    return {
      loadProfile: () => request('/api/profile', {}),
      async send(action, payload = {}) {
        if (pending) return false;
        pending = true;
        const fingerprint = JSON.stringify({ action, payload });
        if (!retryRequest || retryRequest.fingerprint !== fingerprint) {
          retryRequest = { fingerprint, id: Date.now().toString(36) + '-' + Math.random().toString(36).slice(2) };
        }
        try {
          await request('/api/action', { action, payload, requestId: retryRequest.id });
          retryRequest = null;
          return true;
        } finally { pending = false; }
      },
    };
  }
  if (typeof module !== 'undefined' && module.exports) module.exports = createMiniAppSession;
  else root.createMiniAppSession = createMiniAppSession;
})(typeof window === 'undefined' ? {} : window);
