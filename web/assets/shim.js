/* Thay thế google.script.run của Apps Script: gọi máy chủ Supabase (Edge Function chamcong-api).
   Giữ nguyên cú pháp cũ: google.script.run.withSuccessHandler(f).withFailureHandler(g).tenHam(...args) */
(function () {
  var API = 'https://jnzeawdffiwnblrlwtgx.supabase.co/functions/v1/chamcong-api';
  function make(ok, fail) {
    return new Proxy({}, {
      get: function (_, name) {
        if (name === 'withSuccessHandler') return function (f) { return make(f, fail); };
        if (name === 'withFailureHandler') return function (f) { return make(ok, f); };
        return function () {
          var args = Array.prototype.slice.call(arguments);
          fetch(API, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ fn: String(name), args: args }) })
            .then(function (r) { return r.json(); })
            .then(function (j) { if (ok) ok(j); })
            .catch(function (e) { if (fail) fail(e); else console.error(e); });
        };
      }
    });
  }
  window.google = { script: { run: make(null, null) } };
})();
