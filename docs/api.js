/* Thin client for the Supabase REST API (PostgREST). No SDK needed. */
(function () {
  var cfg = window.CONFIG || {};
  var base = String(cfg.SUPABASE_URL || '').replace(/\/+$/, '') + '/rest/v1/requests';

  function isConfigured() {
    return !!(cfg.SUPABASE_URL && cfg.SUPABASE_ANON_KEY &&
      cfg.SUPABASE_URL.indexOf('YOUR-PROJECT') === -1 &&
      cfg.SUPABASE_ANON_KEY.indexOf('YOUR-ANON') === -1);
  }

  function call(method, url, body, prefer) {
    var headers = {
      'apikey': cfg.SUPABASE_ANON_KEY,
      'Authorization': 'Bearer ' + cfg.SUPABASE_ANON_KEY,
      'Content-Type': 'application/json'
    };
    if (prefer) headers['Prefer'] = prefer;
    return fetch(url, {
      method: method,
      headers: headers,
      body: body == null ? undefined : JSON.stringify(body)
    }).then(function (res) {
      if (!res.ok) {
        return res.text().then(function (t) {
          var msg = t;
          try { msg = JSON.parse(t).message || t; } catch (e) { /* keep raw text */ }
          throw new Error(msg || (res.status + ' ' + res.statusText));
        });
      }
      return res.status === 204 ? null : res.json();
    });
  }

  window.API = {
    isConfigured: isConfigured,

    /** Display reference for a request, e.g. REQ-0042. */
    refOf: function (id) { return 'REQ-' + String(id).padStart(4, '0'); },

    /** data: {request_name, requester, date_requested, applications, jira, sprint, priority, description} */
    insertRequest: function (data) {
      return call('POST', base, [data], 'return=representation')
        .then(function (rows) { return rows[0]; });
    },

    /** Newest first. */
    listRequests: function () {
      return call('GET', base + '?select=*&order=created_at.desc');
    },

    /** patch: any subset of {status, priority, jira, sprint} */
    updateRequest: function (id, patch) {
      return call('PATCH', base + '?id=eq.' + encodeURIComponent(id), patch, 'return=representation')
        .then(function (rows) {
          if (!rows || !rows.length) throw new Error('Request not found');
          return rows[0];
        });
    }
  };
})();
