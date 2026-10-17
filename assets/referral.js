/*
 * referral.js — carries a Full Circle affiliate referral from drppc.ai into Orbit.
 *
 * Affiliate links (publish.fullcircleagency.com/creators/go/...) land here with
 * ?ref=<link token>&code=<CREATOR CODE>. Both are kept for 60 days (last click
 * wins) and added to every link into Orbit (fullcircleorbit.com) when it is
 * clicked, so the signup that follows credits the creator. Price is unchanged.
 */
(function () {
  'use strict';
  var KEY = 'fc-creator-referral';
  var MAX_AGE = 60 * 24 * 60 * 60 * 1000;
  var REF = /^[A-Za-z0-9_-]{4,64}$/;
  var CODE = /^[A-Z0-9]{3,24}$/;

  function read() {
    try {
      var v = JSON.parse(localStorage.getItem(KEY) || 'null');
      if (v && Date.now() - v.at < MAX_AGE) return v;
    } catch (e) { /* storage blocked */ }
    return null;
  }

  try {
    var q = new URLSearchParams(location.search);
    var ref = q.get('ref');
    var code = (q.get('code') || '').replace(/[^A-Za-z0-9]/g, '').toUpperCase();
    ref = ref && REF.test(ref) ? ref : null;
    code = CODE.test(code) ? code : null;
    if (ref || code) localStorage.setItem(KEY, JSON.stringify({ ref: ref, code: code, at: Date.now() }));
  } catch (e) { /* storage blocked: links still work, just uncredited */ }

  document.addEventListener('click', function (e) {
    var a = e.target && e.target.closest ? e.target.closest('a[href]') : null;
    if (!a) return;
    var r = read();
    if (!r) return;
    var url;
    try { url = new URL(a.href); } catch (err) { return; }
    if (!/(^|\.)fullcircleorbit\.com$/.test(url.hostname)) return;
    if (r.ref && !url.searchParams.has('ref')) url.searchParams.set('ref', r.ref);
    if (r.code && !url.searchParams.has('code')) url.searchParams.set('code', r.code);
    a.href = url.toString();
  }, true);
})();
