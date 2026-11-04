/*
 * orbit-funnel.js: drppc.ai → Full Circle Orbit sign-up (one product, one account).
 *
 * 1. Meta pixel "Dr. PPC + Orbit" (1054966687594188, Full Circle IND LLC). PageView on every
 *    page. The same pixel runs on fullcircleorbit.com, so Meta sees one funnel across both
 *    domains. Skipped when the browser sends Global Privacy Control.
 * 2. Every "Start for $1" button opens fullcircleorbit.com/signup with from=drppc, plus the ad
 *    details the visitor landed with (utm_*, fbclid, gclid), so Orbit can credit the ad.
 *    A creator referral (referral.js) keeps its utm_source=creator.
 * 3. A click on it fires InitiateCheckout (Meta) and start_trial_click (GA4).
 */
(function () {
  'use strict';
  if (window.__orbitFunnel) return;
  window.__orbitFunnel = 1;

  var PIXEL = '1054966687594188';
  var SIGNUP = 'https://fullcircleorbit.com/signup';
  var KEY = 'drppc-landing-params';
  var KEEP = ['utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'utm_term', 'fbclid', 'gclid'];
  var gpc = !!navigator.globalPrivacyControl;

  if (!gpc && !window.fbq) {
    !function (f, b, e, v, n, t, s) {
      if (f.fbq) return; n = f.fbq = function () { n.callMethod ? n.callMethod.apply(n, arguments) : n.queue.push(arguments); };
      if (!f._fbq) f._fbq = n; n.push = n; n.loaded = !0; n.version = '2.0'; n.queue = [];
      t = b.createElement(e); t.async = !0; t.src = v; s = b.getElementsByTagName(e)[0]; s.parentNode.insertBefore(t, s);
    }(window, document, 'script', 'https://connect.facebook.net/en_US/fbevents.js');
    window.fbq('init', PIXEL);
    window.fbq('track', 'PageView');
  }

  // Ad details from the landing URL; a new ad click replaces the old set (last click wins).
  var saved = {};
  try { saved = JSON.parse(sessionStorage.getItem(KEY) || '{}') || {}; } catch (e) { saved = {}; }
  try {
    var q = new URLSearchParams(location.search), fresh = {}, any = false;
    KEEP.forEach(function (k) { var v = q.get(k); if (v) { fresh[k] = v.slice(0, 300); any = true; } });
    if (any) { saved = fresh; sessionStorage.setItem(KEY, JSON.stringify(saved)); }
  } catch (e) { /* storage blocked: links still work */ }

  function hasCreatorReferral() {
    try {
      var v = JSON.parse(localStorage.getItem('fc-creator-referral') || 'null');
      return !!(v && Date.now() - v.at < 60 * 24 * 60 * 60 * 1000);
    } catch (e) { return false; }
  }
  function isStart(a) {
    if (a.hasAttribute('data-start')) return true;
    if (/\bstart\b[^$]{0,20}\$1\b/i.test(a.textContent || '')) return true;
    return /^https:\/\/fullcircleorbit\.com\/signup/.test(a.href || '');
  }
  function isOrbit(url) { return /(^|\.)fullcircleorbit\.com$/.test(url.hostname); }
  function carry(url) {
    Object.keys(saved).forEach(function (k) { if (!url.searchParams.has(k)) url.searchParams.set(k, saved[k]); });
    if (!url.searchParams.has('utm_source') && !hasCreatorReferral()) {
      url.searchParams.set('utm_source', 'drppc.ai');
      url.searchParams.set('utm_medium', 'site');
    }
    return url;
  }

  function upgrade() {
    var links = document.querySelectorAll('a[href]');
    for (var i = 0; i < links.length; i++) {
      var a = links[i];
      if (!isStart(a)) continue;
      var url = new URL(SIGNUP);
      url.searchParams.set('from', 'drppc');
      a.href = carry(url).toString();
    }
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', upgrade); else upgrade();

  // Capture phase, so the link is final before the browser follows it.
  document.addEventListener('click', function (e) {
    var a = e.target && e.target.closest ? e.target.closest('a[href]') : null;
    if (!a) return;
    var url;
    try { url = new URL(a.href); } catch (err) { return; }
    if (!isOrbit(url)) return;
    a.href = carry(url).toString();
    if (isStart(a)) {
      if (window.fbq) window.fbq('track', 'InitiateCheckout', { content_name: 'Dr. PPC first 30 days', value: 1, currency: 'USD' });
      if (window.gtag) window.gtag('event', 'start_trial_click', { link_url: url.origin + url.pathname });
    }
  }, true);
})();
