/*
 * funnel.js: Hero ASIN skill funnel (drppc.ai/skill/hero-asin → -ty → -guide).
 *
 * One place that ties the three pages to everything else:
 *   1. Meta pixel (1054966687594188) Lead, with an event id the server can match (no double count).
 *   2. GA4 (G-M4G8GYYYXE) events: generate_lead, skill_video_unmute, skill_guide_open, skill_audit_connect.
 *   3. The 6-channel remarketing engine (remarketing-engine/router.mjs):
 *        form submit            → POST /api/lead
 *        guide "connect" click  → POST /api/audit-complete   (stage: connect_clicked)
 *      /api/trial-started is sent by Orbit after the $1 signup, never by these pages.
 *   4. The `drppc_audit` state that assets/site-banner.js reads to greet a returning lead on drppc.ai.
 *
 * ENDPOINT is the public address of the remarketing engine. It is empty until the engine is hosted:
 * until then nothing is posted from the live site (the pixel, GA4 and the saved state still work).
 * On localhost it posts to the engine running on this computer (http://127.0.0.1:8788, dry run).
 */
(function () {
  'use strict';
  var ENDPOINT = '';            // e.g. 'https://remarketing.drppc.ai'  (no trailing slash)

  var LEAD_KEY = 'fc_hormozi_lead';
  var AUDIT_KEY = 'drppc_audit';
  var SPEND = { 'Under $5K a month': 2500, '$5K to $10K a month': 7500, '$10K to $50K a month': 25000, '$50K + a month': 75000 };
  var REVENUE = { 'Under $500K': 250000, '$500K to $1M': 750000, '$1M to $5M': 3000000, '$5M to $20M': 12500000, '$20 Million +': 20000000 };

  var isLocal = /^(localhost|127\.0\.0\.1)$/.test(location.hostname);
  var base = (window.FC_REMARKETING_ENDPOINT || ENDPOINT || (isLocal ? 'http://127.0.0.1:8788' : '')).replace(/\/+$/, '');

  function cookie(name) {
    var m = document.cookie.match(new RegExp('(?:^|; )' + name + '=([^;]*)'));
    return m ? decodeURIComponent(m[1]) : '';
  }
  function landing() {
    try { return JSON.parse(sessionStorage.getItem('drppc-landing-params') || '{}') || {}; } catch (e) { return {}; }
  }
  function storedLead() {
    try { return JSON.parse(sessionStorage.getItem(LEAD_KEY) || 'null'); } catch (e) { return null; }
  }
  function post(route, payload) {
    if (!base) return false;
    try {
      fetch(base + route, { method: 'POST', headers: { 'Content-Type': 'application/json' }, keepalive: true, body: JSON.stringify(payload) })
        .catch(function () { /* never blocks the page */ });
      return true;
    } catch (e) { return false; }
  }
  function ga(name, params) { try { if (window.gtag) window.gtag('event', name, params || {}); } catch (e) { /* blocked */ } }
  function fb(name, params, custom, eventId) {
    try {
      if (!window.fbq) return;
      window.fbq(custom ? 'trackCustom' : 'track', name, params || {}, eventId ? { eventID: eventId } : undefined);
    } catch (e) { /* blocked */ }
  }
  function saveAudit(extra) {
    var state = { ready: true, asin: '', brand: '', estimatedBleed: '', promoCode: 'FOUNDER10', trialStarted: false, source: 'hero_asin_skill', savedAt: new Date().toISOString() };
    for (var k in (extra || {})) state[k] = extra[k];
    var raw = JSON.stringify(state);
    try { localStorage.setItem(AUDIT_KEY, raw); } catch (e) { /* storage blocked */ }
    try { document.cookie = AUDIT_KEY + '=' + encodeURIComponent(raw) + '; path=/; max-age=7776000; SameSite=Lax' + (location.protocol === 'https:' ? '; Secure' : ''); } catch (e) { /* blocked */ }
  }
  function forServer(lead, eventId) {
    var l = landing();
    return {
      first_name: lead.first_name, last_name: lead.last_name, email: lead.email, phone: lead.phone,
      title: lead.role || '',
      monthly_ad_spend: SPEND[lead.monthly_amazon_ad_spend] || 0,
      annual_revenue: REVENUE[lead.annual_amazon_revenue] || 0,
      monthly_ad_spend_band: lead.monthly_amazon_ad_spend || '', annual_revenue_band: lead.annual_amazon_revenue || '',
      biggest_ppc_bottleneck: lead.biggest_ppc_bottleneck || '',
      consent: !!lead.consent, consent_version: lead.consent_version || '',
      source: 'hormozi_skill_funnel', offer: 'hero-asin-architect', code: 'FOUNDER10',
      event_id: eventId, fbp: cookie('_fbp'), fbc: cookie('_fbc'), gclid: l.gclid || '',
      utm_source: l.utm_source || '', utm_medium: l.utm_medium || '', utm_campaign: l.utm_campaign || '', utm_content: l.utm_content || '',
      page_url: location.origin + location.pathname, submitted_at: lead.submitted_at || new Date().toISOString()
    };
  }
  function newId(prefix) { return prefix + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 10); }

  window.FCFunnel = {
    endpoint: base,
    /* Page 1: the form was submitted. */
    lead: function (lead) {
      var id = newId('lead');
      lead.event_id = id;
      try { sessionStorage.setItem(LEAD_KEY, JSON.stringify(lead)); } catch (e) { /* storage blocked */ }
      fb('Lead', { content_name: 'Hero ASIN Architect Skill', content_category: 'Claude Code PPC Skill' }, false, id);
      ga('generate_lead', { method: 'hero_asin_skill_form' });
      saveAudit();
      return post('/api/lead', forServer(lead, id));
    },
    /* Guide page: the visitor started the read-only connect (the audit step). Sent once per visit. */
    audit: function () {
      try { if (sessionStorage.getItem('fc_audit_sent')) return false; sessionStorage.setItem('fc_audit_sent', '1'); } catch (e) { /* storage blocked */ }
      var id = newId('audit');
      fb('AuditConnectClick', { content_name: 'Hero ASIN Architect Skill' }, true, id);
      ga('skill_audit_connect', {});
      saveAudit({ connectClickedAt: new Date().toISOString() });
      var lead = storedLead();
      if (!lead || !lead.email) return false;       // the engine needs an email to act on
      var body = forServer(lead, id);
      body.stage = 'connect_clicked';
      return post('/api/audit-complete', body);
    },
    /* Small page events (video unmuted, guide opened, sample downloaded, install copied). */
    track: function (name) {
      fb(name, { content_name: 'Hero ASIN Architect Skill' }, true);
      ga('skill_' + name.replace(/([a-z])([A-Z])/g, '$1_$2').toLowerCase(), {});
    }
  };
})();
