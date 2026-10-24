// Maya on WhatsApp.
//
// Meta calls this function (https://<site>.netlify.app/whatsapp) for every
// message sent to the Sindooram WhatsApp number. For each guest text it:
//   1. checks the request really came from Meta (X-Hub-Signature-256),
//   2. skips anything already answered (Meta re-delivers slow replies),
//   3. adds today's date and the live booked nights (the public
//      availability feed — Airbnb, Booking.com and Direct) as context,
//   4. asks Claude for Maya's reply, with the conversation so far,
//   5. sends the reply on WhatsApp, and
//   6. if Maya flagged [ESCALATE: …], alerts Negin by email via the
//      alerts Apps Script (alerts-apps-script.gs) and logs it to the Sheet.
//
// Every key and URL lives in Netlify's environment variables — this file is
// public on GitHub. See the setup steps for the list.

import crypto from 'node:crypto';
import Anthropic from '@anthropic-ai/sdk';
import { getStore } from '@netlify/blobs';
import SYSTEM_PROMPT from './prompt.generated.mjs';

export const config = { path: '/whatsapp' };

const env = (key, fallback) => process.env[key] || fallback;

const MODEL = env('CLAUDE_MODEL', 'claude-sonnet-5');
const GRAPH_URL = 'https://graph.facebook.com/' + env('GRAPH_API_VERSION', 'v23.0');
const HISTORY_MESSAGES = 16;                    // kept per guest (8 back-and-forths)
const HISTORY_TTL_MS = 3 * 24 * 60 * 60 * 1000;  // a chat idle for 3 days starts fresh
const DAILY_REPLY_CAP = 30;                     // per guest, keeps API costs bounded
const AVAILABILITY_FRESH_MS = 5 * 60 * 1000;    // re-read the feed at most every 5 min
const AVAILABILITY_STALE_MS = 60 * 60 * 1000;   // never trust a copy older than 1 hour
const FALLBACK_REPLY = "Sorry, I'm having a little trouble right now 🙏 Negin will get back to you shortly!";

export default async (req, context) => {
  if (req.method === 'GET') return verifyWebhook(new URL(req.url));
  if (req.method !== 'POST') return new Response('Method not allowed', { status: 405 });

  const raw = await req.text();
  if (!validSignature(raw, req.headers.get('x-hub-signature-256'))) {
    return new Response('Bad signature', { status: 401 });
  }
  let update;
  try { update = JSON.parse(raw); } catch { return new Response('Bad JSON', { status: 400 }); }

  // Acknowledge Meta straight away and do the slow part (Claude) after,
  // where the platform allows it — otherwise Meta retries and the guest
  // could get the same answer twice (the "seen" check below stops that too).
  const work = handleUpdate(update).catch((err) => console.error('handleUpdate failed:', err));
  if (context && typeof context.waitUntil === 'function') context.waitUntil(work);
  else await work;
  return new Response('OK', { status: 200 });
};

// Meta's one-time check when the webhook URL is saved in the dashboard.
function verifyWebhook(url) {
  const p = url.searchParams;
  const token = env('WHATSAPP_VERIFY_TOKEN');
  if (token && p.get('hub.mode') === 'subscribe' && p.get('hub.verify_token') === token) {
    return new Response(p.get('hub.challenge') || '', { status: 200 });
  }
  return new Response('Forbidden', { status: 403 });
}

function validSignature(raw, header) {
  const secret = env('WHATSAPP_APP_SECRET');
  if (!secret || !header || !header.startsWith('sha256=')) return false;
  const expected = crypto.createHmac('sha256', secret).update(raw, 'utf8').digest();
  const given = Buffer.from(header.slice('sha256='.length), 'hex');
  return given.length === expected.length && crypto.timingSafeEqual(given, expected);
}

async function handleUpdate(update) {
  for (const entry of update.entry || []) {
    for (const change of entry.changes || []) {
      const value = change.value || {};
      // Only our own number (ignore events for any other number on the app).
      const numberId = value.metadata && value.metadata.phone_number_id;
      if (numberId && numberId !== env('WHATSAPP_PHONE_NUMBER_ID')) continue;
      const names = {};
      (value.contacts || []).forEach((c) => { names[c.wa_id] = c.profile && c.profile.name; });
      // value.statuses (sent / delivered / read receipts) need no action.
      for (const msg of value.messages || []) {
        try { await handleMessage(msg, names[msg.from]); }
        catch (err) { console.error('message ' + msg.id + ' failed:', err); }
      }
    }
  }
}

async function handleMessage(msg, name) {
  const phone = String(msg.from || '');
  const allowed = env('ALLOWED_NUMBERS');
  if (allowed && !allowed.split(',').map((n) => n.trim()).includes(phone)) return;

  const store = getStore('maya-whatsapp');
  const seenKey = 'seen/' + msg.id;
  if (await store.get(seenKey)) return;
  await store.set(seenKey, String(Date.now()));

  graph('/messages', { messaging_product: 'whatsapp', status: 'read', message_id: msg.id })
    .catch((err) => console.warn('mark-read failed:', err.message));

  if (msg.type === 'reaction') return;
  if (msg.type !== 'text' || !msg.text || !msg.text.body) {
    await sendText(phone, "Sorry, I can only read text messages for now 🙏 Could you type your question?");
    return;
  }
  const text = msg.text.body.slice(0, 2000);

  const today = indiaToday();
  const convKey = 'conv/' + phone;
  let conv = await store.get(convKey, { type: 'json' });
  if (!conv || Date.now() - conv.updatedAt > HISTORY_TTL_MS) conv = { messages: [] };
  if (conv.day !== today) { conv.day = today; conv.repliesToday = 0; conv.capAlerted = false; }

  if (conv.repliesToday >= DAILY_REPLY_CAP) {
    if (!conv.capAlerted) {
      conv.capAlerted = true;
      await alertTeam({ phone, name, reason: 'Guest passed ' + DAILY_REPLY_CAP + ' messages today — Maya has stopped replying until tomorrow', messages: conv.messages.concat([{ role: 'user', content: text }]) });
    }
    conv.updatedAt = Date.now();
    await store.setJSON(convKey, conv);
    return;
  }

  let reply;
  let reason = null;
  try {
    const context = await buildContext(name, phone, today);
    const messages = mergeTurns(conv.messages.concat([
      { role: 'user', content: context + '\n\nGuest message:\n' + text }
    ]));
    // One retry at most — a guest waiting on WhatsApp is better served by the
    // fallback reply (and Negin) than by a long silence.
    const client = new Anthropic({ apiKey: env('ANTHROPIC_API_KEY'), maxRetries: 1, timeout: 25000 });
    const res = await client.messages.create({
      model: MODEL,
      max_tokens: 700,
      // The prompt is identical on every call — cache it.
      system: [{ type: 'text', text: SYSTEM_PROMPT, cache_control: { type: 'ephemeral' } }],
      messages
    });
    reply = res.content.filter((b) => b.type === 'text').map((b) => b.text).join('').trim();
    const flag = reply.match(/\[ESCALATE:\s*([^\]]*)\]/i);
    if (flag) {
      reason = flag[1].trim() || 'Maya escalated';
      reply = reply.replace(/\s*\[ESCALATE:[^\]]*\]\s*/gi, '\n').trim();
    }
    if (!reply) reply = "Let me get Negin to help you with that — he'll message you shortly 🙌";
  } catch (err) {
    console.error('Claude call failed:', err);
    reply = FALLBACK_REPLY;
    reason = 'Maya could not answer (technical error: ' + String(err.message || err).slice(0, 120) + ')';
  }

  await sendText(phone, reply.slice(0, 4000));

  conv.messages = conv.messages.concat([{ role: 'user', content: text }, { role: 'assistant', content: reply }]).slice(-HISTORY_MESSAGES);
  while (conv.messages.length && conv.messages[0].role !== 'user') conv.messages.shift();
  conv.repliesToday = (conv.repliesToday || 0) + 1;
  conv.updatedAt = Date.now();
  await store.setJSON(convKey, conv);

  if (reason) await alertTeam({ phone, name, reason, messages: conv.messages });
}

// Context Maya needs that the guest didn't type: today's date (India), who
// is writing, and the booked nights from the live calendar.
async function buildContext(name, phone, today) {
  const weekday = new Intl.DateTimeFormat('en-IN', { timeZone: 'Asia/Kolkata', weekday: 'long' }).format(new Date());
  const lines = [
    '[Context from the system — not written by the guest]',
    'Today (India): ' + weekday + ', ' + today,
    'Guest: ' + (name || 'name not shared') + ', +' + phone
  ];
  const avail = await getAvailability();
  if (!avail) {
    lines.push('The live calendar could not be loaded right now. Do not say any dates are available — say you will confirm with Negin, and escalate.');
  } else if (!avail.ranges.length) {
    lines.push('Calendar ' + avail.from + ' to ' + avail.until + ': no nights are booked.');
  } else {
    lines.push('Calendar ' + avail.from + ' to ' + avail.until + '. Booked stays (check-in → check-out; the check-out day itself is free for a new arrival):');
    avail.ranges.forEach((r) => lines.push('- ' + r.start + ' → ' + r.end));
  }
  return lines.join('\n');
}

async function getAvailability() {
  const store = getStore('maya-whatsapp');
  const cached = await store.get('availability', { type: 'json' });
  const age = cached ? Date.now() - cached.fetchedAt : Infinity;
  if (age < AVAILABILITY_FRESH_MS) return cached.data;
  try {
    const res = await fetch(env('AVAILABILITY_FEED_URL'), { redirect: 'follow' });
    const data = await res.json();
    if (!data || data.error || !Array.isArray(data.ranges)) {
      throw new Error('unexpected feed response: ' + JSON.stringify(data).slice(0, 200));
    }
    await store.setJSON('availability', { fetchedAt: Date.now(), data });
    return data;
  } catch (err) {
    console.error('availability feed failed:', err);
    return age < AVAILABILITY_STALE_MS ? cached.data : null;
  }
}

// The Messages API wants user/assistant turns to alternate; a failed earlier
// reply could leave two guest messages in a row, so join those.
function mergeTurns(messages) {
  const out = [];
  messages.forEach((m) => {
    const last = out[out.length - 1];
    if (last && last.role === m.role) last.content += '\n\n' + m.content;
    else out.push({ role: m.role, content: m.content });
  });
  return out;
}

async function graph(path, payload) {
  const res = await fetch(GRAPH_URL + '/' + env('WHATSAPP_PHONE_NUMBER_ID') + path, {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + env('WHATSAPP_TOKEN'), 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
  if (!res.ok) throw new Error('WhatsApp API ' + res.status + ': ' + (await res.text()).slice(0, 300));
  return res.json();
}

function sendText(to, body) {
  return graph('/messages', {
    messaging_product: 'whatsapp',
    recipient_type: 'individual',
    to,
    type: 'text',
    text: { preview_url: true, body }
  });
}

// Emails Negin (and logs to the Sheet) through the alerts Apps Script.
async function alertTeam({ phone, name, reason, messages }) {
  const url = env('ALERT_WEBHOOK_URL');
  if (!url) { console.warn('ALERT_WEBHOOK_URL not set — escalation not sent:', reason); return; }
  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ secret: env('ALERT_SECRET'), phone, name: name || '', reason, messages: (messages || []).slice(-8) }),
      redirect: 'follow'
    });
    const out = await res.text();
    if (!res.ok || /"error"/.test(out)) console.error('alert failed:', res.status, out.slice(0, 300));
  } catch (err) {
    console.error('alert failed:', err);
  }
}

function indiaToday() {
  return new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Kolkata', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
}
