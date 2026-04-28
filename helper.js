#!/usr/bin/env node
/**
 * KSOB PodcastBoard — local helper
 * Creates Microsoft Outlook drafts with the info deck PDF attached.
 *
 * Usage:
 *   node helper.js
 *
 * Optional env overrides:
 *   PDF=/path/to/file.pdf  PORT=9876
 */

const http = require('http');
const fs   = require('fs');
const path = require('path');
const os   = require('os');
const { execFileSync } = require('child_process');

const PDF  = process.env.PDF  || path.join(os.homedir(), 'Desktop', 'Podcast Guest Info_2025.pdf');
const PORT = parseInt(process.env.PORT || '9876', 10);

if (!fs.existsSync(PDF)) {
  console.warn(`⚠  PDF not found at: ${PDF}`);
  console.warn('   Set the PDF env var if it lives elsewhere, e.g.:');
  console.warn('   PDF=~/Documents/deck.pdf node helper.js');
}

function buildAppleScript(to, subject, body) {
  // Escape for AppleScript double-quoted string: backslash then quote
  const esc = s => String(s)
    .replace(/\\/g, '\\\\')
    .replace(/"/g,  '\\"');

  // Build body as concatenated string to handle newlines safely
  const lines = body.split('\n');
  const asBody = lines
    .map(l => `"${esc(l)}"`)
    .join(' & return & ');

  return `tell application "Microsoft Outlook"
  set msgBody to ${asBody}
  set newMsg to make new outgoing message with properties {subject:"${esc(subject)}", content:msgBody}
  make new to recipient at newMsg with properties {email address:{address:"${esc(to)}"}}
  make new attachment at newMsg with properties {file:POSIX file "${esc(PDF)}"}
  open newMsg
  activate
end tell`;
}

const server = http.createServer((req, res) => {
  // CORS — allow the Vercel app and localhost
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return; }
  if (req.method !== 'POST')    { res.writeHead(405); res.end('Method Not Allowed'); return; }

  let raw = '';
  req.on('data', c => raw += c);
  req.on('end', () => {
    try {
      const { to, subject, body } = JSON.parse(raw);
      if (!to || !subject || !body) throw new Error('Missing to / subject / body');

      const script   = buildAppleScript(to, subject, body);
      const tmpFile  = path.join(os.tmpdir(), `ksob_${Date.now()}.applescript`);
      fs.writeFileSync(tmpFile, script, 'utf8');

      try {
        execFileSync('osascript', [tmpFile]);
      } finally {
        fs.unlinkSync(tmpFile);
      }

      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true }));
    } catch (err) {
      console.error('Draft error:', err.message);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false, error: err.message }));
    }
  });
});

server.listen(PORT, '127.0.0.1', () => {
  console.log(`✅  KSOB helper running on http://127.0.0.1:${PORT}`);
  console.log(`📎  PDF attachment: ${PDF}`);
  console.log('    Open the board and click "Draft in Outlook" — Outlook will open with the PDF attached.');
  console.log('    Press Ctrl+C to stop.');
});
