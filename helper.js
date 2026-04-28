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

function buildAppleScript(to, subject, htmlPath) {
  // Escape for AppleScript double-quoted string: backslash then quote
  const esc = s => String(s)
    .replace(/\\/g, '\\\\')
    .replace(/"/g,  '\\"');

  return `tell application "Microsoft Outlook"
  set htmlContent to do shell script "cat " & quoted form of "${esc(htmlPath)}"
  set newMsg to make new outgoing message with properties {subject:"${esc(subject)}", html content:htmlContent}
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
      const { to, subject, body, htmlBody } = JSON.parse(raw);
      if (!to || !subject || !body) throw new Error('Missing to / subject / body');

      const ts       = Date.now();
      const htmlPath = path.join(os.tmpdir(), `ksob_body_${ts}.html`);
      const tmpFile  = path.join(os.tmpdir(), `ksob_${ts}.applescript`);

      // Write HTML body to temp file so AppleScript can read it with `cat`
      fs.writeFileSync(htmlPath, htmlBody || body, 'utf8');
      const script = buildAppleScript(to, subject, htmlPath);
      fs.writeFileSync(tmpFile, script, 'utf8');

      try {
        execFileSync('osascript', [tmpFile]);
      } finally {
        fs.unlinkSync(tmpFile);
        try { fs.unlinkSync(htmlPath); } catch (_) {}
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
