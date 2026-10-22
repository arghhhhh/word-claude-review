// Claude document review helper for the Word "Review Document" button (ClaudeReview.bas).
// Usage: node claude-review.mjs <paragraphs.txt> <out.tsv>
//   paragraphs.txt: UTF-8, one "[n] text" or "[n|location] text" line per non-empty paragraph (written by the macro)
//   out.tsv: first line "OK" or "ERR<TAB>message"; then "SUMMARY<TAB>text";
//            then one edit per line: p, find, prefixLen, oldMiddle, newMiddle, comment, kind (E=edit, C=comment only)
// Edits are sorted last-to-first so applying them in order never shifts the ones still to come.
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { spawn } from 'node:child_process';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const here = path.dirname(fileURLToPath(import.meta.url));
const MODEL = 'opus';
const TIMEOUT_MS = 20 * 60 * 1000;

const [inFile, outFile] = process.argv.slice(2);
const clean = s => String(s ?? '').replace(/[\t\r\n\v]+/g, ' ');
// Document text keeps its tabs / manual line breaks as markers the macro turns back into Word characters.
const enc = s => String(s ?? '').replace(/\t/g, '{TAB}').replace(/\v/g, '{BR}').replace(/[\r\n]+/g, ' ');
const finish = lines => { writeFileSync(outFile, lines.join('\r\n'), 'utf8'); };
const fail = msg => { console.error('\n' + msg); finish(['ERR\t' + clean(msg)]); process.exit(1); };

// Map typographic characters to plain ones, 1:1, so indexes line up with the original text.
const norm = s => s.replace(/[‘’‛′]/g, "'").replace(/[“”‟″]/g, '"')
                  .replace(/[   \t\v]/g, ' ').replace(/[–—]/g, '-');

function locate(para, find) {
    let i = para.indexOf(find);
    if (i >= 0) return i;
    return norm(para).indexOf(norm(find));
}

// Shrink an edit to just the words that change, so the redline doesn't strike the whole phrase.
// If the fix only adds or only removes characters (DEFENDANTS -> DEFENDANT'S), mark just those characters,
// which Word shows as a single change instead of a deletion + insertion pair.
function minimize(oldS, newS) {
    let ca = 0, cb = 0;
    while (ca < oldS.length && ca < newS.length && oldS[ca] === newS[ca]) ca++;
    while (cb < oldS.length - ca && cb < newS.length - ca &&
           oldS[oldS.length - 1 - cb] === newS[newS.length - 1 - cb]) cb++;
    if (ca + cb === oldS.length || ca + cb === newS.length)
        return { prefix: ca, oldMid: oldS.slice(ca, oldS.length - cb), newMid: newS.slice(ca, newS.length - cb) };
    let a = 0;
    while (a < oldS.length && a < newS.length && oldS[a] === newS[a]) a++;
    while (a > 0 && !/\s/.test(oldS[a - 1])) a--;                 // back up to a word boundary
    let b = 0;
    while (b < oldS.length - a && b < newS.length - a &&
           oldS[oldS.length - 1 - b] === newS[newS.length - 1 - b]) b++;
    while (b > 0 && !/\s/.test(oldS[oldS.length - b])) b--;
    return { prefix: a, oldMid: oldS.slice(a, oldS.length - b), newMid: newS.slice(a, newS.length - b) };
}

// Finds the Claude Code program itself (never a shell alias or profile function).
// Official installer: %USERPROFILE%\.local\bin\claude.exe. npm install: claude.cmd on PATH, which Node
// can only start through cmd.exe.
function findClaude() {
    const local = path.join(process.env.USERPROFILE || '', '.local', 'bin', 'claude.exe');
    if (existsSync(local)) return { file: local, viaCmd: false };
    for (const dir of (process.env.PATH || '').split(path.delimiter).filter(Boolean)) {
        for (const [name, viaCmd] of [['claude.exe', false], ['claude.cmd', true]]) {
            const p = path.join(dir.replace(/^"|"$/g, ''), name);
            if (existsSync(p)) return { file: p, viaCmd };
        }
    }
    return null;
}

function runClaude(prompt) {
    return new Promise((resolve, reject) => {
        const claude = findClaude();
        if (!claude) return reject(new Error('Claude Code is not installed (no "claude" found). ' +
            'Install it from https://claude.com/claude-code and sign in, then try again.'));
        const args = ['-p', '--output-format', 'json', '--model', MODEL, '--tools', '',
                      '--no-session-persistence', '--strict-mcp-config'];
        const child = claude.viaCmd
            ? spawn(`"${claude.file}"`, args.map(a => a === '' ? '""' : a), { cwd: tmpdir(), windowsHide: true, shell: true })
            : spawn(claude.file, args, { cwd: tmpdir(), windowsHide: true });
        let out = '', err = '';
        child.stdout.on('data', d => out += d);
        child.stderr.on('data', d => err += d);
        const t0 = Date.now();
        const tick = setInterval(() => {
            const s = Math.round((Date.now() - t0) / 1000);
            process.stdout.write(`\r  Claude is reviewing the document... ${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}  `);
        }, 1000);
        const killer = setTimeout(() => { child.kill(); reject(new Error('Timed out after 20 minutes.')); }, TIMEOUT_MS);
        child.on('error', e => { clearInterval(tick); clearTimeout(killer); reject(e); });
        child.on('close', code => {
            clearInterval(tick); clearTimeout(killer);
            if (code !== 0 && !out) return reject(new Error(`claude exited with code ${code}: ${err.trim().slice(0, 500)}`));
            resolve(out);
        });
        child.stdin.end(prompt, 'utf8');
    });
}

const raw = readFileSync(inFile, 'utf8').replace(/^﻿/, '');
const paras = new Map();
for (const line of raw.split(/\r?\n/)) {
    const m = /^\[(\d+)(?:\|[^\]]*)?\] (.*)$/.exec(line);   // [12] text  or  [205|Footer, section 1] text
    if (m) paras.set(Number(m[1]), m[2]);
}
if (!paras.size) fail('The document has no text to review.');

console.log(`Claude Review: ${paras.size} paragraphs (${raw.length.toLocaleString()} characters).`);
console.log('Word will wait until this finishes. You can leave this window alone.\n');

let template;
try { template = readFileSync(path.join(here, 'review-prompt.md'), 'utf8'); }
catch { fail('Missing review-prompt.md next to claude-review.mjs.'); }
const prompt = template.replace('{{DOCUMENT}}', raw.trim());

let reply;
try {
    const out = await runClaude(prompt);
    const env = JSON.parse(out);
    if (env.is_error) fail('Claude returned an error: ' + (env.result || env.subtype || 'unknown'));
    reply = env.result || '';
} catch (e) { fail('Could not run Claude: ' + e.message); }

let data;
try {
    const s = reply.indexOf('{'), e = reply.lastIndexOf('}');
    data = JSON.parse(reply.slice(s, e + 1));
} catch { fail('Claude did not return a readable review. First part of its reply:\n' + reply.slice(0, 400)); }

const edits = [], missed = [];
for (const ed of Array.isArray(data.edits) ? data.edits : []) {
    const p = Number(ed.p), para = paras.get(p), find = String(ed.find ?? '');
    const comment = String(ed.comment ?? '').trim();
    const pos = para && find ? locate(para, find) : -1;
    if (pos < 0) { missed.push({ p, find, replace: ed.replace, comment }); continue; }
    const actual = para.substr(pos, find.length);            // exact text as it appears in Word
    const isEdit = typeof ed.replace === 'string' && norm(ed.replace) !== norm(actual);
    let prefix = 0, oldMid = actual, newMid = '';
    if (isEdit) ({ prefix, oldMid, newMid } = minimize(actual, ed.replace));
    edits.push({ p, pos, actual, prefix, oldMid, newMid, comment, kind: isEdit ? 'E' : 'C' });
}
edits.sort((x, y) => y.p - x.p || y.pos - x.pos);

let summary = clean(data.summary || '');
if (missed.length) {
    summary += ` [${missed.length} suggestion(s) couldn't be matched to the text and are listed here: ` +
        missed.map(m => `¶${m.p}: "${clean(m.find)}"` + (m.replace != null ? ` → "${clean(m.replace)}"` : '') +
                        (m.comment ? ` (${clean(m.comment)})` : '')).join('; ') + ']';
}

finish(['OK', 'SUMMARY\t' + summary,
    ...edits.map(e => [e.p, enc(e.actual), e.prefix, enc(e.oldMid), enc(e.newMid), clean(e.comment), e.kind].join('\t'))]);
console.log(`\n\nDone: ${edits.filter(e => e.kind === 'E').length} edits, ${edits.filter(e => e.kind === 'C').length} comments` +
            (missed.length ? `, ${missed.length} unmatched (listed in the summary comment).` : '.'));
