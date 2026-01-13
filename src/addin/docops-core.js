// C:\inetpub\wwwroot\addin\docops-core.js
// ============================================================================
// DocOps Core Engine
// Универсальный движок для обработки addin.block операций
// Используется как в локальной, так и в серверной панели
// ============================================================================
(function (global) {
    'use strict';
    // ========================================================================
    // NAMESPACE
    // ========================================================================
    const DocOpsCore = {};
    // ========================================================================
    // CONFIGURATION
    // ========================================================================
    DocOpsCore.config = {
        logsBase: 'https://localhost:8001',
        logsToken: '',
        logsMinLevel: 'debug',
        testIgnoreAnchor: false,
        testSkipDocx: false
    };
    // ========================================================================
    // LOGGING
    // ========================================================================
    function _getLogsToken() {
        const m = document.querySelector('meta[name="imc-logs-token"]');
        return m ? (m.content || '').trim() : DocOpsCore.config.logsToken;
    }
    function _getLogsMinLevel() {
        const m = document.querySelector('meta[name="imc-logs-level"]');
        const v = (m ? (m.content || '') : '').trim().toLowerCase() || DocOpsCore.config.logsMinLevel;
        const map = { debug: 10, info: 20, warn: 30, error: 40 };
        return map[v] || 10;
    }
    function _logsBase() {
        const m = document.querySelector('meta[name="imc-logs-url"]');
        return (m ? (m.content || '').trim() : '') || DocOpsCore.config.logsBase;
    }
    DocOpsCore.plogSend = async function plogSend(level, event, message, data) {
        try {
            const token = _getLogsToken();
            const base = _logsBase();
            const minLvl = _getLogsMinLevel();
            const lvlMap = { debug: 10, info: 20, warn: 30, error: 40 };
            const lvl = lvlMap[(level || 'info').toLowerCase()] || 20;
            if (lvl < minLvl) return true;

            const MAX = 2048;
            const safeData = (data && typeof data === 'object') ? JSON.parse(JSON.stringify(data)) : data;

            try {
                if (safeData && typeof safeData === 'object') {
                    ['text', 'stack'].forEach(k => {
                        if (typeof safeData[k] === 'string' && safeData[k].length > MAX) {
                            safeData[k] = safeData[k].slice(0, MAX) + '';
                            safeData.truncated = true;
                        }
                    });
                }
            } catch (_) { }

            const payload = {
                level: (level || 'info').toLowerCase(),
                phase: 'addin',
                event: event || 'log',
                message: message || '',
                email: (global.__addin_email || ''),
                jobId: (global.__job_id || null),
                traceId: (global.__trace_id || null),
                data: safeData || {}
            };

            const headers = { 'Content-Type': 'application/json' };
            if (token) headers['X-IMC-Logs-Token'] = token;

            const paths = ["/api/logs/ingest/", "/logs/api/logs/ingest/"];
            for (const p of paths) {
                try {
                    const r = await fetch(base + p, {
                        method: 'POST',
                        headers,
                        body: JSON.stringify(payload),
                        credentials: 'omit',
                        mode: 'cors',
                    });
                    if (r.ok) {
                        await r.json().catch(() => null);
                        return true;
                    }
                    if (r.status === 404) {
                        console.log("[plog] 404 on", p, " trying next");
                        continue;
                    }
                    const txt = await r.text().catch(() => "");
                    console.log("[plog] status", r.status, "(not ok) ", txt);
                    return false;
                } catch (e) {
                    console.log("[plog] network error:", e?.message || e);
                }
            }
        } catch (e) {
            console.log("[plog] internal error:", e && (e.message || e));
        }
        return false;
    };
    DocOpsCore.dbgStep = function dbgStep(name, extra) {
        try {
            return DocOpsCore.plogSend('debug', 'docx.step', name, extra || {});
        } catch (_) {
            return Promise.resolve(false);
        }
    };
    // ========================================================================
    // UTILITIES
    // ========================================================================
    DocOpsCore.testIgnoreAnchor = function testIgnoreAnchor() {
        return ((document.querySelector('meta[name="imc-test-ignore-anchor"]')?.content || '').trim() === '1')
            || (global.__TEST_IGNORE_ANCHOR === true)
            || DocOpsCore.config.testIgnoreAnchor;
    };
    DocOpsCore.isAnchorText = function isAnchorText(t) {
        return (/^\s*</).test(String(t || '').trim());
    };
    function _isEmptyText(s) {
        return !String(s || '').replace(/[\s\u00A0\r\n\t]+/g, '');
    }
    function genOpId() {
        return 'op-' + Math.random().toString(36).slice(2, 9) + '-' + Date.now().toString(36);
    }
    function cloneOpData(data) {
        try {
            return (typeof structuredClone === 'function')
                ? structuredClone(data)
                : JSON.parse(JSON.stringify(data));
        } catch (e) {
            return data;
        }
    }
    // ========================================================================
    // ANCHOR HANDLING
    // ========================================================================
    function __normAnchorCode(raw) {
        let s = String(raw || '').trim().replace(/^anchor:/i, '');
        s = s.replace(/^<[\s\u00A0]*/, '').replace(/[\s\u00A0]*>?$/, '');
        s = s.replace(/\u00A0/g, ' ').replace(/[ \t]+/g, '');
        const m = s.match(/^[A-Za-zА-Яа-яЁё0-9][A-Za-zА-Яа-яЁё0-9.\-]*/);
        return m ? m[0] : '';
    }
    DocOpsCore.gotoAnchor = function gotoAnchor(anchorText) {
        try {
            if (DocOpsCore.testIgnoreAnchor()) return Promise.resolve(false);
        } catch (_) { }

        anchorText = String(anchorText || '').trim();
        if (!anchorText) return Promise.resolve(false);

        const code = __normAnchorCode(anchorText);
        if (!code) {
            try {
                DocOpsCore.plogSend('warn', 'anchor.parse.fail', '', { raw: anchorText });
            } catch (_) { }
            return Promise.resolve(false);
        }

        return Word.run(async (ctx) => {
            try {
                await DocOpsCore.plogSend('debug', 'anchor.parse', '', { raw: anchorText, code });
            } catch (_) { }

            const SP = ['', ' ', '\u00A0'];
            const cands = [];
            for (const L of SP) for (const R of SP) {
                cands.push(`<${L}${code}${R}>`);
                cands.push(`<${L}${code}${R}`);
            }

            const ordered = cands.sort((a, b) => (b.endsWith('>') ? 1 : 0) - (a.endsWith('>') ? 1 : 0));

            let found = null, matched = null;
            for (const cand of ordered) {
                const res = ctx.document.body.search(cand, { matchCase: false, matchWholeWord: false, matchWildcards: false });
                ctx.load(res, 'items/text');
                await ctx.sync();

                try {
                    await DocOpsCore.plogSend('debug', 'anchor.try', '', { cand, hits: (res.items || []).length });
                } catch (_) { }

                if (res.items && res.items.length) {
                    const r = res.items[0];
                    if (String(r.text || '') === cand) {
                        found = r;
                        matched = cand;
                        break;
                    }
                }
            }

            if (!found) {
                try {
                    await DocOpsCore.plogSend('warn', 'anchor.not_found', '', { code });
                } catch (_) { }
                return false;
            }

            const para = found.paragraphs.getFirst();
            await ctx.sync();

            const atAfter = para.getRange(Word.RangeLocation.after);
            atAfter.select();
            await ctx.sync();

            try {
                await DocOpsCore.plogSend('info', 'anchor.moved', '', { code, matched, mode: 'after-boundary' });
            } catch (_) { }

            try {
                global.__caret_set_by_anchor = true;
                global.__lastAnchorCode = code;
                global.__anchorAfterTag = null;
            } catch (_) { }

            return true;
        });
    };
    // ========================================================================
    // WORD UTILITIES
    // ========================================================================
    DocOpsCore.pickSafeInsertHost = async function pickSafeInsertHost(ctx, rangeLike) {
        // Default implementation - can be overridden
        return rangeLike;
    };
    async function safeDeleteIfVisuallyEmptyParagraph(ctx, p) {
        try {
            p.load("text");
            await ctx.sync();

            if (!_isEmptyText(p.text)) return false;

            const xmlRes = p.getRange().getOoxml();
            await ctx.sync();

            const xml = String(xmlRes.value || "");
            if (/\b(w:drawing|w:pict|w:object|w:tbl)\b/i.test(xml)) return false;
            if (/<w:t\b[^>]*>[^<]*\S[^<]*<\/w:t>/i.test(xml)) return false;

            p.getRange().delete();
            await ctx.sync();

            return true;
        } catch (_) {
            return false;
        }
    }
    async function safeInsertFileFromBase64(ctx, rangeLike, base64) {
        try {
            await DocOpsCore.plogSend('debug', 'docx.insert.try', '', { len: (base64 || '').length });
        } catch (_) { }

        try {
            rangeLike.insertFileFromBase64(base64, Word.InsertLocation.replace, { importStyles: false });
            await ctx.sync();
            try {
                await DocOpsCore.plogSend('info', 'docx.insert.ok', '', { via: 'range.replace(importStyles:false)' });
            } catch (_) { }
            return true;
        } catch (e1) {
            try {
                rangeLike.insertFileFromBase64(base64, Word.InsertLocation.replace);
                await ctx.sync();
                try {
                    await DocOpsCore.plogSend('info', 'docx.insert.ok', '', { via: 'range.replace(fallback-no-options)', err1: String(e1 && e1.message || e1) });
                } catch (_) { }
                return true;
            } catch (e2) {
                try {
                    ctx.document.insertFileFromBase64(base64, Word.InsertLocation.end, { importStyles: false });
                    await ctx.sync();
                    try {
                        await DocOpsCore.plogSend('warn', 'docx.insert.ok', '', { via: 'document.end(fallback)', err1: String(e1), err2: String(e2) });
                    } catch (_) { }
                    return true;
                } catch (e3) {
                    try {
                        await DocOpsCore.plogSend('error', 'docx.insert.fail', String(e3 && e3.message || e3), { e1: String(e1), e2: String(e2), e3: String(e3) });
                    } catch (_) { }
                    throw e3;
                }
            }
        }
    }
    // ========================================================================
    // FOOTNOTES
    // ========================================================================
    async function wrapFootnoteReferenceWithBrackets(ctx, footnote) {
        if (!footnote) return;
        try {
            footnote.load("reference");
            await ctx.sync();

            const ref = footnote.reference;
            const left = ref.insertText("<", Word.InsertLocation.before);
            const right = ref.insertText(">", Word.InsertLocation.after);

            try {
                left.font.superscript = true;
            } catch (_) { }
            try {
                right.font.superscript = true;
            } catch (_) { }

            await ctx.sync();
        } catch (_) { }
    }
    DocOpsCore.convertEmbeddedMarkersToFootnotes = function convertEmbeddedMarkersToFootnotes(marker = "&SRC&", scope = "last") {
        marker = String(marker || "&SRC&");
        scope = DocOpsCore.testIgnoreAnchor() ? "doc" : (scope === "doc" ? "doc" : "last");

        return Word.run(async (ctx) => {
            const t0 = Date.now();
            const attempted = ["A.expandTo", "B.wildcards", "C.by-paragraph"];
            let stageUsed = null;

            await DocOpsCore.plogSend('info', 'footnotes.conv.start', '', { scope, marker });

            let root = ctx.document.body, useLast = false;
            if (scope === "last" && (global.__lastDocOpsCCTag || global.__anchorAfterTag)) {
                const tag = String(global.__lastDocOpsCCTag || global.__anchorAfterTag);
                const cc = ctx.document.contentControls.getByTag(tag).getFirstOrNullObject();
                ctx.load(cc, "isNullObject");
                await ctx.sync();

                if (!cc.isNullObject) {
                    root = (cc.getRange ? cc.getRange() : cc);
                    useLast = true;
                }
            }

            const probe = root.search(marker, { matchCase: false, matchWholeWord: false });
            probe.load("items");
            await ctx.sync();

            const markersBefore = (probe.items || []).length;
            if (markersBefore < 2) {
                await DocOpsCore.plogSend('debug', 'footnotes.result', '', { scope, applied: 0, markersBefore, markersLeft: markersBefore, result: 'none-found' });
                return 0;
            }

            let applied = 0;

            // Stage A: expandTo
            try {
                const hits = root.search(marker, { matchCase: false, matchWholeWord: false, matchWildcards: false });
                hits.load("items");
                await ctx.sync();

                const items = hits.items || [];
                if (items.length >= 2) {
                    for (let i = 0; i + 1 < items.length; i += 2) {
                        const r1 = items[i].getRange ? items[i].getRange() : items[i];
                        const r2 = items[i + 1].getRange ? items[i + 1].getRange() : items[i + 1];
                        const seg = r1.expandTo ? r1.expandTo(r2) : null;
                        if (!seg) break;

                        seg.load("text,paragraphs");
                        await ctx.sync();

                        const raw = String(seg.text || "");
                        const foot = raw.replace(new RegExp("^" + marker), "").replace(new RegExp(marker + "$"), "");

                        let host = null;
                        try {
                            const p = seg.paragraphs.getFirst();
                            const tbl = p.parentTable;
                            ctx.load(tbl, "isNullObject");
                            await ctx.sync();

                            if (tbl && !tbl.isNullObject) {
                                const tr = tbl.getRange ? tbl.getRange() : tbl;
                                host = tr.getRange ? tr.getRange(Word.RangeLocation.after) : tr;
                            } else {
                                host = p.getRange ? p.getRange(Word.RangeLocation.end) : p;
                            }
                        } catch (_) { }

                        if (!host) host = seg;

                        try {
                            const norm = String(foot || "").replace(/^[ \t\u00A0]+/, "");
                            const fn = host.insertFootnote("\t" + norm);
                            if (fn) {
                                await wrapFootnoteReferenceWithBrackets(ctx, fn);
                            }
                        } catch (_) { }

                        try {
                            seg.insertText("", Word.InsertLocation.replace);
                        } catch (_) { }

                        await ctx.sync();
                        applied++;
                        if (!stageUsed) stageUsed = "A.expandTo";
                    }
                }
            } catch (_) { }

            // Stage B: wildcards
            if (!applied) {
                while (true) {
                    const pattern = marker + "[!^13]@" + marker;
                    const res = root.search(pattern, { matchWildcards: true, matchCase: false, matchWholeWord: false });
                    res.load("items");
                    await ctx.sync();

                    if (!res.items.length) break;

                    const seg = res.items[0];
                    seg.load("text,paragraphs");
                    await ctx.sync();

                    const raw = String(seg.text || "");
                    const foot = raw.replace(new RegExp("^" + marker), "").replace(new RegExp(marker + "$"), "");

                    let host = null;
                    try {
                        const p = seg.paragraphs.getFirst();
                        const tbl = p.parentTable;
                        ctx.load(tbl, "isNullObject");
                        await ctx.sync();

                        if (tbl && !tbl.isNullObject) {
                            const tr = tbl.getRange ? tbl.getRange() : tbl;
                            host = tr.getRange ? tr.getRange(Word.RangeLocation.after) : tr;
                        } else {
                            host = p.getRange ? p.getRange(Word.RangeLocation.end) : p;
                        }
                    } catch (_) { }

                    if (!host) host = seg;

                    try {
                        const norm = String(foot || "").replace(/^[ \t\u00A0]+/, "");
                        const fn = host.insertFootnote("\t" + norm);
                        if (fn) {
                            await wrapFootnoteReferenceWithBrackets(ctx, fn);
                        }
                    } catch (_) { }

                    try {
                        seg.insertText("", Word.InsertLocation.replace);
                    } catch (_) { }

                    await ctx.sync();
                    applied++;
                    if (!stageUsed) stageUsed = "B.wildcards";
                }
            }

            // Stage C: by-paragraph
            if (!applied) {
                const ps = root.paragraphs;
                ps.load("items/text");
                await ctx.sync();

                for (const p of (ps.items || [])) {
                    const s = String(p.text || "");
                    const i1 = s.indexOf(marker);
                    const i2 = (i1 >= 0) ? s.indexOf(marker, i1 + marker.length) : -1;

                    if (i1 >= 0 && i2 > i1) {
                        const foot = s.slice(i1 + marker.length, i2);
                        const newS = s.slice(0, i1) + s.slice(i2 + marker.length);

                        try {
                            const pr = p.getRange ? p.getRange() : p;
                            pr.insertText(newS, Word.InsertLocation.replace);
                            await ctx.sync();
                        } catch (_) { }

                        let host = null;
                        try {
                            const tbl = p.parentTable;
                            ctx.load(tbl, "isNullObject");
                            await ctx.sync();

                            if (tbl && !tbl.isNullObject) {
                                const tr = tbl.getRange ? tbl.getRange() : tbl;
                                host = tr.getRange ? tr.getRange(Word.RangeLocation.after) : tr;
                            } else {
                                host = p.getRange ? p.getRange(Word.RangeLocation.end) : p;
                            }
                        } catch (_) { }

                        if (!host) host = p;

                        try {
                            const norm = String(foot || "").replace(/^[ \t\u00A0]+/, "");
                            const fn = host.insertFootnote(" \t" + norm);
                            if (fn) {
                                await wrapFootnoteReferenceWithBrackets(ctx, fn);
                            }
                        } catch (_) { }

                        applied++;
                        if (!stageUsed) stageUsed = "C.by-paragraph";
                    }
                }
            }

            let leftCount = 0;
            try {
                const left = root.search(marker, { matchCase: false, matchWholeWord: false });
                left.load("items");
                await ctx.sync();
                leftCount = (left.items || []).length;
            } catch (_) { }

            const level = applied ? 'info' : (markersBefore ? 'warn' : 'debug');
            await DocOpsCore.plogSend(level, 'footnotes.result', '', { scope, applied, markersBefore, markersLeft: leftCount, stageUsed, durationMs: Date.now() - t0 });

            return applied;
        });
    };
    // ========================================================================
    // CAPTION HANDLING
    // ========================================================================
    DocOpsCore.fixCaptionStylesHard = function fixCaptionStylesHard() {
        return Word.run(async function (ctx) {
            const paras = ctx.document.body.paragraphs;
            paras.load('items/style,text');
            await ctx.sync();

            let fixedCount = 0;

            for (let i = 0; i < paras.items.length; i++) {
                const para = paras.items[i];
                const s = String(para.style || "").trim();
                const t = String(para.text || "").replace(/\s+/g, " ").trim();

                if (/^Название объекта\d+$/.test(s)) {
                    try {
                        para.style = "Название объекта";
                    } catch (_) {
                        try {
                            para.styleBuiltIn = Word.Style.caption;
                        } catch (_) {
                            try {
                                para.styleBuiltIn = "Caption";
                            } catch (_) { }
                        }
                    }
                    fixedCount++;
                    continue;
                }

                if (/^(Табл\.?|Таблица|Рис\.?|Рисунок|Table|Figure)\b/i.test(t)) {
                    try {
                        para.styleBuiltIn = Word.Style.caption;
                    } catch (_) {
                        try {
                            para.styleBuiltIn = "Caption";
                        } catch (_) { }
                    }
                    fixedCount++;
                }
            }

            if (fixedCount > 0) {
                await ctx.sync();
                DocOpsCore.plogSend('info', 'caption_styles.fixed', `Fixed ${fixedCount} caption styles`);
            }
        });
    };
    DocOpsCore.flattenForeignContentControls = function flattenForeignContentControls() {
        return Word.run(function (ctx) {
            const all = ctx.document.contentControls;
            ctx.load(all, "items,tag");

            return ctx.sync().then(function () {
                let chain = Promise.resolve();
                (all.items || []).forEach(function (cc) {
                    chain = chain.then(function () {
                        const t = String(cc.tag || "");
                        if (!t || t.indexOf("DocOps") !== 0) {
                            try {
                                cc.delete(false);
                            } catch (_) { }
                        }
                        return ctx.sync().catch(function () { });
                    });
                });

                return chain;
            });
        });
    };
    // ========================================================================
    // BULLETPROOF QUEUE
    // ========================================================================
    function initBulletproofQueue() {
        if (!global.__bulletproofQueue) {
            global.__bulletproofQueue = {
                operations: [],
                isProcessing: false,
                currentOpId: null,
                stats: {
                    total: 0,
                    pending: 0,
                    processing: 0,
                    success: 0,
                    retry: 0,
                    failed: 0
                }
            };
            DocOpsCore.plogSend('info', 'queue.init', 'Bulletproof queue initialized');
        }
        return global.__bulletproofQueue;
    }
    function updateQueueStats(queue) {
        queue.stats = {
            total: queue.operations.length,
            pending: queue.operations.filter(op => op.status === 'pending').length,
            processing: queue.operations.filter(op => op.status === 'processing').length,
            success: queue.operations.filter(op => op.status === 'success').length,
            retry: queue.operations.filter(op => op.status === 'retry').length,
            failed: queue.operations.filter(op => op.status === 'failed').length
        };
    }
    DocOpsCore.enqueueOperation = function enqueueOperation(rawOp) {
        const queue = initBulletproofQueue();

        const op = {
            id: rawOp.__opId || genOpId(),
            kind: String(rawOp.op || rawOp.kind || rawOp.type || '').toLowerCase(),
            index: rawOp.index,
            jobId: rawOp.jobId || rawOp.job_id || global.__job_id,
            traceId: rawOp.traceId || rawOp.trace_id || global.__trace_id,
            status: 'pending',
            attempt: 0,
            maxAttempts: 5,
            data: cloneOpData(rawOp),
            error: null,
            errors: [],
            createdAt: Date.now(),
            startTime: null,
            lastAttemptTime: null,
            completedAt: null,
            retryDelays: [100, 500, 2000, 5000, 10000],
            nextRetryAt: null
        };

        queue.operations.push(op);
        updateQueueStats(queue);

        DocOpsCore.plogSend('debug', 'queue.enqueue', op.kind + '#' + op.id, {
            queueLength: queue.operations.length,
            pending: queue.stats.pending
        });

        return op.id;
    };
    async function waitWordReady() {
        try {
            const timeout = new Promise(function (resolve) {
                setTimeout(function () { resolve(false); }, 3000);
            });

            const check = Word.run(async function (ctx) {
                const sel = ctx.document.getSelection();
                sel.load("text");
                await ctx.sync();
                return true;
            });

            const result = await Promise.race([check, timeout]);

            if (!result) {
                DocOpsCore.plogSend('warn', 'word.not_ready', 'Word API not responsive within 3s');
            }

            return result;
        } catch (e) {
            DocOpsCore.plogSend('error', 'word.ready_check.error', String(e && e.message || e));
            return false;
        }
    }
    async function getNextInsertionPoint(ctx) {
        let sel = ctx.document.getSelection();

        if (typeof DocOpsCore.pickSafeInsertHost === 'function') {
            sel = await DocOpsCore.pickSafeInsertHost(ctx, sel);
        }

        const emptyPara = sel.insertParagraph("", Word.InsertLocation.after);
        await ctx.sync();

        emptyPara.getRange(Word.RangeLocation.end).select();
        await ctx.sync();

        DocOpsCore.plogSend('debug', 'insert_point.created', 'Empty paragraph created after selection');

        return emptyPara;
    }
    async function moveCaretToEnd(ctx, para) {
        try {
            para.getRange(Word.RangeLocation.end).select();
            await ctx.sync();
            DocOpsCore.plogSend('debug', 'caret.moved', 'Caret moved to end of paragraph');
        } catch (e) {
            DocOpsCore.plogSend('warn', 'caret.move_failed', String(e && e.message || e));
        }
    }
    // ========================================================================
    // OPERATION HANDLERS
    // ========================================================================
    // List state
    const __list = { active: false, items: [], styleName: "Маркированный список" };
    async function flushList() {
        if (!__list.active || !(__list.items && __list.items.length)) {
            __list.active = false;
            __list.items = [];
            return;
        }

        const items = __list.items.slice();
        __list.active = false;
        __list.items = [];

        await Word.run(async function (ctx) {
            const body = ctx.document.body;

            const p = body.insertParagraph(items[0] || "", Word.InsertLocation.end);
            p.startNewList();
            await ctx.sync();

            try {
                p.style = __list.styleName;
            } catch (_) { }
            await ctx.sync();

            for (let i = 1; i < items.length; i++) {
                const pItem = p.insertParagraph(items[i] || "", Word.InsertLocation.after);
                await ctx.sync();

                try {
                    pItem.style = __list.styleName;
                } catch (_) { }
                await ctx.sync();
            }
        });
    }
    async function executeDocxInsert(op) {
        const b64 = String(op.data.base64 || '');
        const loc = String(op.data.location || '');

        if (!b64) {
            throw new Error('No base64 data for docx.insert');
        }

        DocOpsCore.plogSend('info', 'docx.insert.begin', '', { len: b64.length, location: loc });

        if (/^anchor:/i.test(loc)) {
            const anchorText = loc.replace(/^anchor:/i, "").trim();
            if (typeof DocOpsCore.gotoAnchor === 'function') {
                await DocOpsCore.gotoAnchor(anchorText);
            }
        }

        await Word.run(async function (ctx) {
            let sel = ctx.document.getSelection();

            if (typeof DocOpsCore.pickSafeInsertHost === 'function') {
                sel = await DocOpsCore.pickSafeInsertHost(ctx, sel);
            }

            const targetPara = sel.insertParagraph("", Word.InsertLocation.after);
            await ctx.sync();

            DocOpsCore.plogSend('debug', 'docx.target.created', 'Empty paragraph created for DOCX');

            const targetRange = targetPara.getRange();
            let insertedRange;

            try {
                insertedRange = targetRange.insertFileFromBase64(b64, Word.InsertLocation.replace, { importStyles: false });
                await ctx.sync();
                DocOpsCore.plogSend('info', 'docx.insert.ok', 'method: replace');
            } catch (e1) {
                try {
                    insertedRange = targetRange.insertFileFromBase64(b64, Word.InsertLocation.replace);
                    await ctx.sync();
                    DocOpsCore.plogSend('info', 'docx.insert.ok', 'method: replace (no options)');
                } catch (e2) {
                    DocOpsCore.plogSend('error', 'docx.insert.failed', String(e2 && e2.message || e2));
                    throw e2;
                }
            }

            const insertedParas = insertedRange.paragraphs;
            insertedParas.load('items');
            await ctx.sync();

            if (insertedParas.items.length > 0) {
                const lastInsertedPara = insertedParas.items[insertedParas.items.length - 1];

                try {
                    lastInsertedPara.load('parentTable');
                    await ctx.sync();

                    if (lastInsertedPara.parentTable && !lastInsertedPara.parentTable.isNullObject) {
                        const table = lastInsertedPara.parentTable;
                        table.getRange(Word.RangeLocation.after).select();
                        await ctx.sync();
                        DocOpsCore.plogSend('info', 'docx.caret.positioned', 'Caret moved AFTER table');
                    } else {
                        lastInsertedPara.getRange(Word.RangeLocation.end).select();
                        await ctx.sync();
                        DocOpsCore.plogSend('info', 'docx.caret.positioned', 'Caret moved to end of last paragraph');
                    }
                } catch (tableCheckError) {
                    DocOpsCore.plogSend('warn', 'docx.table_check.failed', String(tableCheckError.message || tableCheckError));
                    insertedRange.select(Word.RangeLocation.end);
                    await ctx.sync();
                    DocOpsCore.plogSend('info', 'docx.caret.positioned', 'Caret moved to end of insertedRange (fallback)');
                }
            } else {
                insertedRange.select(Word.RangeLocation.end);
                await ctx.sync();
                DocOpsCore.plogSend('info', 'docx.caret.positioned', 'Caret moved to end of insertedRange (no paragraphs)');
            }
        });

        try {
            await DocOpsCore.flattenForeignContentControls();
            await DocOpsCore.convertEmbeddedMarkersToFootnotes("&SRC&", "last");
            if (typeof DocOpsCore.fixCaptionStylesHard === 'function') {
                await DocOpsCore.fixCaptionStylesHard();
                DocOpsCore.plogSend('info', 'docx.caption_styles.fixed', 'Caption styles corrected');
            }
        } catch (e) {
            DocOpsCore.plogSend('warn', 'docx.post.failed', String(e && e.message || e));
        }
    }
    async function executeCaptionAdd(op) {
        const text = String(op.data.text || op.data.caption || '').trim();

        if (!text) {
            DocOpsCore.plogSend('warn', 'caption.empty', 'No caption text');
            return;
        }

        DocOpsCore.plogSend('info', 'caption.add.begin', text.slice(0, 60));

        await Word.run(async function (ctx) {
            const captionPara = await getNextInsertionPoint(ctx);

            try {
                captionPara.styleBuiltIn = Word.Style.caption;
            } catch (_) {
                try {
                    captionPara.style = 'Название объекта';
                } catch (_) { }
            }

            await ctx.sync();

            const rng = captionPara.getRange(Word.RangeLocation.end);
            rng.insertText("Рис. ", Word.InsertLocation.end);
            await ctx.sync();

            let usedFields = false;
            try {
                if (typeof Word.FieldType !== 'undefined' && Word.FieldType.seq && Word.FieldType.styleRef) {
                    try {
                        rng.insertField(Word.InsertLocation.end, Word.FieldType.styleRef, '"Заголовок 1" \\n', true);
                        await ctx.sync();
                    } catch (_) {
                        rng.insertField(Word.InsertLocation.end, Word.FieldType.styleRef, '"Heading 1" \\n', true);
                        await ctx.sync();
                    }

                    rng.insertText("-", Word.InsertLocation.end);
                    await ctx.sync();

                    rng.insertField(Word.InsertLocation.end, Word.FieldType.seq, 'Рисунок \\* ARABIC \\s 1', true);
                    await ctx.sync();

                    usedFields = true;
                }
            } catch (e) {
                DocOpsCore.plogSend('warn', 'caption.fields.failed', String(e && e.message || e));
            }

            if (!usedFields) {
                const paras = ctx.document.body.paragraphs;
                paras.load("items/text");
                await ctx.sync();

                let maxN = 0;
                for (const para of (paras.items || [])) {
                    const m = String(para.text || "").match(/^Рис\.\s*(\d+)/);
                    if (m) {
                        const n = parseInt(m[1], 10);
                        if (!isNaN(n) && n > maxN) maxN = n;
                    }
                }

                const nextN = maxN + 1;
                rng.insertText(String(nextN), Word.InsertLocation.end);
                await ctx.sync();
            }

            rng.insertText(" " + text, Word.InsertLocation.end);
            await ctx.sync();

            await moveCaretToEnd(ctx, captionPara);
        });

        DocOpsCore.plogSend('info', 'caption.add.done', '');
    }
    async function executeParagraphInsert(op) {
        const text = String(op.data.text || '');

        if (DocOpsCore.isAnchorText(text)) {
            DocOpsCore.plogSend('debug', 'paragraph.is_anchor', 'Moving caret, not inserting text', { text: text.slice(0, 50) });

            if (typeof DocOpsCore.gotoAnchor === 'function') {
                try {
                    const moved = await DocOpsCore.gotoAnchor(text);
                    if (moved) {
                        DocOpsCore.plogSend('info', 'paragraph.anchor.moved', 'Caret moved to anchor');
                    } else {
                        DocOpsCore.plogSend('warn', 'paragraph.anchor.not_found', 'Anchor not found');
                    }
                } catch (e) {
                    DocOpsCore.plogSend('error', 'paragraph.anchor.error', String(e && e.message || e));
                }
            }

            return;
        }

        await Word.run(async function (ctx) {
            const para = await getNextInsertionPoint(ctx);

            para.getRange().insertText(text, Word.InsertLocation.replace);
            await ctx.sync();

            try {
                para.styleBuiltIn = 'Normal';
            } catch (_) { }

            await ctx.sync();

            await moveCaretToEnd(ctx, para);
        });
    }
    async function executeListStart(op) {
        __list.active = true;
        __list.items = [];
        __list.styleName = op.data.styleName || op.data.styleNameHint || 'Маркированный список';

        DocOpsCore.plogSend('debug', 'list.start', __list.styleName);
    }
    async function executeListItem(op) {
        const text = String(op.data.text || '');

        if (!__list.active) {
            __list.active = true;
            __list.items = [];
        }

        __list.items.push(text);
        DocOpsCore.plogSend('debug', 'list.item', 'count=' + __list.items.length);
    }
    async function executeListEnd(op) {
        await flushList();
        DocOpsCore.plogSend('debug', 'list.end', '');
    }
    async function executeFootnotes(op) {
        const marker = String(op.data.marker || '&SRC&');
        const scope = (op.data.scope === 'doc') ? 'doc' : 'last';

        DocOpsCore.plogSend('info', 'footnotes.begin', '', { marker, scope });

        if (typeof DocOpsCore.convertEmbeddedMarkersToFootnotes === 'function') {
            const applied = await DocOpsCore.convertEmbeddedMarkersToFootnotes(marker, scope);
            DocOpsCore.plogSend('info', 'footnotes.applied', '', { applied });
        } else {
            DocOpsCore.plogSend('warn', 'footnotes.handler.missing', '');
        }
    }
    async function executeJobTail(op) {
        DocOpsCore.plogSend('info', 'job.tail.begin', '');

        await Word.run(async function (ctx) {
            const marker = `<BLOCK: ${op.data.jobId}>`;
            const sel = ctx.document.getSelection();

            const blockPara = sel.insertParagraph(marker, Word.InsertLocation.after);

            try {
                blockPara.style = 'Normal';
            } catch (e) { }

            blockPara.select(Word.RangeLocation.end);
            await ctx.sync();

            DocOpsCore.plogSend('info', 'job.marker.appended', '');
        });

        DocOpsCore.plogSend('info', 'job.tail.end', '');
    }
    async function executeOperationByKind(op) {
        switch (op.kind) {
            case 'docx.insert':
                await executeDocxInsert(op);
                break;
            case 'caption.add':
                await executeCaptionAdd(op);
                break;
            case 'paragraph.insert':
            case 'paragraph':
                await executeParagraphInsert(op);
                break;
            case 'list.start':
                await executeListStart(op);
                break;
            case 'list.item':
                await executeListItem(op);
                break;
            case 'list.end':
                await executeListEnd(op);
                break;
            case 'footnotes.apply-embedded':
                await executeFootnotes(op);
                break;
            case 'job.marker.tail':
            case 'job.tail':
            case 'block.end':
                await executeJobTail(op);
                break;
            default:
                DocOpsCore.plogSend('warn', 'op.unknown_kind', op.kind);
                throw new Error('Unknown operation kind: ' + op.kind);
        }
    }
    async function safeExecuteOperation(op) {
        let lastStep = 'init';

        try {
            lastStep = 'validate';
            if (!op.data || !op.kind) {
                throw new Error('Invalid operation data');
            }

            lastStep = 'execute';
            await executeOperationByKind(op);

            lastStep = 'complete';
            return { success: true };
        } catch (e) {
            DocOpsCore.plogSend('error', 'op.execute.error', op.kind + '#' + op.id, {
                step: lastStep,
                error: String(e && e.message || e)
            });

            return { success: false, error: e, step: lastStep };
        }
    }
    async function processSingleOperation(op) {
        const queue = global.__bulletproofQueue;

        op.status = 'processing';
        op.attempt++;
        op.lastAttemptTime = Date.now();
        if (!op.startTime) op.startTime = Date.now();

        queue.currentOpId = op.id;

        DocOpsCore.plogSend('info', 'op.process.start', op.kind + '#' + op.id, {
            attempt: op.attempt,
            maxAttempts: op.maxAttempts
        });

        try {
            const ready = await waitWordReady();
            if (!ready) {
                throw new Error('Word API not ready');
            }

            const result = await safeExecuteOperation(op);

            if (result.success) {
                op.status = 'success';
                op.completedAt = Date.now();
                op.error = null;

                DocOpsCore.plogSend('info', 'op.process.success', op.kind + '#' + op.id, {
                    attempt: op.attempt,
                    duration: op.completedAt - (op.startTime || op.createdAt)
                });
            } else {
                handleOperationError(op, result.error, result.step);
            }
        } catch (e) {
            handleOperationError(op, e, 'unknown');
        }
    }
    function handleOperationError(op, error, step) {
        const errorMsg = String(error && error.message || error || 'Unknown error');

        op.errors.push({
            attempt: op.attempt,
            error: errorMsg,
            timestamp: Date.now(),
            step: step
        });

        if (op.attempt < op.maxAttempts) {
            op.status = 'retry';
            op.error = errorMsg;

            const delay = op.retryDelays[op.attempt - 1] || 10000;
            op.nextRetryAt = Date.now() + delay;

            DocOpsCore.plogSend('warn', 'op.process.retry', op.kind + '#' + op.id, {
                attempt: op.attempt,
                maxAttempts: op.maxAttempts,
                nextRetryIn: delay,
                error: errorMsg,
                step: step
            });
        } else {
            op.status = 'failed';
            op.error = errorMsg;
            op.completedAt = Date.now();

            DocOpsCore.plogSend('error', 'op.process.failed', op.kind + '#' + op.id, {
                attempts: op.attempt,
                errors: op.errors,
                finalError: errorMsg
            });
        }
    }
    DocOpsCore.processQueue = async function processQueue() {
        const queue = initBulletproofQueue();

        if (queue.isProcessing) {
            DocOpsCore.plogSend('debug', 'queue.already_processing', '');
            return;
        }

        queue.isProcessing = true;

        DocOpsCore.plogSend('info', 'queue.process.start', '', { total: queue.operations.length });

        try {
            while (true) {
                const nextOp = findNextOperation(queue);
                if (!nextOp) {
                    DocOpsCore.plogSend('info', 'queue.process.complete', 'No more operations');
                    break;
                }

                if (nextOp.nextRetryAt && Date.now() < nextOp.nextRetryAt) {
                    const waitMs = nextOp.nextRetryAt - Date.now();
                    DocOpsCore.plogSend('debug', 'queue.wait_retry', nextOp.kind + '#' + nextOp.id, { waitMs });
                    await new Promise(function (r) { setTimeout(r, waitMs); });
                }

                await processSingleOperation(nextOp);
                updateQueueStats(queue);
            }
        } finally {
            queue.isProcessing = false;
            queue.currentOpId = null;

            DocOpsCore.plogSend('info', 'queue.process.end', '', queue.stats);
        }
    };
    function findNextOperation(queue) {
        const now = Date.now();

        for (let i = 0; i < queue.operations.length; i++) {
            const op = queue.operations[i];

            if (op.status === 'success' || op.status === 'failed') continue;
            if (op.status === 'processing') continue;
            if (op.status === 'retry' && op.nextRetryAt && now < op.nextRetryAt) continue;

            return op;
        }

        return null;
    }
    // ========================================================================
    // DEBUG UTILITIES
    // ========================================================================
    DocOpsCore.getQueueStats = function getQueueStats() {
        const queue = global.__bulletproofQueue;
        if (!queue) return null;

        return {
            stats: queue.stats,
            isProcessing: queue.isProcessing,
            currentOpId: queue.currentOpId,
            operations: queue.operations.map(function (op) {
                return {
                    id: op.id,
                    kind: op.kind,
                    status: op.status,
                    attempt: op.attempt,
                    error: op.error
                };
            })
        };
    };
    DocOpsCore.debugQueue = function debugQueue() {
        console.table((DocOpsCore.getQueueStats() || {}).operations || []);
    };
    // ========================================================================
    // EXPORTS
    // ========================================================================
    // Expose to global scope
    global.DocOpsCore = DocOpsCore;
    // Expose compatibility aliases
    global.enqueueOperation = DocOpsCore.enqueueOperation;
    global.processQueue = DocOpsCore.processQueue;
    global.getQueueStats = DocOpsCore.getQueueStats;
    global.debugQueue = DocOpsCore.debugQueue;
    global.plogSend = DocOpsCore.plogSend;
    global.pickSafeInsertHost = DocOpsCore.pickSafeInsertHost;
    console.log("[docops-core] DocOps Core Engine loaded");
})(typeof window !== 'undefined' ? window : this);