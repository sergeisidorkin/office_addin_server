/** C:\inetpub\wwwroot\addin\selftest.js **/
(function(){
  'use strict';

  // ========================================================================
  // LOG
  // ========================================================================
  var LOG = null;
  function log(s){
    try { console.log(s); } catch(_) {}
    if(!LOG) { LOG = document.getElementById("log"); }
    if(!LOG){
      LOG = document.createElement("pre");
      LOG.id = "log";
      LOG.style.cssText = "font:12px/1.4 monospace;white-space:pre-wrap;background:#111;color:#0f0;padding:8px;max-height:50vh;overflow:auto;border:1px solid #333;";
      document.body.appendChild(LOG);
    }
    LOG.textContent += "["+new Date().toISOString()+"] "+s+"\n";
    LOG.scrollTop = LOG.scrollHeight;
  }

  // ========================================================================
  // QUERY / CONFIG LOADER
  // ========================================================================
  function getQuery(){
    var q = {}, src = location.search || "";
    if(src.charAt(0) === "?") src = src.substring(1);
    src.split("&").forEach(function(p){
      if(!p) return;
      var kv = p.split("=");
      q[decodeURIComponent(kv[0] || "")] = decodeURIComponent(kv[1] || "");
    });
    return q;
  }
  var Q = getQuery();

  function injectScriptOnce(src){
    return new Promise(function(resolve){
      try{
        // Уже есть такой тег?
        var exists = Array.from(document.scripts || []).some(function(s){ return (s.src||"").indexOf(src) >= 0; });
        if(exists){ resolve(true); return; }

        var el = document.createElement("script");
        el.src = src;
        el.onload = function(){ resolve(true); };
        el.onerror = function(){ resolve(false); };
        document.head.appendChild(el);
      }catch(_){ resolve(false); }
    });
  }

  async function ensureConfigLoaded(){
    // Если уже есть глобальная переменная  ничего не делаем
    if(typeof window.DOCOPS_BASE === "string" || typeof window.DOCOPS_AGENT === "string"){
      return true;
    }
    // Пытаемся загрузить конфиг с кэш-бастером
    var ok = await injectScriptOnce("/addin/config.js?v=" + Date.now());
    if(!ok){
      log("WARN: config.js not loaded (will rely on autodetect/fallback)");
    }
    return ok;
  }

  // ========================================================================
  // BASE RESOLUTION
  // ========================================================================
  var _BASE = null;
  function autoDetectBaseByDocUrl(){
    try{
      var u = (Office && Office.context && Office.context.document && Office.context.document.url) || "";
      if(/imcmontanai-my\.sharepoint\.com/i.test(u)) return "https://imcmontanai.ru"; // прод
      // можно расширить правила при необходимости
    }catch(_){}
    return "";
  }
  function resolveBase(){
    var chosen = null, source = null;

    // 1) query ?base=
    var qBase = (Q.base || "").trim();
    if(qBase){
      chosen = qBase; source = "query";
    }

    // 2) window.DOCOPS_BASE из config.js
    if(!chosen && typeof window.DOCOPS_BASE === "string" && window.DOCOPS_BASE.trim()){
      chosen = window.DOCOPS_BASE.trim(); source = "config.js";
    }

    // 3) авто-детект по URL документа (тенант)
    if(!chosen){
      var ad = autoDetectBaseByDocUrl();
      if(ad){ chosen = ad; source = "autodetect"; }
    }

    // 4) дефолт  cloudpub
    if(!chosen){
      chosen = "https://snappishly-primed-blackfish.cloudpub.ru";
      source = "default";
    }

    _BASE = chosen;
    log("BASE resolved: " + chosen + " (source=" + source + ")");
    return _BASE;
  }
  function getBASE(){ return _BASE || resolveBase(); }

  // ========================================================================
  // HTTP
  // ========================================================================
  var HDRS_JSON = {"ngrok-skip-browser-warning":"1","Content-Type":"application/json"};

  function postJSON(base, path, body){
    return new Promise(function(res, rej){
      var x = new XMLHttpRequest();
      x.open("POST", base + path, true);
      for(var k in HDRS_JSON){ x.setRequestHeader(k, HDRS_JSON[k]); }
      x.onreadystatechange = function(){
        if(x.readyState === 4){
          if(x.status >= 200 && x.status < 300){
            var d = null;
            try { d = x.responseText ? JSON.parse(x.responseText) : null; } catch(e) {}
            res({ ok:true, status:x.status, data:d, text:x.responseText });
          } else {
            rej(new Error("HTTP "+x.status+" "+base+path));
          }
        }
      };
      x.onerror = function(){ rej(new Error("network "+base+path)); };
      x.send(JSON.stringify(body || {}));
    });
  }

  async function postWithFallback(primaryBase, altBase, path, body, tag){
    try{
      log("HTTP  " + tag + " @ " + primaryBase + path);
      return await postJSON(primaryBase, path, body);
    }catch(e1){
      log("WARN: primary failed ("+e1+"), trying fallback @ " + altBase + path);
      try{
        var r2 = await postJSON(altBase, path, body);
        // если фолбэк сработал  переключаем BASE
        _BASE = altBase;
        log("BASE switched to fallback: " + _BASE + " (after " + tag + ")");
        return r2;
      }catch(e2){
        throw new Error(tag + " failed both: primary=" + e1 + " ; fallback=" + e2);
      }
    }
  }

  // ========================================================================
  // MARKER / WORD
  // ========================================================================
  var MARKER_PREFIX = "<BLOCK:";
  var MARKER_SUFFIX = ">";
  function makeDoneMarker(jobId){ return MARKER_PREFIX + " " + jobId + MARKER_SUFFIX; }

  function hasMarker(jobId){
    var needle = makeDoneMarker(jobId);
    return Word.run(function(ctx){
      var r = ctx.document.body.getRange("Whole");
      r.load("text");
      return ctx.sync().then(function(){
        var txt = r.text || "";
        return txt.indexOf(needle) >= 0;
      });
    }).catch(function(){ return false; });
  }

  function tryBridgeSync(){
    try {
      if(window.chrome && chrome.webview && chrome.webview.hostObjects && chrome.webview.hostObjects.sync){
        chrome.webview.hostObjects.sync.PutUpdate();
        log("Sync.PutUpdate() called");
        return true;
      }
    } catch(e) { log("PutUpdate error: "+e); }
    return false;
  }

  // ========================================================================
  // DOC-CENTRIC JOB PULLING
  // ========================================================================
  var inflight = false, loopStarted = false;

  function getDocUrl(){
    try { return (Office && Office.context && Office.context.document && Office.context.document.url) || ""; }
    catch(_){ return ""; }
  }

  async function claimNextForThisDoc(){
    var url = getDocUrl();
    if(!url){
      log("No document URL yet, skip tick");
      return null;
    }

    var base = getBASE();
    var alt  = (base.indexOf("imcmontanai.ru") >= 0)
      ? "https://snappishly-primed-blackfish.cloudpub.ru"
      : "https://imcmontanai.ru";

    var r = await postWithFallback(base, alt, "/api/docs/next", { url:url }, "docs/next");
    var job = (r && r.data && r.data.job) ? r.data.job : null;
    if(job){
      log("docs/next  job " + job.id + " (trace=" + (job.traceId || "") + ")");
    }else{
      log("docs/next  no job");
    }
    return job;
  }

  async function completeJob(jobId, ok, message){
    var base = getBASE();
    var alt  = (base.indexOf("imcmontanai.ru") >= 0)
      ? "https://snappishly-primed-blackfish.cloudpub.ru"
      : "https://imcmontanai.ru";

    try {
      await postWithFallback(base, alt, "/api/jobs/"+jobId+"/complete", { ok: !!ok, message: message || "" }, "jobs/complete");
    } catch(e) {
      log("complete error: " + e);
    }
  }

  // ========================================================================
  // PROCESS (DocOps Core)
  // ========================================================================
  function processOnce(){
    if(inflight) return Promise.resolve(false);
    inflight = true;

    return claimNextForThisDoc()
      .then(function(job){
        if(!job){
          inflight = false;
          return false;
        }

        var jobId   = job.id;
        var traceId = job.traceId || "";
        var payload = job.payload || {};

        log("Job claimed: id=" + jobId);

        return hasMarker(jobId).then(function(already){
          if(already){
            log("Marker present  complete (already done)");
            return completeJob(jobId, true, "already-present").then(function(){
              inflight = false;
              return true;
            });
          }

          // Проверка формата
          var isAddinBlock = (String(payload.type || "").toLowerCase() === "addin.block");
          if(!isAddinBlock){
            log("WARN: payload is not addin.block, type=" + (payload.type || "unknown"));
            return completeJob(jobId, false, "invalid payload type").then(function(){
              inflight = false; return false;
            });
          }

          var blocks = payload.blocks || [];
          if(!blocks.length){
            log("WARN: empty blocks array");
            return completeJob(jobId, false, "empty blocks").then(function(){
              inflight = false; return false;
            });
          }

          // DocOps Core?
          if(typeof window.DocOpsCore === "undefined" || typeof window.enqueueOperation !== "function"){
            log("ERROR: DocOps Core not loaded!");
            return completeJob(jobId, false, "DocOps Core not available").then(function(){
              inflight = false; return false;
            });
          }

          try { window.__job_id = jobId; window.__trace_id = traceId; } catch(_){}

          var enqueued = 0;
          for(var i=0;i<blocks.length;i++){
            try{
              blocks[i].jobId = jobId;
              blocks[i].traceId = traceId;
              blocks[i].index = i;
              window.enqueueOperation(blocks[i]);
              enqueued++;
            }catch(e){ log("ERROR enqueue #" + i + ": " + e); }
          }
          log("Enqueued " + enqueued + " operations, starting processing...");

          return window.processQueue()
            .then(function(){
              log("Queue processing completed");
              tryBridgeSync();

              try {
                fetch("http://127.0.0.1:17603/done?jobId=" + encodeURIComponent(jobId), { method:"GET", mode:"no-cors" }).catch(function(){});
              } catch(_) {}

              return completeJob(jobId, true, "processed via DocOps Core");
            })
            .catch(function(e){
              log("ERROR processing queue: " + (e && e.message || e));
              return completeJob(jobId, false, "queue processing failed: " + e);
            })
            .then(function(){
              inflight = false;
              return true;
            });
        });
      })
      .catch(function(e){
        log("process error: " + (e && e.message || e));
        inflight = false;
        return false;
      });
  }

  // ========================================================================
  // BACKGROUND LOOP
  // ========================================================================
  function backgroundLoop(){
    if(loopStarted) return;
    loopStarted = true;

    var idle = 0;
    function tick(){
      processOnce().then(function(done){
        idle = done ? 0 : Math.min(idle + 1, 4);
        var delay = done ? 400 : (idle < 2 ? 600 : 1200);
        setTimeout(tick, delay);
      });
    }
    setTimeout(tick, 150);
  }

  // ========================================================================
  // BOOTSTRAP
  // ========================================================================
  async function start(){
    await ensureConfigLoaded(); // подтянуть window.DOCOPS_BASE / DOCOPS_AGENT если есть
    var base = resolveBase();

    log("Panel ready (DocOps Core integrated). base=" + base);

    if(typeof window.DocOpsCore !== "undefined"){
      log(" DocOps Core loaded");
      if(window.DocOpsCore.config){
        window.DocOpsCore.config.logsBase = base;
        window.DocOpsCore.config.logsToken = "";
      }
    } else {
      log(" WARNING: DocOps Core NOT loaded!");
    }

    backgroundLoop();
  }

  if(window.Office && Office.onReady){
    Office.onReady().then(function(){ start(); });
  } else {
    setTimeout(start, 800);
  }

  // ========================================================================
  // TEST BUTTONS (optional)
  // ========================================================================
  document.addEventListener("DOMContentLoaded", function(){
    var btnInsert = document.getElementById("btn-insert");
    if(btnInsert){
      btnInsert.addEventListener("click", function(){
        Word.run(function(ctx){
          ctx.document.body.insertParagraph("SELFTEST from button", Word.InsertLocation.end);
          return ctx.sync();
        }).then(function(){ log("Test paragraph inserted"); })
          .catch(function(e){ log("Error: " + (e && e.message || e)); });
      });
    }

    var btnPull = document.getElementById("btn-pull");
    if(btnPull){
      btnPull.addEventListener("click", function(){
        log("Manual pull triggered...");
        processOnce().then(function(done){
          log(done ? "Manual pull: job processed" : "Manual pull: no job");
        });
      });
    }
  });

})();