/**
 * Overall-size measure for Studio Delta QC.
 * Default: bright tape/manual entry (never a black screen).
 * Optional WebXR AR on Android Chrome + ARCore only — iPhone skips AR
 * because Safari often opens a black camera session.
 */
(function (global) {
  "use strict";

  const OVERLAY_ID = "sdArMeasureMask";

  function roundMm(meters) {
    if (!Number.isFinite(meters) || meters < 0) return 0;
    const mm = Math.round(meters * 1000);
    if (mm <= 0) return 0;
    return mm;
  }

  function distMeters(a, b) {
    const dx = a.x - b.x;
    const dy = a.y - b.y;
    const dz = a.z - b.z;
    return Math.sqrt(dx * dx + dy * dy + dz * dz);
  }

  function isAppleMobile() {
    const ua = String((global.navigator && navigator.userAgent) || "");
    return /iPhone|iPad|iPod/i.test(ua) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  }

  async function isArMeasureSupported() {
    try {
      if (isAppleMobile()) return { ok: false, reason: "ios" };
      if (!global.isSecureContext) return { ok: false, reason: "https" };
      if (!navigator.xr || typeof navigator.xr.isSessionSupported !== "function") {
        return { ok: false, reason: "no-xr" };
      }
      const ok = await navigator.xr.isSessionSupported("immersive-ar");
      return ok ? { ok: true, reason: "webxr" } : { ok: false, reason: "no-immersive-ar" };
    } catch (e) {
      return { ok: false, reason: "error" };
    }
  }

  function ensureOverlay() {
    let el = document.getElementById(OVERLAY_ID);
    if (el) return el;
    el = document.createElement("div");
    el.id = OVERLAY_ID;
    el.className = "sd-ar-mask";
    el.hidden = true;
    el.innerHTML =
      '<div class="sd-ar-sheet" role="dialog" aria-modal="true" aria-labelledby="sdArTitle">' +
      '  <header class="sd-ar-head">' +
      '    <div><h2 id="sdArTitle">Measure overall size</h2>' +
      '    <p class="sd-ar-hint" id="sdArHint">Type the tape size in millimetres.</p></div>' +
      '    <button type="button" class="sd-ar-x" id="sdArClose" aria-label="Close">Close</button>' +
      "  </header>" +
      '  <div class="sd-ar-body" id="sdArBody">' +
      '    <p class="sd-ar-lead" id="sdArFallbackText"></p>' +
      '    <label class="sd-ar-label">Millimetres' +
      '      <input id="sdArManual" inputmode="numeric" type="number" min="1" step="1" placeholder="e.g. 1200">' +
      "    </label>" +
      '    <button type="button" class="sd-ar-primary" id="sdArManualUse">Use this size</button>' +
      '    <button type="button" class="sd-ar-secondary" id="sdArTryCam" hidden>Try AR camera</button>' +
      '    <p class="sd-ar-note" id="sdArNote"></p>' +
      "  </div>" +
      '  <div class="sd-ar-stage" id="sdArStage" hidden>' +
      '    <canvas id="sdArCanvas"></canvas>' +
      '    <div class="sd-ar-readout" id="sdArReadout">—</div>' +
      '    <p class="sd-ar-status" id="sdArStatus"></p>' +
      '    <div class="sd-ar-foot">' +
      '      <button type="button" class="sd-ar-ghost" id="sdArBack">Back</button>' +
      '      <button type="button" class="sd-ar-ghost" id="sdArReset">Reset</button>' +
      '      <button type="button" class="sd-ar-primary" id="sdArUse" disabled>Use measurement</button>' +
      "    </div>" +
      "  </div>" +
      "</div>";
    document.body.appendChild(el);
    if (!document.getElementById("sdArMeasureStyles")) {
      const style = document.createElement("style");
      style.id = "sdArMeasureStyles";
      style.textContent =
        ".sd-ar-mask{position:fixed;inset:0;z-index:2000;background:#f8fafc;color:#1d2939;display:flex;flex-direction:column}" +
        ".sd-ar-mask[hidden]{display:none!important}" +
        ".sd-ar-sheet{flex:1;display:flex;flex-direction:column;min-height:0;max-width:560px;width:100%;margin:0 auto}" +
        ".sd-ar-head{display:flex;gap:12px;align-items:flex-start;padding:16px;border-bottom:1px solid #d0d5dd;background:#fff}" +
        ".sd-ar-head h2{margin:0;font-size:18px;font-weight:700}" +
        ".sd-ar-hint{margin:4px 0 0;color:#667085;font-size:13px}" +
        ".sd-ar-x{margin-left:auto;border:1px solid #d0d5dd;background:#fff;color:#1d2939;border-radius:8px;padding:10px 14px;font-weight:600;min-height:44px}" +
        ".sd-ar-body{padding:20px 16px;display:flex;flex-direction:column;gap:14px;background:#f8fafc;flex:1}" +
        ".sd-ar-body[hidden]{display:none!important}" +
        ".sd-ar-lead{margin:0;font-size:15px;line-height:1.45;color:#344054}" +
        ".sd-ar-label{display:flex;flex-direction:column;gap:6px;font-size:12px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;color:#667085}" +
        ".sd-ar-label input{min-height:52px;border-radius:10px;border:1px solid #d0d5dd;background:#fff;color:#1d2939;padding:12px;font-size:18px}" +
        ".sd-ar-primary,.sd-ar-secondary,.sd-ar-ghost{min-height:48px;border-radius:10px;font-weight:700;font-size:15px;padding:12px 16px}" +
        ".sd-ar-primary{border:0;background:#1d2939;color:#fff}" +
        ".sd-ar-primary:disabled{opacity:.45}" +
        ".sd-ar-secondary{border:1px solid #1d2939;background:#fff;color:#1d2939}" +
        ".sd-ar-secondary[hidden]{display:none!important}" +
        ".sd-ar-note{margin:0;font-size:12px;color:#667085;line-height:1.4}" +
        ".sd-ar-stage{position:relative;flex:1;min-height:280px;background:#111;color:#fff;display:flex;flex-direction:column}" +
        ".sd-ar-stage[hidden]{display:none!important}" +
        ".sd-ar-stage canvas{position:absolute;inset:0 0 72px 0;width:100%;height:calc(100% - 72px);touch-action:none}" +
        ".sd-ar-readout{position:absolute;left:50%;top:16%;transform:translate(-50%,-50%);font-size:40px;font-weight:800;text-shadow:0 2px 12px #000;pointer-events:none}" +
        ".sd-ar-status{position:absolute;left:16px;right:16px;bottom:84px;margin:0;font-size:14px;text-shadow:0 1px 8px #000;pointer-events:none}" +
        ".sd-ar-foot{margin-top:auto;display:flex;gap:8px;padding:12px 16px calc(12px + env(safe-area-inset-bottom,0px));background:#1c1917;position:relative;z-index:2}" +
        ".sd-ar-foot button{flex:1}" +
        ".sd-ar-ghost{border:1px solid #57534e;background:transparent;color:#fafaf9}";
      document.head.appendChild(style);
    }
    return el;
  }

  function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
  }

  function showManual(mask, label, support) {
    const body = document.getElementById("sdArBody");
    const stage = document.getElementById("sdArStage");
    const tryCam = document.getElementById("sdArTryCam");
    if (body) body.hidden = false;
    if (stage) stage.hidden = true;
    setText("sdArTitle", "Measure " + label);
    setText("sdArHint", "Type the tape reading in millimetres.");
    let why = "Use a tape measure, type the size, then Use this size.";
    if (support && support.reason === "ios") {
      why = "iPhone cannot run this in-app AR measure reliably (black camera). Use a tape, type millimetres below.";
    } else if (support && support.reason === "https") {
      why = "AR camera needs https. This preview may be http on your phone — use a tape and type millimetres below.";
    } else if (support && !support.ok) {
      why = "AR camera is not available on this phone/browser. Use a tape and type millimetres below.";
    }
    setText("sdArFallbackText", why);
    setText("sdArNote", "AR is optional. Tape is the source of truth when sizes are tight.");
    if (tryCam) tryCam.hidden = !(support && support.ok);
  }

  async function openArMeasure(opts) {
    const options = opts || {};
    const label = String(options.label || "edge").trim() || "edge";
    const onResult = typeof options.onResult === "function" ? options.onResult : null;
    const mask = ensureOverlay();
    const canvas = document.getElementById("sdArCanvas");
    const useBtn = document.getElementById("sdArUse");
    const resetBtn = document.getElementById("sdArReset");
    const closeBtn = document.getElementById("sdArClose");
    const backBtn = document.getElementById("sdArBack");
    const tryCam = document.getElementById("sdArTryCam");
    const body = document.getElementById("sdArBody");
    const stage = document.getElementById("sdArStage");
    const manual = document.getElementById("sdArManual");
    const manualUse = document.getElementById("sdArManualUse");

    let session = null;
    let gl = null;
    let refSpace = null;
    let viewerSpace = null;
    let hitTestSource = null;
    let points = [];
    let latestMm = 0;
    let closed = false;

    function cleanup() {
      closed = true;
      try {
        if (hitTestSource && hitTestSource.cancel) hitTestSource.cancel();
      } catch (e) {}
      hitTestSource = null;
      try {
        if (session) session.end();
      } catch (e) {}
      session = null;
      mask.hidden = true;
    }

    function finishMm(mm) {
      cleanup();
      if (onResult) onResult(Math.max(1, Math.round(Number(mm) || 0)));
    }

    function updateReadout() {
      if (latestMm > 0) {
        setText("sdArReadout", latestMm + " mm");
        useBtn.disabled = false;
      } else {
        setText("sdArReadout", "—");
        useBtn.disabled = true;
      }
    }

    async function stopArAndShowManual(support) {
      try {
        if (hitTestSource && hitTestSource.cancel) hitTestSource.cancel();
      } catch (e) {}
      hitTestSource = null;
      try {
        if (session) await session.end();
      } catch (e) {}
      session = null;
      showManual(mask, label, support || { ok: false, reason: "fallback" });
    }

    async function startArCamera() {
      if (body) body.hidden = true;
      if (stage) stage.hidden = false;
      setText("sdArStatus", "Starting camera…");
      setText("sdArReadout", "—");
      latestMm = 0;
      points = [];
      useBtn.disabled = true;

      try {
        gl = canvas.getContext("webgl", { xrCompatible: true, alpha: true });
        if (!gl) throw new Error("WebGL not available");

        // Prefer session without DOM overlay — overlay often paints a black sheet over the camera.
        session = await navigator.xr.requestSession("immersive-ar", {
          requiredFeatures: ["hit-test"],
          optionalFeatures: ["local", "local-floor", "dom-overlay"],
          domOverlay: { root: mask }
        }).catch(() =>
          navigator.xr.requestSession("immersive-ar", {
            requiredFeatures: ["hit-test"],
            optionalFeatures: ["local", "local-floor"]
          })
        );

        const layer = new XRWebGLLayer(session, gl, { alpha: true });
        await session.updateRenderState({ baseLayer: layer });
        refSpace = await session.requestReferenceSpace("local").catch(() => session.requestReferenceSpace("local-floor"));
        viewerSpace = await session.requestReferenceSpace("viewer");
        if (!session.requestHitTestSource) throw new Error("Hit-test unavailable");
        hitTestSource = await session.requestHitTestSource({ space: viewerSpace });

        setText("sdArStatus", "Tap the first end of the " + label + ".");

        session.addEventListener("end", () => {
          if (!closed && stage && !stage.hidden) {
            showManual(mask, label, { ok: false, reason: "ended" });
          }
        });

        session.addEventListener("select", (ev) => {
          try {
            const frame = ev.frame;
            if (!frame || !hitTestSource) return;
            const hits = frame.getHitTestResults(hitTestSource);
            if (!hits || !hits.length) {
              setText("sdArStatus", "No surface found — aim at the item and tap again.");
              return;
            }
            const pose = hits[0].getPose(refSpace);
            if (!pose) return;
            const p = {
              x: pose.transform.position.x,
              y: pose.transform.position.y,
              z: pose.transform.position.z
            };
            if (points.length >= 2) points = [];
            points.push(p);
            if (points.length === 1) {
              latestMm = 0;
              updateReadout();
              setText("sdArStatus", "First point set. Tap the other end.");
              return;
            }
            latestMm = roundMm(distMeters(points[0], points[1]));
            updateReadout();
            setText("sdArStatus", "Measured " + latestMm + " mm. Use measurement, or Reset.");
          } catch (err) {
            setText("sdArStatus", "Could not place that point. Try again.");
          }
        });

        let frames = 0;
        const onXRFrame = (_time, frame) => {
          if (closed || !session) return;
          session.requestAnimationFrame(onXRFrame);
          frames += 1;
          const base = session.renderState.baseLayer;
          if (!base) return;
          gl.bindFramebuffer(gl.FRAMEBUFFER, base.framebuffer);
          // Transparent clear so the XR camera feed stays visible.
          gl.clearColor(0, 0, 0, 0);
          gl.clear(gl.COLOR_BUFFER_BIT | gl.DEPTH_BUFFER_BIT);
          if (points.length === 1 && hitTestSource) {
            const hits = frame.getHitTestResults(hitTestSource);
            if (hits && hits.length) {
              const pose = hits[0].getPose(refSpace);
              if (pose) {
                const live = roundMm(
                  distMeters(points[0], {
                    x: pose.transform.position.x,
                    y: pose.transform.position.y,
                    z: pose.transform.position.z
                  })
                );
                setText("sdArReadout", live + " mm");
              }
            }
          }
        };
        session.requestAnimationFrame(onXRFrame);

        // If still black / no useful frames, bounce to manual after a short wait.
        setTimeout(() => {
          if (closed || !session) return;
          if (frames < 5) {
            stopArAndShowManual({ ok: false, reason: "no-frames" });
            setText("sdArFallbackText", "AR camera did not start on this phone. Type the tape size in millimetres.");
          }
        }, 2500);

        return true;
      } catch (e) {
        await stopArAndShowManual({ ok: false, reason: "error" });
        setText(
          "sdArFallbackText",
          "Could not start AR (" + ((e && e.message) || "unsupported") + "). Type the tape size below."
        );
        return false;
      }
    }

    mask.hidden = false;
    if (manual) manual.value = "";
    latestMm = 0;
    points = [];

    closeBtn.onclick = () => cleanup();
    backBtn.onclick = () => stopArAndShowManual(lastSupport);
    resetBtn.onclick = () => {
      points = [];
      latestMm = 0;
      updateReadout();
      setText("sdArStatus", "Tap the first end of the " + label + ".");
    };
    useBtn.onclick = () => {
      if (latestMm > 0) finishMm(latestMm);
    };
    manualUse.onclick = () => {
      const n = Number(manual && manual.value);
      if (!Number.isFinite(n) || n < 1) {
        setText("sdArFallbackText", "Type a size in millimetres (for example 1200).");
        return;
      }
      finishMm(n);
    };

    const lastSupport = await isArMeasureSupported();
    showManual(mask, label, lastSupport);
    tryCam.onclick = () => startArCamera();

    // Auto-focus the input so the phone keyboard is ready.
    setTimeout(() => {
      try {
        if (manual) manual.focus();
      } catch (e) {}
    }, 50);

    return { mode: lastSupport.ok ? "manual-with-ar" : "manual" };
  }

  global.sdArMeasureSupported = isArMeasureSupported;
  global.sdOpenArMeasure = openArMeasure;
  global.sdArRoundMm = roundMm;
  global.sdArDistMeters = distMeters;
})(typeof window !== "undefined" ? window : globalThis);
