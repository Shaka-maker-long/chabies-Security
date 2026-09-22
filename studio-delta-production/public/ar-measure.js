/**
 * Phone AR edge measure for Studio Delta QC (WebXR hit-test).
 * Tap two points on a real edge → length in millimetres.
 * Works best on Android Chrome with ARCore. iPhone Safari support is limited;
 * unsupported phones get a clear tape/manual fallback.
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

  async function isArMeasureSupported() {
    try {
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
      '    <div><h2 id="sdArTitle">Measure with phone</h2>' +
      '    <p class="sd-ar-hint" id="sdArHint">Tap two ends of the edge you are measuring.</p></div>' +
      '    <button type="button" class="sd-ar-x" id="sdArClose" aria-label="Close">Close</button>' +
      "  </header>" +
      '  <div class="sd-ar-stage">' +
      '    <canvas id="sdArCanvas"></canvas>' +
      '    <div class="sd-ar-readout" id="sdArReadout">—</div>' +
      '    <p class="sd-ar-status" id="sdArStatus"></p>' +
      "  </div>" +
      '  <div class="sd-ar-foot">' +
      '    <button type="button" class="sd-ar-ghost" id="sdArReset">Reset points</button>' +
      '    <button type="button" class="sd-ar-primary" id="sdArUse" disabled>Use measurement</button>' +
      "  </div>" +
      '  <div class="sd-ar-fallback" id="sdArFallback" hidden>' +
      "    <p id=\"sdArFallbackText\"></p>" +
      '    <label>Type millimetres' +
      '      <input id="sdArManual" inputmode="numeric" type="number" min="1" step="1" placeholder="e.g. 1200">' +
      "    </label>" +
      '    <button type="button" class="sd-ar-primary" id="sdArManualUse">Use typed size</button>' +
      "  </div>" +
      "</div>";
    document.body.appendChild(el);
    if (!document.getElementById("sdArMeasureStyles")) {
      const style = document.createElement("style");
      style.id = "sdArMeasureStyles";
      style.textContent =
        ".sd-ar-mask{position:fixed;inset:0;z-index:2000;background:#0c0a09;color:#fafaf9;display:flex;flex-direction:column}" +
        ".sd-ar-mask[hidden]{display:none!important}" +
        ".sd-ar-sheet{flex:1;display:flex;flex-direction:column;min-height:0}" +
        ".sd-ar-head{display:flex;gap:12px;align-items:flex-start;padding:14px 16px;border-bottom:1px solid #292524}" +
        ".sd-ar-head h2{margin:0;font-size:18px;font-weight:700}" +
        ".sd-ar-hint{margin:4px 0 0;color:#a8a29e;font-size:13px}" +
        ".sd-ar-x{margin-left:auto;border:1px solid #44403c;background:#1c1917;color:#fff;border-radius:8px;padding:10px 14px;font-weight:600;min-height:44px}" +
        ".sd-ar-stage{position:relative;flex:1;min-height:220px;background:#000}" +
        ".sd-ar-stage canvas{position:absolute;inset:0;width:100%;height:100%;touch-action:none}" +
        ".sd-ar-readout{position:absolute;left:50%;top:18%;transform:translate(-50%,-50%);font-size:42px;font-weight:800;letter-spacing:.02em;text-shadow:0 2px 12px #000;pointer-events:none}" +
        ".sd-ar-status{position:absolute;left:16px;right:16px;bottom:16px;margin:0;font-size:14px;color:#e7e5e4;text-shadow:0 1px 8px #000;pointer-events:none}" +
        ".sd-ar-foot{display:flex;gap:8px;padding:12px 16px calc(12px + env(safe-area-inset-bottom,0px));border-top:1px solid #292524;background:#1c1917}" +
        ".sd-ar-foot button{flex:1;min-height:48px;border-radius:10px;font-weight:700;font-size:15px}" +
        ".sd-ar-primary{border:0;background:#fafaf9;color:#1c1917}" +
        ".sd-ar-primary:disabled{opacity:.45}" +
        ".sd-ar-ghost{border:1px solid #57534e;background:transparent;color:#fafaf9}" +
        ".sd-ar-fallback{padding:16px;display:flex;flex-direction:column;gap:12px}" +
        ".sd-ar-fallback[hidden]{display:none!important}" +
        ".sd-ar-fallback label{display:flex;flex-direction:column;gap:6px;font-size:12px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;color:#a8a29e}" +
        ".sd-ar-fallback input{min-height:48px;border-radius:10px;border:1px solid #44403c;background:#0c0a09;color:#fff;padding:12px;font-size:16px}";
      document.head.appendChild(style);
    }
    return el;
  }

  function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
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
    const fallback = document.getElementById("sdArFallback");
    const stage = mask.querySelector(".sd-ar-stage");
    const foot = mask.querySelector(".sd-ar-foot");
    const manual = document.getElementById("sdArManual");
    const manualUse = document.getElementById("sdArManualUse");

    let session = null;
    let gl = null;
    let refSpace = null;
    let viewerSpace = null;
    let hitTestSource = null;
    let xrLayer = null;
    let points = [];
    let latestMm = 0;
    let closed = false;
    let raf = 0;

    function cleanup() {
      closed = true;
      if (raf) cancelAnimationFrame(raf);
      raf = 0;
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

    mask.hidden = false;
    fallback.hidden = true;
    if (stage) stage.hidden = false;
    if (foot) foot.hidden = false;
    setText("sdArTitle", "Measure " + label);
    setText("sdArHint", "Tap the first end, then the second end of the " + label + ".");
    setText("sdArStatus", "Starting camera…");
    latestMm = 0;
    points = [];
    updateReadout();
    if (manual) manual.value = "";

    closeBtn.onclick = () => cleanup();
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

    const support = await isArMeasureSupported();
    if (!support.ok) {
      if (stage) stage.hidden = true;
      if (foot) foot.hidden = true;
      fallback.hidden = false;
      const why =
        support.reason === "https"
          ? "AR measure needs a secure (https) page."
          : "This phone browser cannot run in-app AR measure (needs ARCore Chrome on Android, or a browser with WebXR AR).";
      setText(
        "sdArFallbackText",
        why + " Use a tape, type the millimetres below, then Use typed size. Results from AR are a check — tape wins when unsure."
      );
      return { mode: "manual" };
    }

    try {
      gl = canvas.getContext("webgl", { xrCompatible: true });
      if (!gl) throw new Error("WebGL not available");
      session = await navigator.xr.requestSession("immersive-ar", {
        requiredFeatures: ["hit-test", "local", "dom-overlay"],
        optionalFeatures: ["dom-overlay"],
        domOverlay: { root: mask }
      }).catch(async () => {
        // Older stacks may refuse dom-overlay — retry without it.
        return navigator.xr.requestSession("immersive-ar", {
          requiredFeatures: ["hit-test"],
          optionalFeatures: ["local", "local-floor"]
        });
      });

      xrLayer = new XRWebGLLayer(session, gl);
      await session.updateRenderState({ baseLayer: xrLayer });
      refSpace = await session.requestReferenceSpace("local").catch(() => session.requestReferenceSpace("local-floor"));
      viewerSpace = await session.requestReferenceSpace("viewer");
      if (session.requestHitTestSource) {
        hitTestSource = await session.requestHitTestSource({ space: viewerSpace });
      } else {
        throw new Error("Hit-test is not available on this device.");
      }

      setText("sdArStatus", "Tap the first end of the " + label + ".");

      session.addEventListener("end", () => {
        if (!closed) cleanup();
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
            setText("sdArStatus", "First point set. Tap the other end of the " + label + ".");
            return;
          }
          latestMm = roundMm(distMeters(points[0], points[1]));
          updateReadout();
          setText(
            "sdArStatus",
            "Measured " +
              latestMm +
              " mm. Use measurement, or Reset to try again. Allow a few mm tolerance — tape wins if tight."
          );
        } catch (err) {
          setText("sdArStatus", "Could not place that point. Try again.");
        }
      });

      const onXRFrame = (_time, frame) => {
        if (closed || !session) return;
        session.requestAnimationFrame(onXRFrame);
        const layer = session.renderState.baseLayer;
        if (!layer) return;
        gl.bindFramebuffer(gl.FRAMEBUFFER, layer.framebuffer);
        gl.clearColor(0, 0, 0, 0);
        gl.clear(gl.COLOR_BUFFER_BIT | gl.DEPTH_BUFFER_BIT);
        // Keep session alive; camera feed is composited by the XR runtime.
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
      return { mode: "webxr" };
    } catch (e) {
      if (stage) stage.hidden = true;
      if (foot) foot.hidden = true;
      fallback.hidden = false;
      setText(
        "sdArFallbackText",
        "Could not start AR (" +
          ((e && e.message) || "unsupported") +
          "). Use a tape and type millimetres below."
      );
      try {
        if (session) session.end();
      } catch (err) {}
      session = null;
      return { mode: "manual" };
    }
  }

  global.sdArMeasureSupported = isArMeasureSupported;
  global.sdOpenArMeasure = openArMeasure;
  global.sdArRoundMm = roundMm;
  global.sdArDistMeters = distMeters;
})(typeof window !== "undefined" ? window : globalThis);
