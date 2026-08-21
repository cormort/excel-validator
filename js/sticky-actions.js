/**
 * 表格上方的浮動操作列：清除選取 / 執行驗算 / 錯誤摘要
 * ponytail: 純代理既有 toolbar 按鈕，不動 ui-controller 邏輯
 */
(function () {
  function build(p) {
    const grid = document.getElementById(p + "gridContainer");
    if (!grid) return;

    const bar = document.createElement("div");
    bar.className = "sticky-actions hidden";
    bar.innerHTML = `
      <span class="sticky-actions-grip" title="拖曳移動">⠿</span>
      <span class="sticky-actions-status" data-status>—</span>
      <div class="sticky-actions-btns">
        <button class="btn btn-outline btn-sm" data-act="clear">✕ 清除選取</button>
        <button class="btn btn-success btn-sm" data-act="validate">✓ 執行驗算</button>
      </div>`;
    grid.parentNode.insertBefore(bar, grid);

    const src = {
      clear: document.getElementById(p + "btnClearLogicToolbar"),
      validate: document.getElementById(p + "btnValidate"),
    };
    const status = bar.querySelector("[data-status]");
    const btns = {
      clear: bar.querySelector('[data-act="clear"]'),
      validate: bar.querySelector('[data-act="validate"]'),
    };
    Object.keys(btns).forEach((k) =>
      btns[k].addEventListener("click", () => src[k]?.click())
    );

    makeDraggable(bar);

    const errPanel = document.getElementById(p + "errorPanel");
    const errCount = document.getElementById(p + "errorCount");
    const errDiff = document.getElementById(p + "errorDiff");
    const banner = document.getElementById(p + "successBanner");

    function sync() {
      bar.classList.toggle("hidden", grid.classList.contains("hidden"));
      btns.validate.disabled = !!src.validate?.disabled;

      const hasErr = errPanel && !errPanel.classList.contains("hidden");
      const ok = banner && !banner.classList.contains("hidden");
      status.classList.toggle("is-error", !!hasErr);
      status.classList.toggle("is-ok", !hasErr && !!ok);
      if (hasErr) {
        status.textContent = `⚠️ 發現 ${errCount?.textContent || 0} 個錯誤（差異 ${errDiff?.textContent || 0}）`;
      } else if (ok) {
        status.textContent = "✅ 驗算完成，未發現錯誤";
      } else {
        status.textContent = "尚未驗算";
      }
    }

    const obs = new MutationObserver(sync);
    [grid, errPanel, errCount, banner, src.validate]
      .filter(Boolean)
      .forEach((n) =>
        obs.observe(n, {
          attributes: true,
          childList: true,
          subtree: true,
          characterData: true,
        })
      );
    sync();
  }

  // 拖曳移動（避開按鈕），位置存在 localStorage
  function makeDraggable(bar) {
    const key = "stickyActionsPos";
    const saved = JSON.parse(localStorage.getItem(key) || "null");
    if (saved) place(saved.x, saved.y);

    function place(x, y) {
      const w = bar.offsetWidth || 300;
      const h = bar.offsetHeight || 50;
      x = Math.min(Math.max(x, 0), Math.max(innerWidth - w, 0));
      y = Math.min(Math.max(y, 0), Math.max(innerHeight - h, 0));
      bar.style.left = x + "px";
      bar.style.top = y + "px";
      bar.style.right = "auto";
      bar.style.bottom = "auto";
    }

    bar.addEventListener("pointerdown", (e) => {
      if (e.target.closest("button")) return;
      const r = bar.getBoundingClientRect();
      const dx = e.clientX - r.left;
      const dy = e.clientY - r.top;
      bar.setPointerCapture(e.pointerId);
      bar.classList.add("is-dragging");
      e.preventDefault();

      const move = (ev) => place(ev.clientX - dx, ev.clientY - dy);
      const up = () => {
        bar.classList.remove("is-dragging");
        bar.removeEventListener("pointermove", move);
        bar.removeEventListener("pointerup", up);
        const r2 = bar.getBoundingClientRect();
        localStorage.setItem(key, JSON.stringify({ x: r2.left, y: r2.top }));
      };
      bar.addEventListener("pointermove", move);
      bar.addEventListener("pointerup", up);
    });
  }

  document.addEventListener("DOMContentLoaded", () => {
    build("");
    build("p-");
  });
})();
