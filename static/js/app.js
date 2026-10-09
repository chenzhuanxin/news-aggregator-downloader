/* ============================================================
   新闻资讯下载器 —— 前端逻辑
   ============================================================ */
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);

  const state = {
    sites: [],
    jobId: null,
    pollTimer: null,
    startedAt: null,
  };

  /* ---------- 日期标签 ---------- */
  function fmtDate(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}${m}${day}`;
  }
  function fmtDateDash(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }
  function initDateLabels() {
    const today = new Date();
    const yest = new Date(today.getTime() - 86400000);
    $("todayLabel").textContent = fmtDate(today);
    $("yesterdayLabel").textContent = fmtDate(yest);
  }

  /* ---------- 站点渲染 ---------- */
  async function loadSites() {
    const res = await fetch("/api/sites");
    const sites = await res.json();
    state.sites = sites;
    const grid = $("siteGrid");
    grid.innerHTML = "";
    sites.forEach((s) => {
      const card = document.createElement("label");
      card.className = "site-card";
      card.dataset.id = s.id;
      card.innerHTML = `
        <input type="checkbox" value="${s.id}">
        <span class="checkbox-dot"></span>
        <div class="site-name">${escapeHtml(s.name)}</div>
        <div class="site-desc">${escapeHtml(s.desc)}</div>`;
      card.addEventListener("click", (e) => {
        // label 点击自动切换，但需同步视觉状态
        const cb = card.querySelector("input");
        card.classList.toggle("checked", cb.checked);
        updateCount();
      });
      grid.appendChild(card);
    });
    updateCount();
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    }[c]));
  }

  function selectedSites() {
    return [...document.querySelectorAll("#siteGrid input:checked")].map((i) => i.value);
  }

  function updateCount() {
    $("selectedCount").textContent = `已选 ${selectedSites().length} / ${state.sites.length}`;
  }

  /* ---------- 全选 ---------- */
  $("selectAll").addEventListener("change", (e) => {
    const on = e.target.checked;
    document.querySelectorAll("#siteGrid input").forEach((cb) => {
      cb.checked = on;
      cb.closest(".site-card").classList.toggle("checked", on);
    });
    updateCount();
  });

  /* ---------- 格式切换 ---------- */
  document.querySelectorAll(".fmt-chip").forEach((chip) => {
    chip.addEventListener("click", () => {
      document.querySelectorAll(".fmt-chip").forEach((c) => c.classList.remove("active"));
      chip.classList.add("active");
    });
  });

  /* ---------- 目录选择 ---------- */
  async function browse() {
    const res = await fetch("/api/browse", { method: "POST" });
    const data = await res.json();
    if (data.path) {
      $("outDir").value = data.path;
      $("btnOpen").disabled = false;
    } else if (data.error) {
      alert(data.error);
    }
  }
  $("btnBrowse").addEventListener("click", browse);

  async function openDir() {
    const path = $("outDir").value.trim();
    if (!path) return;
    const res = await fetch("/api/open", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path }),
    });
    const data = await res.json();
    if (!data.ok) alert(data.error);
  }
  $("btnOpen").addEventListener("click", openDir);
  // 手动输入路径后也启用“打开”按钮
  $("outDir").addEventListener("input", () => {
    $("btnOpen").disabled = !$("outDir").value.trim();
  });

  /* ---------- 一键下载 ---------- */
  $("btnRun").addEventListener("click", run);

  /* ---------- 启动下载 ---------- */
  async function run() {
    const sites = selectedSites();
    const dates = [...document.querySelectorAll("#dateOptions input:checked")].map((i) => i.value);
    const outDir = $("outDir").value.trim();
    const format = document.querySelector('input[name="format"]:checked').value;

    if (!sites.length) return alert("请至少勾选一个新闻站点");
    if (!dates.length) return alert("请至少勾选一个日期（今天 / 前一天）");
    if (!outDir) return alert("请先选择保存目录");

    const res = await fetch("/api/run", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sites, dates, out_dir: outDir, format }),
    });
    const data = await res.json();
    if (data.error) return alert(data.error);

    state.jobId = data.job_id;
    state.startedAt = Date.now();
    resetProgress();
    showProgress();
    $("btnRun").disabled = true;
    $("btnStop").disabled = false;
    poll();
  }

  /* ---------- 进度 ---------- */
  function resetProgress() {
    const box = $("siteProgress");
    box.innerHTML = "";
    state.sites.filter((s) => selectedSites().includes(s.id)).forEach((s) => {
      const row = document.createElement("div");
      row.className = "sp-row";
      row.dataset.site = s.id;
      row.innerHTML = `
        <span class="sp-status wait">…</span>
        <div class="sp-name">${escapeHtml(s.name)}</div>
        <div class="sp-msg">等待中</div>`;
      box.appendChild(row);
    });
  }

  function showProgress() {
    $("progressPanel").style.display = "block";
    $("progressPanel").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function setOverall(pct, text, stateLabel, cls) {
    $("overallBar").style.width = pct + "%";
    $("overallText").textContent = text;
    const pill = $("overallState");
    pill.textContent = stateLabel;
    pill.className = "state-pill" + (cls ? " " + cls : "");
    if (state.startedAt) {
      const sec = Math.round((Date.now() - state.startedAt) / 1000);
      const mm = String(Math.floor(sec / 60)).padStart(2, "0");
      const ss = String(sec % 60).padStart(2, "0");
      $("elapsed").textContent = `已用 ${mm}:${ss}`;
    }
  }

  function renderStep(step, idx) {
    // 用站点 id 匹配进度行（step.site_id 对应 resetProgress 里的 data-site）
    const row = document.querySelector(`.sp-row[data-site="${step.site_id}"]`);
    if (!row) return;
    const dot = row.querySelector(".sp-status");
    const msg = row.querySelector(".sp-msg");
    const map = {
      running: ["run", "…", "抓取中…"],
      success: ["ok", "✓", step.message || "完成"],
      fail: ["fail", "✕", step.message || "失败"],
      empty: ["empty", "·", step.message || "无内容"],
      cancel: ["fail", "■", step.message || "已停止"],
    };
    const [cls, icon, text] = map[step.status] || ["wait", "…", step.message || ""];
    dot.className = "sp-status " + cls;
    dot.textContent = icon;
    msg.textContent = text;
  }

  async function poll() {
    if (!state.jobId) return;
    const res = await fetch(`/api/jobs/${state.jobId}`);
    if (res.status === 404) { stopPoll(); return; }
    const job = await res.json();

    // 渲染每站点状态
    job.steps.forEach((step, idx) => renderStep(step, idx));

    const total = job.total || 0;
    const doneCount = job.steps.filter((s) =>
      ["success", "fail", "empty", "cancel"].includes(s.status)).length;
    const pct = total ? Math.round((doneCount / total) * 100) : 0;
    setOverall(pct, `${doneCount} / ${total}`, "进行中");

    if (job.state === "done" || job.state === "stopped") {
      stopPoll();
      const finished = job.state === "done";
      setOverall(100, `${total} / ${total}`, finished ? "全部完成" : "已停止", finished ? "done" : "");
      $("btnRun").disabled = false;
      $("btnStop").disabled = true;
      if (job.out_dir) {
        $("outDir").value = job.out_dir;
        $("btnOpen").disabled = false;
      }
      const failed = job.steps.filter((s) => s.status === "fail").length;
      if (failed) alert(`下载完成，但有 ${failed} 个站点失败（可查看进度详情）。`);
    } else if (job.state === "error") {
      stopPoll();
      setOverall(0, "出错", "出错了", "err");
      $("btnRun").disabled = false;
      $("btnStop").disabled = true;
    } else if (job.state === "stopping") {
      setOverall(pct, `${doneCount} / ${total}`, "正在停止…");
    }
  }

  function stopPoll() {
    if (state.pollTimer) { clearInterval(state.pollTimer); state.pollTimer = null; }
    state.jobId = null;
  }

  /* ---------- 停止 ---------- */
  async function stop() {
    if (!state.jobId) return;
    const res = await fetch("/api/stop", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ job_id: state.jobId }),
    });
    const data = await res.json();
    if (data.ok) {
      $("btnStop").disabled = true;
      setOverall(0, "-", "停止中…");
    }
  }
  $("btnStop").addEventListener("click", stop);

  /* ---------- 轮询 ---------- */
  function startPoll() {
    state.pollTimer = setInterval(poll, 1200);
  }

  /* ---------- 初始化 ---------- */
  async function init() {
    initDateLabels();
    await loadSites();
    startPoll();
  }
  init();
})();
