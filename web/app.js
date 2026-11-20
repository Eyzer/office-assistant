/* office办公助手 · 前端
   结构：首页功能墙（按类别排列）→ 点卡片进入工作区 → 运行，日志走底部抽屉。 */
const TOKEN = window.__TOKEN__;
const API_VER = window.__API_VER__ || 0;
const APP = window.__APP__ || { name: "office办公助手" };
const $ = (s) => document.querySelector(s);

const state = {
  tools: [],
  tool: null,
  inputs: [],
  params: {},
  taskId: null,
  polling: null,
  logCount: 0,
  unread: 0,
  groupFilter: "*",
  fsMode: "files",
  fsPath: "",
  fsParent: "",
  lastOutput: null,
};

const GROUP_LABEL = { common: "通用文件", excel: "表格处理", word: "Word 文档" };
const GROUP_DESC = {
  common: "重命名、复制、压缩、编码、清单等与格式无关的操作",
  excel: "xlsx / csv 的拆分、合并、清洗、格式、公式与导出",
  word: "docx 的内容处理、排版、结构与转换",
};

/* ---------- 基础设施 ---------- */
async function api(path, opts = {}) {
  const sep = path.includes("?") ? "&" : "?";
  const url = path + sep + "t=" + encodeURIComponent(TOKEN);
  let res;
  try {
    res = await fetch(url, { headers: { "Content-Type": "application/json" }, ...opts });
  } catch (e) {
    return { ok: false, error: "连不上后台服务：" + e };
  }
  const text = await res.text();
  try {
    return JSON.parse(text);
  } catch (e) {
    return { ok: false, error: "服务端返回异常（HTTP " + res.status + "）" };
  }
}

let toastTimer = null;
function toast(msg, kind = "") {
  const el = $("#toast");
  el.textContent = msg;
  el.className = "toast " + kind;
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (el.hidden = true), 3200);
}

function showBanner(msg) {
  const b = $("#banner");
  b.textContent = msg;
  b.hidden = !msg;
}

/* ---------- 日志抽屉 ---------- */
function log(msg, kind = "") {
  const el = $("#log");
  const line = document.createElement("div");
  line.textContent = msg;
  if (kind === "err") line.style.color = "#ef4444";
  if (kind === "ok") line.style.color = "#22c55e";
  if (kind === "warn") line.style.color = "#f59e0b";
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
  if ($("#logDrawer").hidden) {
    state.unread++;
    ["#logBadge", "#fabBadge"].forEach((s) => {
      const b = $(s);
      b.textContent = state.unread > 99 ? "99+" : state.unread;
      b.hidden = false;
    });
  }
}

function openLog() {
  $("#logDrawer").hidden = false;
  $("#logFab").hidden = true;
  state.unread = 0;
  $("#logBadge").hidden = true;
  $("#log").scrollTop = $("#log").scrollHeight;
}

function closeLog() {
  $("#logDrawer").hidden = true;
  $("#logFab").hidden = false;
}

/* ---------- 视图切换 ---------- */
function showHome(push) {
  $("#viewHome").hidden = false;
  $("#viewWork").hidden = true;
  if (push !== false && location.hash) history.pushState(null, "", location.pathname + location.search);
  const s = $("#toolSearch");
  if (s) s.focus();
}

function showWork(push) {
  $("#viewHome").hidden = true;
  $("#viewWork").hidden = false;
  $(".main").scrollTop = 0;
  if (push && state.tool && location.hash !== "#tool=" + state.tool.id) {
    history.pushState(null, "", "#tool=" + state.tool.id);
  }
}

/* ---------- 首页功能墙 ---------- */
function cardOf(t) {
  const el = document.createElement("button");
  el.className = "card";
  el.dataset.id = t.id;
  el.dataset.group = t.group;
  el.dataset.cat = t.cat || "";
  el.dataset.search = (t.name + " " + (t.desc || "") + " " + t.id + " " + (t.cat || "") +
    " " + (GROUP_LABEL[t.group] || "")).toLowerCase();

  const ico = document.createElement("span");
  ico.className = "card-ico g-" + t.group;
  // 用功能名首字当图标，避免依赖 emoji 字体
  const short = (t.name || "?").replace(/[（(].*$/, "").slice(0, 1);
  ico.textContent = short || "·";

  const body = document.createElement("span");
  body.className = "card-body";
  const name = document.createElement("span");
  name.className = "card-name";
  name.textContent = t.name;
  const desc = document.createElement("span");
  desc.className = "card-desc";
  desc.textContent = t.desc || "";
  body.append(name, desc);
  el.append(ico, body);
  el.onclick = () => selectTool(t);
  return el;
}

function renderHome() {
  const box = $("#homeGroups");
  box.innerHTML = "";
  const order = ["common", "excel", "word"];
  order.forEach((g) => {
    const items = state.tools.filter((t) => t.group === g);
    if (!items.length) return;
    const sec = document.createElement("section");
    sec.className = "gsec";
    sec.dataset.group = g;

    const head = document.createElement("div");
    head.className = "gsec-head";
    const h2 = document.createElement("h2");
    h2.textContent = GROUP_LABEL[g] || g;
    const sp = document.createElement("span");
    sp.textContent = (GROUP_DESC[g] || "") + " · " + items.length + " 个功能";
    head.append(h2, sp);
    sec.appendChild(head);

    const cats = [];
    items.forEach((t) => {
      const c = t.cat || "其他";
      if (cats.indexOf(c) < 0) cats.push(c);
    });
    cats.forEach((c) => {
      const blk = document.createElement("div");
      blk.className = "cat-block";
      const ct = document.createElement("div");
      ct.className = "cat-title";
      ct.textContent = c;
      const cards = document.createElement("div");
      cards.className = "cards";
      items.filter((t) => (t.cat || "其他") === c).forEach((t) => cards.appendChild(cardOf(t)));
      blk.append(ct, cards);
      sec.appendChild(blk);
    });
    box.appendChild(sec);
  });
  renderChips();
  applySearch($("#toolSearch") ? $("#toolSearch").value : "");
}

function renderChips() {
  const box = $("#groupChips");
  box.innerHTML = "";
  const mk = (key, label, n) => {
    const b = document.createElement("button");
    b.className = "chip" + (state.groupFilter === key ? " on" : "");
    b.textContent = label;
    if (n !== undefined) {
      const s = document.createElement("span");
      s.className = "n";
      s.textContent = n;
      b.appendChild(s);
    }
    b.onclick = () => {
      state.groupFilter = key;
      renderChips();
      applySearch($("#toolSearch").value);
    };
    return b;
  };
  box.appendChild(mk("*", "全部", state.tools.length));
  ["common", "excel", "word"].forEach((g) => {
    const n = state.tools.filter((t) => t.group === g).length;
    if (n) box.appendChild(mk(g, GROUP_LABEL[g] || g, n));
  });
}

function applySearch(q) {
  q = (q || "").trim().toLowerCase();
  let visible = 0;
  document.querySelectorAll(".card").forEach((el) => {
    const hitGroup = state.groupFilter === "*" || el.dataset.group === state.groupFilter;
    const hitText = !q || (el.dataset.search || "").includes(q);
    el.hidden = !(hitGroup && hitText);
    if (!el.hidden) visible++;
  });
  document.querySelectorAll(".cat-block").forEach((b) => {
    b.hidden = !Array.from(b.querySelectorAll(".card")).some((x) => !x.hidden);
  });
  document.querySelectorAll(".gsec").forEach((s) => {
    s.hidden = !Array.from(s.querySelectorAll(".card")).some((x) => !x.hidden);
  });
  $("#homeEmpty").hidden = visible > 0;
}

async function loadTools() {
  const r = await api("/api/tools");
  if (!r.ok) {
    showBanner("读取功能清单失败：" + (r.error || "") + "（若刚升级过程序，请关闭后重新启动）");
    return;
  }
  state.tools = r.tools || [];
  renderHome();
  route(false);
}

function route(push) {
  if (location.hash === "#deps") return showDeps();
  if (location.hash === "#about") return showAbout();
  if (location.hash === "#log") return openLog();
  const m = /tool=([\w.\-]+)/.exec(location.hash || "");
  if (!m) return showHome(false);
  const t = state.tools.find((x) => x.id === m[1]);
  if (!t) return showHome(false);
  if (!state.tool || state.tool.id !== t.id) selectTool(t, push);
}

/* ---------- 进入某个功能 ---------- */
function selectTool(t, push) {
  state.tool = t;
  state.params = {};
  localStorage.setItem("dbt_last_tool", t.id);
  $("#toolName").textContent = t.name;
  $("#toolCat").textContent = (GROUP_LABEL[t.group] || "") + (t.cat ? " / " + t.cat : "");
  $("#toolDesc").textContent = t.desc || "";
  renderParams(t.fields || []);
  showWork(push !== false);
}

function renderParams(fields) {
  const box = $("#paramForm");
  box.innerHTML = "";
  if (!fields.length) {
    box.innerHTML = '<div class="empty">该功能无需参数</div>';
    return;
  }
  fields.forEach((f) => {
    const wrap = document.createElement("div");
    if (f.type === "checkbox") {
      wrap.className = "field check";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.id = "p_" + f.key;
      cb.checked = !!f.default;
      state.params[f.key] = cb.checked;
      cb.onchange = () => (state.params[f.key] = cb.checked);
      const lb = document.createElement("label");
      lb.textContent = f.label;
      wrap.append(cb, lb);
    } else {
      wrap.className = f.wide ? "field wide" : "field";
      const lb = document.createElement("label");
      lb.textContent = f.label;
      wrap.appendChild(lb);
      let input;
      if (f.type === "select") {
        input = document.createElement("select");
        (f.options || []).forEach((o) => {
          // 兼容三种写法：{value,label} / [value,label] / "value"
          let val, txt;
          if (Array.isArray(o)) {
            val = o[0];
            txt = o.length > 1 ? o[1] : o[0];
          } else if (o !== null && typeof o === "object") {
            val = o.value;
            txt = o.label !== undefined ? o.label : o.value;
          } else {
            val = txt = o;
          }
          const opt = document.createElement("option");
          opt.value = val;
          opt.textContent = txt;
          input.appendChild(opt);
        });
        if (f.default !== undefined) input.value = f.default;
      } else if (f.type === "textarea") {
        // 长文本占满整行，否则后面的参数会被挤到看不见的地方
        wrap.className = "field wide";
        input = document.createElement("textarea");
        input.rows = f.rows || 3;
        if (f.default !== undefined) input.value = f.default;
        if (f.placeholder) input.placeholder = f.placeholder;
      } else {
        input = document.createElement("input");
        input.type = f.type === "number" ? "number" : "text";
        if (f.default !== undefined) input.value = f.default;
        if (f.placeholder) input.placeholder = f.placeholder;
      }
      input.id = "p_" + f.key;
      state.params[f.key] = input.value;
      input.oninput = () => {
        state.params[f.key] = f.type === "number" ? Number(input.value) : input.value;
      };
      wrap.appendChild(input);
      if (f.help) {
        const h = document.createElement("div");
        h.className = "help";
        h.textContent = f.help;
        wrap.appendChild(h);
      }
    }
    box.appendChild(wrap);
  });
}

/* ---------- 文件列表 ---------- */
function renderFiles() {
  const ul = $("#fileList");
  ul.innerHTML = "";
  $("#fileCount").textContent = state.inputs.length;
  if (!state.inputs.length) {
    ul.innerHTML = '<li class="empty">尚未添加文件</li>';
    return;
  }
  state.inputs.forEach((p, i) => {
    const li = document.createElement("li");
    const s = document.createElement("span");
    s.textContent = p;
    s.title = p;
    const rm = document.createElement("span");
    rm.className = "rm";
    rm.textContent = "移除";
    rm.onclick = () => {
      state.inputs.splice(i, 1);
      renderFiles();
    };
    li.append(s, rm);
    ul.appendChild(li);
  });
}

function addPaths(paths, opts) {
  let added = 0;
  (paths || []).forEach((p) => {
    if (!p) return;
    const n = String(p).replace(/\//g, "\\").replace(/\\+$/, "");
    if (!state.inputs.includes(n)) {
      state.inputs.push(n);
      added++;
    }
  });
  // 批量添加时传 {render:false}，由调用方统一重绘一次，避免逐条重排 DOM
  if (!opts || opts.render !== false) renderFiles();
  return added;
}

/* ---------- 打开输出目录 ---------- */
async function openDir(target) {
  const p = target || state.lastOutput || $("#outputDir").value.trim() || "output";
  const r = await api("/api/open", { method: "POST", body: JSON.stringify({ path: p }) });
  if (r && r.ok) {
    toast("已打开 " + r.path, "ok");
    log("打开目录：" + r.path, "ok");
  } else {
    const why = (r && (r.error || r.reason)) || "未知原因";
    toast("打不开：" + why, "err");
    log("打开目录失败：" + why, "err");
  }
}

/* ---------- 系统文件对话框 ----------
   后台能调 zenity/kdialog 就弹真正的系统框（返回真实路径，不用上传副本）；
   调不到就退回浏览器自带的 <input type=file>。 */
async function pickNative(mode) {
  const r = await api("/api/pick?mode=" + encodeURIComponent(mode));
  if (r.ok) return r;
  log("系统文件对话框不可用：" + (r.error || r.reason || "未知原因") + "，改用浏览器选择框", "warn");
  return null;
}

/* 对话框弹出前先给按钮一个"打开中…"的即时反馈，避免看起来像点了没反应 */
async function withBusy(btn, fn) {
  const label = btn.textContent;
  btn.disabled = true;
  btn.textContent = "打开中…";
  try {
    return await fn();
  } finally {
    btn.disabled = false;
    btn.textContent = label;
  }
}

async function addFiles() {
  const r = await withBusy($("#addFiles"), () => pickNative("files"));
  if (r) {
    if (r.cancelled || !r.paths || !r.paths.length) return;
    const n = addPaths(r.paths);
    toast("已添加 " + n + " 个文件", "ok");
    log("已添加 " + n + " 个文件（系统对话框）", "ok");
    return;
  }
  $("#fileInput").click();
}

async function addFolder() {
  const r = await withBusy($("#addFolder"), () => pickNative("dir"));
  if (r) {
    if (r.cancelled || !r.paths || !r.paths.length) return;
    const n = addPaths(r.paths);
    toast("已添加 " + n + " 个文件夹", "ok");
    log("已添加文件夹：" + r.paths.join("，"), "ok");
    return;
  }
  $("#dirInput").click();
}

/* ---------- 服务器目录浏览（备用） ---------- */
async function openFs(mode) {
  state.fsMode = mode;
  $("#fsModal").hidden = false;
  $("#fsTitle").textContent = mode === "outdir" ? "选择输出目录" : "浏览目录（选择文件或文件夹）";
  const start = mode === "outdir" ? ($("#outputDir").value || "") : "";
  await browse(start);
}

async function browse(path) {
  const r = await api("/api/fs?path=" + encodeURIComponent(path || ""));
  if (!r.ok) {
    log("目录读取失败: " + (r.error || ""), "err");
    return;
  }
  state.fsPath = r.path;
  state.fsParent = r.parent;
  $("#fsPath").value = r.path;
  const ul = $("#fsList");
  ul.innerHTML = "";
  (r.dirs || []).forEach((d) => {
    const li = document.createElement("li");
    li.textContent = "📁 " + d;
    li.onclick = () => browse(r.path + (r.path.endsWith("/") || r.path.endsWith("\\") ? "" : "/") + d);
    ul.appendChild(li);
  });
  if (state.fsMode === "files") {
    const sep = r.path.endsWith("/") || r.path.endsWith("\\") ? "" : "/";
    (r.files || []).forEach((f) => {
      const li = document.createElement("li");
      li.textContent = "📄 " + f;
      li.onclick = () => {
        addPaths([r.path + sep + f]);
        log("已添加: " + r.path + sep + f);
      };
      ul.appendChild(li);
    });
  }
}

/* ---------- 拖拽上传 ----------
   浏览器不给拖入文件的真实路径，只能把内容传给本机服务落盘一次，
   再拿回真实路径交给后端处理（全程 127.0.0.1，不出本机）。 */
const ACCEPT_EXT = new Set([
  ".xlsx", ".xlsm", ".xlsb", ".xls", ".csv", ".tsv", ".txt", ".json", ".md",
  ".docx", ".docm", ".doc", ".pptx", ".ppt", ".pdf",
  ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp",
  ".dbf", ".vcf",
]);
const SID = "u" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

function extOf(name) {
  const i = name.lastIndexOf(".");
  return i < 0 ? "" : name.slice(i).toLowerCase();
}

/* 限流并发：目录里上千个文件时串行 await 会明显卡住 */
async function mapLimit(items, limit, fn) {
  const out = new Array(items.length);
  let cursor = 0;
  async function run() {
    while (cursor < items.length) {
      const k = cursor++;
      out[k] = await fn(items[k]);
    }
  }
  const lanes = Math.min(limit, items.length) || 1;
  await Promise.all(Array.from({ length: lanes }, run));
  return out;
}

function fmtSize(n) {
  if (!n) return "0";
  if (n < 1024) return n + " B";
  if (n < 1024 * 1024) return (n / 1024).toFixed(0) + " KB";
  if (n < 1024 * 1024 * 1024) return (n / 1024 / 1024).toFixed(1) + " MB";
  return (n / 1024 / 1024 / 1024).toFixed(2) + " GB";
}

function readAllEntries(reader) {
  // readEntries 一次最多返回 100 条，必须循环读到空
  return new Promise((resolve) => {
    const all = [];
    const step = () =>
      reader.readEntries((batch) => {
        if (!batch.length) return resolve(all);
        all.push.apply(all, batch);
        step();
      }, () => resolve(all));
    step();
  });
}

async function walkEntry(entry, prefix) {
  if (!entry) return [];
  if (entry.isFile) {
    const f = await new Promise((res) => entry.file((x) => res(x), () => res(null)));
    return f ? [{ file: f, rel: prefix + entry.name }] : [];
  }
  if (entry.isDirectory) {
    const kids = await readAllEntries(entry.createReader());
    // 兄弟节点并行读，读完后按原顺序合并，保证列表顺序稳定
    const groups = await mapLimit(kids, 8, (k) => walkEntry(k, prefix + entry.name + "/"));
    const out = [];
    groups.forEach((g) => {
      for (const it of g) out.push(it);
    });
    return out;
  }
  return [];
}

async function handleDrop(entries, files) {
  const raw = [];
  if (entries && entries.length) {
    $("#dzProgress").textContent = "正在读取文件…";
    const groups = await mapLimit(entries, 4, (e) => walkEntry(e, ""));
    groups.forEach((g) => {
      for (const it of g) raw.push(it);
    });
  }
  if (!raw.length && files && files.length) {
    files.forEach((f) => raw.push({ file: f, rel: f.webkitRelativePath || f.name }));
  }
  if (!raw.length) {
    toast("没有读到文件，请直接把文件或文件夹拖进窗口", "err");
    return;
  }

  const keep = [];
  let skipped = 0;
  raw.forEach((it) => {
    const n = it.file.name;
    if (n.startsWith(".") || n.startsWith("~$") || !ACCEPT_EXT.has(extOf(n))) {
      skipped++;
      return;
    }
    keep.push(it);
  });
  if (!keep.length) {
    toast("拖进来的 " + skipped + " 个文件类型不支持（支持 xlsx/docx/csv/pdf/图片等）", "warn");
    log("没有可处理的文件：跳过 " + skipped + " 个不支持的（如 .et 请先另存为 xlsx）", "warn");
    return;
  }
  await uploadAll(keep, skipped);
}

/* 上传并发数：浏览器对同一主机通常 6 条连接，留 2 条给界面轮询 */
const UPLOAD_CONCURRENCY = 4;

async function uploadAll(list, skipped) {
  const box = $("#dzProgress");
  const start = Date.now();
  const failed = [];
  const totalBytes = list.reduce((a, it) => a + (it.file.size || 0), 0);
  let sentBytes = 0;
  let done = 0;
  let added = 0;
  let lastPaint = 0;

  function paint(force) {
    const now = Date.now();
    if (!force && now - lastPaint < 100) return;
    lastPaint = now;
    const pct = totalBytes ? Math.min(100, Math.round((sentBytes / totalBytes) * 100)) : 0;
    box.textContent =
      `上传中 ${done}/${list.length} · ${fmtSize(sentBytes)} / ${fmtSize(totalBytes)}（${pct}%）`;
  }

  function uploadOne(it) {
    return new Promise((resolve, reject) => {
      const url =
        "/api/upload?t=" + encodeURIComponent(TOKEN) +
        "&sid=" + encodeURIComponent(SID) +
        "&name=" + encodeURIComponent(it.rel);
      const xhr = new XMLHttpRequest();
      xhr.open("POST", url, true);
      let prev = 0;
      // fetch 拿不到上传进度，只有 XHR 有 upload.onprogress
      xhr.upload.onprogress = (e) => {
        if (!e.lengthComputable) return;
        sentBytes += e.loaded - prev;
        prev = e.loaded;
        paint(false);
      };
      xhr.onload = () => {
        let j;
        try {
          j = JSON.parse(xhr.responseText);
        } catch (e) {
          j = { ok: false, error: "HTTP " + xhr.status };
        }
        if (j.ok && j.path) resolve(j);
        else reject(new Error(j.error || "HTTP " + xhr.status));
      };
      xhr.onerror = () => reject(new Error("网络中断"));
      xhr.send(it.file);
    });
  }

  let cursor = 0;
  const sjzHit = [];
  async function worker() {
    while (cursor < list.length) {
      const it = list[cursor++];
      try {
        const r = await uploadOne(it);
        // 关掉逐条重绘，否则 N 个文件要重排 N 次 DOM，几百个文件就卡死
        added += addPaths([r.path], { render: false });
        if (r.sjz && (r.sjz.detected || r.sjz.error)) sjzHit.push({ name: it.rel, info: r.sjz });
      } catch (e) {
        failed.push(it.rel + "：" + ((e && e.message) || e));
      }
      done++;
      paint(false);
    }
  }

  paint(true);
  const lanes = Math.min(UPLOAD_CONCURRENCY, list.length);
  await Promise.all(Array.from({ length: lanes }, worker));

  renderFiles();
  const secs = ((Date.now() - start) / 1000).toFixed(1);
  box.textContent = `已加入 ${added} 个文件，共 ${fmtSize(totalBytes)}（${secs}s）`;
  setTimeout(() => (box.textContent = ""), 6000);

  if (failed.length) {
    failed.slice(0, 5).forEach((m) => log("上传失败 " + m, "err"));
    toast("有 " + failed.length + " 个文件上传失败，详见日志", "err");
  }
  if (sjzHit.length) {
    const okN = sjzHit.filter((x) => x.info.decrypted).length;
    const badN = sjzHit.filter((x) => x.info.error).length;
    sjzHit.slice(0, 5).forEach((x) => {
      const s = x.info;
      log(
        s.decrypted
          ? `sjz 已解密：${x.name}（判据 ${s.via}，${fmtSize(s.size)}）`
          : `sjz 识别到但未解密：${x.name}${s.error ? "（" + s.error + "）" : ""}`,
        s.decrypted ? "ok" : "warn"
      );
    });
    if (okN) toast(`已自动解密 ${okN} 个 sjz 文件`, "ok");
    if (badN) toast(`有 ${badN} 个 sjz 文件解密失败，详见日志`, "warn");
  }
  log(
    `拖入 ${list.length} 个文件（${fmtSize(totalBytes)}），加入 ${added} 个` +
      (skipped ? `，跳过 ${skipped} 个不支持的文件` : "") +
      `，用时 ${secs}s`,
    added ? "ok" : "warn"
  );
  if (added && !state.tool) toast("已加入 " + added + " 个文件，接着选一个功能", "ok");
}

function bindDrop() {
  let depth = 0;
  const on = () => document.body.classList.add("dragging");
  const off = () => document.body.classList.remove("dragging");

  window.addEventListener("dragenter", (e) => {
    e.preventDefault();
    depth++;
    on();
  });
  window.addEventListener("dragover", (e) => {
    e.preventDefault();
    if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
  });
  window.addEventListener("dragleave", () => {
    depth--;
    if (depth <= 0) {
      depth = 0;
      off();
    }
  });
  // 关键：数据只在事件同步阶段有效，先取出来再异步处理
  window.addEventListener("drop", (e) => {
    e.preventDefault();
    depth = 0;
    off();
    const dt = e.dataTransfer;
    if (!dt) return;
    const files = Array.from(dt.files || []);
    const entries = [];
    if (dt.items) {
      for (const it of Array.from(dt.items)) {
        if (it.kind !== "file") continue;
        try {
          const en = it.webkitGetAsEntry ? it.webkitGetAsEntry() : null;
          if (en) entries.push(en);
        } catch (err) {
          /* 忽略，退回 files */
        }
      }
    }
    if (!files.length && !entries.length) {
      toast("没读到拖进来的内容，试试「添加文件」按钮", "warn");
      return;
    }
    handleDrop(entries, files);
  });

  const dz = $("#dropzone");
  if (dz) dz.onclick = addFiles;

  const fi = $("#fileInput");
  if (fi) {
    fi.onchange = async () => {
      const list = Array.from(fi.files || []).map((f) => ({ file: f, rel: f.name }));
      fi.value = "";
      if (!list.length) return;
      const keep = list.filter(
        (it) => !it.file.name.startsWith(".") && ACCEPT_EXT.has(extOf(it.file.name))
      );
      if (!keep.length) return log("所选文件类型不支持", "warn");
      await uploadAll(keep, list.length - keep.length);
    };
  }

  const di = $("#dirInput");
  if (di) {
    di.onchange = async () => {
      const list = Array.from(di.files || []).map((f) => ({
        file: f,
        rel: f.webkitRelativePath || f.name,
      }));
      di.value = "";
      if (!list.length) return;
      const keep = list.filter(
        (it) => !it.file.name.startsWith(".") && ACCEPT_EXT.has(extOf(it.file.name))
      );
      if (!keep.length) return log("所选文件夹里没有支持的文件", "warn");
      await uploadAll(keep, list.length - keep.length);
    };
  }
}

/* ---------- 环境自检 ---------- */
async function showDeps() {
  const box = $("#infoBody");
  $("#infoTitle").textContent = "环境自检";
  $("#infoModal").hidden = false;
  box.innerHTML = '<div class="empty">检查中…</div>';

  const r = await api("/api/deps");
  if (!r.ok) {
    box.innerHTML =
      '<div class="empty" style="color:#f59e0b">自检失败：' +
      (r.error || "未知错误") +
      "<br><br>如果刚升级过程序，请关闭程序后重新启动（后台服务仍是旧版本时功能接口会对不上）。</div>";
    log("自检失败：" + (r.error || ""), "err");
    return;
  }

  const rows = (r.deps || [])
    .map((d) => {
      const st = d.ok ? '<span class="st ok">✓</span>' : '<span class="st no">✗</span>';
      const extra = d.ok
        ? '<span class="dep-ver">' + (d.version || "已安装") + "</span>"
        : '<span class="dep-ver">apt install ' + (d.apt || "") + "</span>";
      return (
        '<div class="dep-row">' + st +
        '<span>' + d.name + '<span class="dep-use"> · ' + (d.use || "") + "</span></span>" +
        extra + "</div>"
      );
    })
    .join("");
  const bad = (r.deps || []).filter((d) => !d.ok).length;
  box.innerHTML =
    '<div class="kv"><dt>版本</dt><dd>v' + (r.version || APP.version) +
    "（接口 " + (r.api_version !== undefined ? r.api_version : "?") + "）</dd>" +
    "<dt>暂存目录</dt><dd>" + (r.upload_dir || "-") + "</dd></div>" +
    '<div class="about-sec">依赖组件</div>' + rows +
    (bad
      ? '<div class="about-foot" style="color:#f59e0b">有 ' + bad +
        " 项缺失，相关功能会提示缺依赖。装系统包或在 vendor/ 放入 wheel 即可。</div>"
      : '<div class="about-foot" style="color:#22c55e">全部就绪。</div>');
  log("环境自检完成：" + ((r.deps || []).length - bad) + " 项可用，缺 " + bad + " 项", bad ? "warn" : "ok");
}

/* ---------- 关于 ---------- */
async function showAbout() {
  const box = $("#infoBody");
  $("#infoTitle").textContent = "关于";
  $("#infoModal").hidden = false;
  box.innerHTML = '<div class="empty">读取中…</div>';
  const r = await api("/api/about");
  if (!r.ok) {
    box.innerHTML = '<div class="empty">' + (r.error || "读取失败") + "</div>";
    return;
  }
  const a = r.about;
  const tp = (a.third_party || [])
    .map((t) => '<div class="tp-row"><span>' + t.name + "</span><span>" + t.use + '</span><span class="lic">' + t.license + "</span></div>")
    .join("");
  box.innerHTML =
    '<div class="about-head"><span class="logo">OA</span><div>' +
    '<div class="about-name">' + a.name + "</div>" +
    '<div class="about-tag">' + a.tagline + " · v" + a.version + "</div></div></div>" +
    '<div class="kv">' +
    "<dt>作者</dt><dd>" + a.author + "</dd>" +
    "<dt>邮箱</dt><dd><a href='mailto:" + a.email + "'>" + a.email + "</a></dd>" +
    "<dt>主页</dt><dd><a href='" + a.homepage + "' target='_blank' rel='noopener'>" + a.homepage + "</a></dd>" +
    "<dt>许可</dt><dd>" + a.license + "</dd>" +
    "<dt>运行环境</dt><dd>" + a.python + " · " + a.platform + "</dd>" +
    "</div>" +
    '<div class="muted" style="margin-top:10px">' + (a.description || "") + "</div>" +
    '<div class="about-sec">第三方组件</div>' + tp +
    '<div class="about-foot">本程序为本地离线工具，只在 127.0.0.1 上提供界面，不上传任何文件到外网。</div>';
}

/* ---------- 启动 ---------- */
document.addEventListener("DOMContentLoaded", () => {
  bindDrop();
  renderFiles();
  loadTools();

  api("/api/ping").then((p) => {
    if (!p.ok) {
      showBanner("后台服务版本过旧或接口不匹配：" + (p.error || "") + " 请关闭程序后重新启动。");
    } else if (API_VER && p.api_version !== API_VER) {
      showBanner("页面与后台服务版本不一致（页面 " + API_VER + " / 服务 " + p.api_version + "），请重启程序。");
    }
  });

  // 预热系统文件对话框进程：点「添加文件」时才不用现启一个 PowerShell
  api("/api/pick?mode=warm");

  const ts = $("#toolSearch");
  ts.oninput = () => applySearch(ts.value);
  window.addEventListener("popstate", () => route(false));
  window.addEventListener("hashchange", () => route(false));
  $("#brandHome").onclick = () => showHome();
  $("#btnHome").onclick = () => showHome();
  $("#btnBack").onclick = () => showHome();
  $("#btnDeps").onclick = showDeps;
  $("#btnAbout").onclick = showAbout;
  $("#btnLog").onclick = () => ($("#logDrawer").hidden ? openLog() : closeLog());
  $("#logFab").onclick = openLog;
  $("#closeLog").onclick = closeLog;
  $("#clearLog").onclick = () => {
    $("#log").innerHTML = "";
    state.unread = 0;
    $("#logBadge").hidden = true;
    $("#fabBadge").hidden = true;
  };
  $("#infoClose").onclick = () => ($("#infoModal").hidden = true);

  $("#addFiles").onclick = addFiles;
  $("#addFolder").onclick = addFolder;
  $("#addServerDir").onclick = () => openFs("files");
  $("#clearFiles").onclick = () => {
    state.inputs = [];
    renderFiles();
  };
  $("#pickOutDir").onclick = async () => {
    const r = await withBusy($("#pickOutDir"), () => pickNative("dir"));
    if (r && r.ok) {
      if (r.paths && r.paths.length) {
        $("#outputDir").value = r.paths[0];
        api("/api/config", { method: "POST", body: JSON.stringify({ output_dir: r.paths[0] }) });
      }
      return;
    }
    if (r) return;
    openFs("outdir");
  };

  $("#openOutDir").onclick = () => openDir();
  $("#openResultDir").onclick = () => openDir(state.lastOutput);

  $("#fsClose").onclick = () => ($("#fsModal").hidden = true);
  $("#fsUp").onclick = () => browse(state.fsParent);
  $("#fsGo").onclick = () => browse($("#fsPath").value);
  $("#fsPick").onclick = () => {
    if (state.fsMode === "outdir") {
      $("#outputDir").value = state.fsPath;
      api("/api/config", { method: "POST", body: JSON.stringify({ output_dir: state.fsPath }) });
    } else {
      addPaths([state.fsPath]);
    }
    $("#fsModal").hidden = true;
  };

  $("#runBtn").onclick = run;
  $("#cancelBtn").onclick = async () => {
    if (state.taskId) {
      await api("/api/cancel", { method: "POST", body: JSON.stringify({ id: state.taskId }) });
      log("已请求取消…", "warn");
    }
  };

  $("#saveProfile").onclick = async () => {
    const name = $("#profileName").value.trim();
    if (!name) return toast("请先填方案名", "warn");
    await api("/api/profile/save", {
      method: "POST",
      body: JSON.stringify({
        name,
        data: { tool: state.tool ? state.tool.id : "", inputs: state.inputs, params: state.params },
      }),
    });
    log("方案已保存: " + name, "ok");
    loadProfiles();
  };

  $("#profileSelect").onchange = async (e) => {
    const name = e.target.value;
    if (!name) return;
    const r = await api("/api/profile?name=" + encodeURIComponent(name));
    if (!r.ok) return log("方案载入失败", "err");
    const d = r.data;
    const tool = state.tools.find((t) => t.id === d.tool);
    if (tool) selectTool(tool);
    state.inputs = d.inputs || [];
    renderFiles();
    Object.entries(d.params || {}).forEach(([k, v]) => {
      const el = document.getElementById("p_" + k);
      if (el) {
        if (el.type === "checkbox") el.checked = !!v;
        else el.value = v;
        state.params[k] = v;
      }
    });
    log("方案已载入: " + name, "ok");
  };

  async function loadProfiles() {
    const r = await api("/api/profiles");
    const sel = $("#profileSelect");
    sel.innerHTML = '<option value="">载入方案…</option>';
    (r.profiles || []).forEach((p) => {
      const o = document.createElement("option");
      o.value = p;
      o.textContent = p;
      sel.appendChild(o);
    });
  }
  loadProfiles();

  api("/api/config").then((r) => {
    if (r.ok && r.output_dir) $("#outputDir").value = r.output_dir;
  });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      $("#fsModal").hidden = true;
      $("#infoModal").hidden = true;
    }
  });
});

/* ---------- 运行与轮询 ---------- */
async function run() {
  if (!state.tool) return toast("请先选择一个功能", "warn");
  if (!state.inputs.length) return toast("请先添加文件或文件夹", "warn");
  $("#log").innerHTML = "";
  state.logCount = 0;
  $("#resultBar").hidden = true;
  openLog();
  const outDir = $("#outputDir").value.trim() || "output";
  const params = Object.assign({}, state.params);
  if ($("#dryRun").checked) params.__dry_run__ = true;
  const r = await api("/api/run", {
    method: "POST",
    body: JSON.stringify({
      tool: state.tool.id,
      inputs: state.inputs,
      output_dir: outDir,
      params,
    }),
  });
  if (!r.ok) {
    log("启动失败: " + (r.error || ""), "err");
    return;
  }
  state.taskId = r.task_id;
  $("#runBtn").disabled = true;
  $("#cancelBtn").disabled = false;
  setProgress(0, 0, "已启动");
  poll();
}

function setProgress(cur, total, name) {
  const pct = total > 0 ? Math.round((cur / total) * 100) : 0;
  $("#progressFill").style.width = pct + "%";
  $("#progressText").textContent = total > 0
    ? `进度 ${cur}/${total} ${name ? "· " + name : ""}`
    : (name || "准备中");
}

async function poll() {
  if (!state.taskId) return;
  const r = await api(`/api/task?id=${state.taskId}&since=${state.logCount}`);
  if (!r.ok) return;
  const t = r.task;
  setProgress(t.progress.current, t.progress.total, t.progress.name);
  (t.logs || []).forEach((line) => {
    const kind = line.includes("[错误]") ? "err" : line.includes("[警告]") ? "warn" : line.includes("[成功]") ? "ok" : "";
    log(line, kind);
  });
  state.logCount = t.log_count || state.logCount;

  if (["完成", "失败", "已取消"].includes(t.status)) {
    $("#runBtn").disabled = false;
    $("#cancelBtn").disabled = true;
    if (t.result) {
      log(`结果：成功 ${t.result.success_count}，失败 ${t.result.fail_count}`, t.result.fail_count ? "warn" : "ok");
      (t.result.output_files || []).slice(0, 20).forEach((f) => log("输出: " + f));
    }
    const outs = (t.result && t.result.output_files) || [];
    state.lastOutput = outs[0] || null;
    if (outs.length) {
      $("#resultText").textContent = `已生成 ${outs.length} 个文件`;
      $("#resultBar").hidden = false;
    } else {
      $("#resultBar").hidden = true;
    }
    log("状态: " + t.status, t.status === "失败" ? "err" : "ok");
    toast(
      t.status + "：成功 " + ((t.result && t.result.success_count) || 0) +
        "，失败 " + ((t.result && t.result.fail_count) || 0),
      t.status === "失败" ? "err" : "ok"
    );
    state.taskId = null;
    return;
  }
  state.polling = setTimeout(poll, 700);
}
