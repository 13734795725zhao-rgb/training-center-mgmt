/* ========== 前端公共组件 v2.0 ========== */
/* 提供：Toast 通知、自定义模态框、fetch 封装、表单校验、防重复提交 */

/* ---------- Toast 通知 ---------- */
const Toast = {
  container: null,
  init() {
    if (this.container) return;
    this.container = document.createElement("div");
    this.container.className = "toast-container";
    document.body.appendChild(this.container);
  },
  show(message, type = "info", duration = 3000) {
    this.init();
    const toast = document.createElement("div");
    toast.className = `toast toast-${type}`;
    const icons = { success: "\u2713", error: "\u2717", warning: "\u26a0", info: "i" };
    toast.innerHTML = `<span class="toast-icon">${icons[type] || "i"}</span><span class="toast-msg">${this.escape(message)}</span>`;
    this.container.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add("show"));
    setTimeout(() => {
      toast.classList.remove("show");
      setTimeout(() => toast.remove(), 300);
    }, duration);
  },
  success(msg, dur) { this.show(msg, "success", dur); },
  error(msg, dur) { this.show(msg, "error", dur || 5000); },
  warning(msg, dur) { this.show(msg, "warning", dur); },
  info(msg, dur) { this.show(msg, "info", dur); },
  escape(s) {
    const d = document.createElement("div");
    d.textContent = s;
    return d.innerHTML;
  },
};

/* ---------- 自定义模态框（替代 alert/confirm） ---------- */
const Modal = {
  show({ title, body, buttons }) {
    return new Promise((resolve) => {
      const overlay = document.createElement("div");
      overlay.className = "modal-overlay";
      const modal = document.createElement("div");
      modal.className = "modal-box";
      let btnHtml = "";
      (buttons || [{ label: "\u786e\u5b9a", value: true, type: "primary" }]).forEach(b => {
        btnHtml += `<button class="modal-btn modal-btn-${b.type || "default"}" data-value="${b.value}">${this.escape(b.label)}</button>`;
      });
      modal.innerHTML = `
        <div class="modal-header">
          <span class="modal-title">${this.escape(title || "\u63d0\u793a")}</span>
          <button class="modal-close">&times;</button>
        </div>
        <div class="modal-body">${body || ""}</div>
        <div class="modal-footer">${btnHtml}</div>
      `;
      overlay.appendChild(modal);
      document.body.appendChild(overlay);
      requestAnimationFrame(() => overlay.classList.add("show"));

      const close = (val) => {
        overlay.classList.remove("show");
        setTimeout(() => overlay.remove(), 200);
        resolve(val);
      };
      modal.querySelector(".modal-close").onclick = () => close(null);
      overlay.onclick = (e) => { if (e.target === overlay) close(null); };
      modal.querySelectorAll(".modal-btn").forEach(btn => {
        btn.onclick = () => close(btn.dataset.value);
      });
    });
  },
  alert(message, title) {
    return this.show({ title: title || "\u63d0\u793a", body: `<p>${this.escape(message)}</p>` });
  },
  confirm(message, title) {
    return this.show({
      title: title || "\u786e\u8ba4",
      body: `<p>${this.escape(message)}</p>`,
      buttons: [
        { label: "\u53d6\u6d88", value: false, type: "default" },
        { label: "\u786e\u8ba4", value: true, type: "danger" },
      ],
    });
  },
  escape(s) {
    const d = document.createElement("div");
    d.textContent = s;
    return d.innerHTML;
  },
};

/* ---------- fetch 封装（自动处理 401/403 + JSON） ---------- */
async function apiFetch(url, options = {}) {
  const opts = { ...options };
  opts.headers = opts.headers || {};
  if (opts.body && !(opts.body instanceof FormData) && !opts.headers["Content-Type"]) {
    opts.headers["Content-Type"] = "application/json";
  }
  if (opts.body && typeof opts.body === "object" && !(opts.body instanceof FormData)) {
    opts.body = JSON.stringify(opts.body);
  }
  try {
    const res = await fetch(url, opts);
    if (res.status === 401) {
      const data = await res.json().catch(() => ({}));
      if (data.redirect) {
        window.location.href = data.redirect;
        return null;
      }
      Toast.error("\u672a\u767b\u5f55\u6216\u767b\u5f55\u5df2\u8fc7\u671f");
      setTimeout(() => window.location.href = "/login", 1500);
      return null;
    }
    if (res.status === 403) {
      const data = await res.json().catch(() => ({}));
      Toast.error(data.error || "\u6743\u9650\u4e0d\u8db3");
      return null;
    }
    return res;
  } catch (e) {
    Toast.error("\u7f51\u7edc\u8fde\u63a5\u5f02\u5e38");
    return null;
  }
}

/* ---------- HTML 转义 ---------- */
function escapeHtml(s) {
  if (s === null || s === undefined) return "";
  const d = document.createElement("div");
  d.textContent = String(s);
  return d.innerHTML;
}

/* ---------- 防重复提交 ---------- */
function withLock(btn, fn) {
  if (btn.dataset.locked === "1") return Promise.resolve(null);
  btn.dataset.locked = "1";
  btn.disabled = true;
  const orig = btn.textContent;
  btn.textContent = "\u5904\u7406\u4e2d...";
  return Promise.resolve(fn()).finally(() => {
    btn.dataset.locked = "0";
    btn.disabled = false;
    btn.textContent = orig;
  });
}

/* ---------- 表单校验 ---------- */
function validateField(input, rules) {
  const val = (input.value || "").trim();
  for (const rule of rules) {
    if (rule.required && !val) return rule.message || "\u6b64\u5b57\u6bb5\u4e3a\u5fc5\u586b\u9879";
    if (rule.minlength && val.length < rule.minlength) return rule.message || `\u81f3\u5c11 ${rule.minlength} \u4e2a\u5b57\u7b26`;
    if (rule.maxlength && val.length > rule.maxlength) return rule.message || `\u6700\u591a ${rule.maxlength} \u4e2a\u5b57\u7b26`;
    if (rule.pattern && !rule.pattern.test(val)) return rule.message || "\u683c\u5f0f\u4e0d\u6b63\u786e";
    if (rule.min !== undefined && parseFloat(val) < rule.min) return rule.message || `\u4e0d\u80fd\u5c0f\u4e8e ${rule.min}`;
    if (rule.max !== undefined && parseFloat(val) > rule.max) return rule.message || `\u4e0d\u80fd\u5927\u4e8e ${rule.max}`;
  }
  return null;
}

/* ---------- 用户菜单 ---------- */
function initUserMenu() {
  const menu = document.querySelector(".user-menu");
  if (!menu) return;
  const trigger = menu.querySelector(".user-trigger");
  const dropdown = menu.querySelector(".user-dropdown");
  if (!trigger || !dropdown) return;
  trigger.onclick = (e) => {
    e.stopPropagation();
    dropdown.classList.toggle("show");
  };
  document.addEventListener("click", () => dropdown.classList.remove("show"));
}

document.addEventListener("DOMContentLoaded", initUserMenu);
