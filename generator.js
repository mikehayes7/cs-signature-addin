// generator.js
// Standalone Cellular Sales email signature builder.
// No Office.js dependency -- works as a plain static page (e.g. on GitHub Pages).
//
// IMPORTANT: Replace ASSET_BASE below once logo/icon/banner images are hosted
// (add them under an /assets folder in this repo, or point at another host).
// SVG is NOT supported by most email clients' signature fields -- use PNG or JPG only.

const ASSET_BASE = "assets";

const BASE_STYLE = "font-family: Arial, sans-serif; font-size: 12px; color: #000000; line-height: 1.4;";

function baseSignature({ showCareersLink = false, showAwardBanner = false } = {}) {
  const websiteLine = showCareersLink ? "cellularsales.com/careers" : "cellularsales.com";

  return `
<table style="${BASE_STYLE}" cellpadding="0" cellspacing="0" border="0">
  <tr><td style="font-weight:bold;">{{name}}</td></tr>
  <tr><td>{{title}}</td></tr>
  <tr><td>&nbsp;</td></tr>
  <tr><td>9040 Executive Park Dr.</td></tr>
  <tr><td>Knoxville, TN 37923</td></tr>
  <tr><td><b>w.</b> {{workPhone}}</td></tr>
  <tr><td><b>c.</b> {{cellPhone}}</td></tr>
  <tr><td>${websiteLine}</td></tr>
  <tr><td><a href="https://cellularsales.com/careers" style="color:#c00; text-decoration:underline;">Work with us</a></td></tr>
  <tr><td>&nbsp;</td></tr>
  <tr><td style="font-weight:bold;">great ${showCareersLink ? '<span style="color:#c00;">career</span> ' : ''}experiences start here.</td></tr>
  <tr><td>&nbsp;</td></tr>
  <tr>
    <td>
      <img src="${ASSET_BASE}/cellularsales-logo.png" alt="cellularsales" height="24" style="vertical-align:middle;" />
      &nbsp;&nbsp;
      <img src="${ASSET_BASE}/verizon-authorized-retailer.png" alt="Verizon Authorized Retailer" height="24" style="vertical-align:middle;" />
    </td>
  </tr>
  <tr><td>&nbsp;</td></tr>
  <tr>
    <td>
      <a href="https://instagram.com/cellularsales"><img src="${ASSET_BASE}/icon-instagram.png" width="16" height="16" alt="Instagram" style="margin-right:6px;" /></a>
      <a href="https://tiktok.com/@cellularsales"><img src="${ASSET_BASE}/icon-tiktok.png" width="16" height="16" alt="TikTok" style="margin-right:6px;" /></a>
      <a href="https://x.com/cellularsales"><img src="${ASSET_BASE}/icon-x.png" width="16" height="16" alt="X" style="margin-right:6px;" /></a>
      <a href="https://facebook.com/cellularsales"><img src="${ASSET_BASE}/icon-facebook.png" width="16" height="16" alt="Facebook" style="margin-right:6px;" /></a>
      <a href="https://linkedin.com/company/cellularsales"><img src="${ASSET_BASE}/icon-linkedin.png" width="16" height="16" alt="LinkedIn" /></a>
    </td>
  </tr>
  ${showAwardBanner ? `
  <tr><td>&nbsp;</td></tr>
  <tr><td><img src="${ASSET_BASE}/greatest-workplaces-2026.png" alt="America's Greatest Workplaces 2026 - Newsweek" height="70" /></td></tr>
  ` : ''}
</table>`;
}

const TEMPLATES = {
  salesRep: {
    label: "Sales Rep Signature",
    html: baseSignature({ showCareersLink: false, showAwardBanner: false })
  },
  recruiter: {
    label: "Recruiter Signature",
    html: baseSignature({ showCareersLink: true, showAwardBanner: true })
  },
  hqMarketOffice: {
    label: "HQ / Recruiter / Market Office Employee Signature",
    html: baseSignature({ showCareersLink: false, showAwardBanner: false })
  }
};

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function getFields() {
  return {
    name: document.getElementById("nameInput").value.trim(),
    title: document.getElementById("titleInput").value.trim(),
    workPhone: document.getElementById("workPhoneInput").value.trim(),
    cellPhone: document.getElementById("cellPhoneInput").value.trim()
  };
}

function renderSignature(templateKey, fields) {
  const template = TEMPLATES[templateKey];
  if (!template) throw new Error("Unknown template: " + templateKey);

  let html = template.html;
  html = html.replaceAll("{{name}}", escapeHtml(fields.name || ""));
  html = html.replaceAll("{{title}}", escapeHtml(fields.title || ""));
  html = html.replaceAll("{{workPhone}}", escapeHtml(fields.workPhone || ""));
  html = html.replaceAll("{{cellPhone}}", escapeHtml(fields.cellPhone || ""));
  return html;
}

function updatePreview() {
  const templateKey = document.getElementById("templateSelect").value;
  const html = renderSignature(templateKey, getFields());
  const frame = document.getElementById("previewFrame");
  // Use srcdoc (an attribute on the iframe element itself) rather than
  // reaching into frame.contentDocument -- contentDocument access can be
  // blocked when this page is displayed inside a sandboxed iframe, like
  // SharePoint's Embed web part. srcdoc works regardless.
  frame.srcdoc = html;
  return html;
}

function showStatus(message, isError) {
  const el = document.getElementById("statusMsg");
  el.textContent = message;
  el.className = isError ? "error" : "";
}

async function copySignature() {
  const fields = getFields();
  if (!fields.name || !fields.title) {
    showStatus("Please fill in at least your name and title before copying.", true);
    return;
  }

  const templateKey = document.getElementById("templateSelect").value;
  const html = renderSignature(templateKey, fields);
  const plain = document.createElement("div");
  plain.innerHTML = html;
  const text = plain.textContent.replace(/\n\s*\n+/g, "\n").trim();

  // Preferred: rich clipboard write (keeps formatting when pasted into Outlook/Gmail).
  if (navigator.clipboard && window.ClipboardItem) {
    try {
      const item = new ClipboardItem({
        "text/html": new Blob([html], { type: "text/html" }),
        "text/plain": new Blob([text], { type: "text/plain" })
      });
      await navigator.clipboard.write([item]);
      showStatus("Signature copied! Paste it into your email client's signature editor.", false);
      return;
    } catch (err) {
      // Fall through to legacy copy method below.
    }
  }

  // Fallback: select a hidden rendered copy and use execCommand("copy").
  try {
    const holder = document.createElement("div");
    holder.setAttribute("contenteditable", "true");
    holder.style.position = "fixed";
    holder.style.left = "-9999px";
    holder.innerHTML = html;
    document.body.appendChild(holder);

    const range = document.createRange();
    range.selectNodeContents(holder);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);

    const ok = document.execCommand("copy");
    selection.removeAllRanges();
    document.body.removeChild(holder);

    if (ok) {
      showStatus("Signature copied! Paste it into your email client's signature editor.", false);
    } else {
      showStatus("Copy failed -- select the preview above and copy manually (Ctrl/Cmd+C).", true);
    }
  } catch (err) {
    showStatus("Copy failed -- select the preview above and copy manually (Ctrl/Cmd+C).", true);
  }
}

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("templateSelect").addEventListener("change", updatePreview);
  document.getElementById("refreshBtn").addEventListener("click", updatePreview);
  document.getElementById("copyBtn").addEventListener("click", copySignature);
  ["nameInput", "titleInput", "workPhoneInput", "cellPhoneInput"].forEach((id) => {
    document.getElementById(id).addEventListener("input", updatePreview);
  });

  updatePreview();
});
