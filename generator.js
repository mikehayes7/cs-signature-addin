// generator.js
// Standalone Cellular Sales email signature builder.
// No Office.js dependency -- works as a plain static page (e.g. on GitHub Pages).
//
// CHANGE LOG (this pass):
// 1. ASSET_BASE switched from a relative path to an absolute URL. Relative
//    paths only resolve correctly while the images are viewed *on this page*.
//    Once the signature HTML is copied into Outlook's signature editor, it no
//    longer lives at this URL, so relative paths broke -> Outlook showed the
//    logo/icon spots as grey placeholder boxes. Absolute URLs fix that.
//    If GitHub Pages is not the final host, update ASSET_BASE to wherever
//    IT ends up hosting the assets folder.
// 2. Sales Rep template's title is now hard-set to "Sales Representative" --
//    the title field is disabled and ignored for that template.
// 3. copySignature() now cascades through multiple copy strategies so it
//    degrades gracefully on mobile browsers that don't support rich
//    clipboard writes: rich HTML -> legacy rich copy -> plain-text clipboard
//    -> plain-text legacy copy -> manual "select this box" fallback.

const ASSET_BASE = "https://mikehayes7.github.io/cs-signature-addin/assets";

const BASE_STYLE = "font-family: Arial, sans-serif; font-size: 12px; color: #000000; line-height: 1.4;";

const HARD_SET_TITLES = {
  salesRep: "Sales Representative"
};

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

function applyHardSetTitle(templateKey) {
  const titleInput = document.getElementById("titleInput");
  const hardTitle = HARD_SET_TITLES[templateKey];
  if (hardTitle) {
    titleInput.value = hardTitle;
    titleInput.disabled = true;
    titleInput.dataset.autoset = "true";
    titleInput.title = "This template always uses a fixed title.";
  } else {
    titleInput.disabled = false;
    titleInput.title = "";
    // Only clear if the current value was auto-filled by us (e.g. leftover
    // "Sales Representative" from switching away from that template) --
    // never wipe out something the user typed themselves.
    if (titleInput.dataset.autoset === "true") {
      titleInput.value = "";
    }
    titleInput.dataset.autoset = "false";
  }
}

function renderSignature(templateKey, fields) {
  const template = TEMPLATES[templateKey];
  if (!template) throw new Error("Unknown template: " + templateKey);

  const title = HARD_SET_TITLES[templateKey] || fields.title || "";

  let html = template.html;
  html = html.replaceAll("{{name}}", escapeHtml(fields.name || ""));
  html = html.replaceAll("{{title}}", escapeHtml(title));
  html = html.replaceAll("{{workPhone}}", escapeHtml(fields.workPhone || ""));
  html = html.replaceAll("{{cellPhone}}", escapeHtml(fields.cellPhone || ""));
  return html;
}

// Simplified, image-free layout for Outlook Mobile's signature editor, which
// doesn't reliably support images/rich HTML the way desktop Outlook does.
function renderMobileSignature(templateKey, fields) {
  const title = HARD_SET_TITLES[templateKey] || fields.title || "";
  const website = templateKey === "recruiter" ? "cellularsales.com/careers" : "cellularsales.com";

  return `
<table style="${BASE_STYLE}" cellpadding="0" cellspacing="0" border="0">
  <tr><td style="font-weight:bold;">${escapeHtml(fields.name || "")}</td></tr>
  <tr><td>${escapeHtml(title)}</td></tr>
  <tr><td>&nbsp;</td></tr>
  <tr><td>9040 Executive Park Dr.</td></tr>
  <tr><td>Knoxville, TN 37923</td></tr>
  <tr><td><b>w.</b> ${escapeHtml(fields.workPhone || "")}</td></tr>
  <tr><td><b>c.</b> ${escapeHtml(fields.cellPhone || "")}</td></tr>
  <tr><td>${website}</td></tr>
  <tr><td><a href="https://cellularsales.com/careers" style="color:#c00; text-decoration:underline;">Work with us</a></td></tr>
</table>`;
}

function updatePreview() {
  const templateKey = document.getElementById("templateSelect").value;
  const isMobile = document.getElementById("mobileToggle").checked;
  applyHardSetTitle(templateKey);
  const fields = getFields();
  const html = isMobile
    ? renderMobileSignature(templateKey, fields)
    : renderSignature(templateKey, fields);
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

function htmlToPlainText(html) {
  const plain = document.createElement("div");
  plain.innerHTML = html;
  return plain.textContent.replace(/\n\s*\n+/g, "\n").trim();
}

// Legacy rich-copy: render the HTML into a real, in-viewport (but invisible)
// contenteditable element, select it, and use execCommand("copy"). Some
// mobile browsers refuse to copy content that's positioned off-screen, so
// this keeps the holder within the visible viewport using opacity instead.
function legacyCopyHtml(html) {
  const holder = document.createElement("div");
  holder.setAttribute("contenteditable", "true");
  holder.style.position = "fixed";
  holder.style.top = "0";
  holder.style.left = "0";
  holder.style.opacity = "0";
  holder.style.pointerEvents = "none";
  holder.style.zIndex = "-1";
  holder.innerHTML = html;
  document.body.appendChild(holder);

  let ok = false;
  try {
    const range = document.createRange();
    range.selectNodeContents(holder);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    ok = document.execCommand("copy");
    selection.removeAllRanges();
  } catch (err) {
    ok = false;
  } finally {
    document.body.removeChild(holder);
  }
  return ok;
}

function showManualCopyFallback(html) {
  const frame = document.getElementById("previewFrame");
  frame.scrollIntoView({ behavior: "smooth", block: "center" });
  showStatus(
    "Couldn't copy automatically on this device. Tap and hold inside the preview above, choose \"Select All\", then \"Copy\".",
    true
  );
}

async function copySignature() {
  const fields = getFields();
  const templateKey = document.getElementById("templateSelect").value;
  const isMobile = document.getElementById("mobileToggle").checked;
  const effectiveTitle = HARD_SET_TITLES[templateKey] || fields.title;

  if (!fields.name || !effectiveTitle) {
    showStatus("Please fill in at least your name and title before copying.", true);
    return;
  }

  const html = isMobile
    ? renderMobileSignature(templateKey, fields)
    : renderSignature(templateKey, fields);
  const text = htmlToPlainText(html);

  // 1) Preferred: rich clipboard write (keeps formatting when pasted into Outlook/Gmail).
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
      // Fall through -- common on mobile Safari/Chrome, which either lack
      // ClipboardItem support or block it outside a direct user gesture.
    }
  }

  // 2) Legacy rich copy via execCommand.
  if (legacyCopyHtml(html)) {
    showStatus("Signature copied! Paste it into your email client's signature editor.", false);
    return;
  }

  // 3) Plain-text clipboard write -- widely supported, including most mobile browsers.
  if (navigator.clipboard && navigator.clipboard.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      showStatus(
        "Formatting couldn't be copied on this device -- copied plain text instead. You'll need to add logos/formatting manually, or set this up from a desktop browser.",
        true
      );
      return;
    } catch (err) {
      // Fall through to the final fallback.
    }
  }

  // 4) Plain-text legacy copy.
  try {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    textarea.style.top = "0";
    textarea.style.left = "0";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(textarea);
    if (ok) {
      showStatus(
        "Formatting couldn't be copied on this device -- copied plain text instead. You'll need to add logos/formatting manually, or set this up from a desktop browser.",
        true
      );
      return;
    }
  } catch (err) {
    // Fall through to manual fallback.
  }

  // 5) Nothing worked -- point the user at manual copy.
  showManualCopyFallback(html);
}

function resetForm() {
  document.getElementById("templateSelect").value = "salesRep";
  document.getElementById("nameInput").value = "";
  document.getElementById("titleInput").value = "";
  document.getElementById("workPhoneInput").value = "";
  document.getElementById("cellPhoneInput").value = "";
  document.getElementById("mobileToggle").checked = false;
  showStatus("", false);
  updatePreview();
}

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("templateSelect").addEventListener("change", updatePreview);
  document.getElementById("mobileToggle").addEventListener("change", updatePreview);
  document.getElementById("refreshBtn").addEventListener("click", resetForm);
  document.getElementById("copyBtn").addEventListener("click", copySignature);
  document.getElementById("titleInput").addEventListener("input", () => {
    document.getElementById("titleInput").dataset.autoset = "false";
  });
  ["nameInput", "titleInput", "workPhoneInput", "cellPhoneInput"].forEach((id) => {
    document.getElementById(id).addEventListener("input", updatePreview);
  });

  updatePreview();
});
