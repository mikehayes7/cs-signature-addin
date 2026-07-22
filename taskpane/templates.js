// templates.js
// Signature HTML templates, matching the layouts on the "Email Signature Set Up" SharePoint page.
// {{tokens}} get replaced with the values the user enters in the task pane.
//
// IMPORTANT: Replace the placeholder image URLs below with your actual hosted logo/banner assets.
// SVG is NOT supported in Outlook signatures — use PNG or JPG only.

const ASSET_BASE = "https://YOUR-HOSTING-DOMAIN/assets";

const BASE_STYLE = `font-family: Arial, sans-serif; font-size: 12px; color: #000000; line-height: 1.4;`;

function baseSignature({ showCareersLink = false, showAwardBanner = false } = {}) {
  const websiteLine = showCareersLink
    ? `cellularsales.com/careers`
    : `cellularsales.com`;

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

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}
