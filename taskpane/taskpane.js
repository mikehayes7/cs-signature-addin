// taskpane.js

Office.onReady(() => {
  document.getElementById("previewBtn").addEventListener("click", updatePreview);
  document.getElementById("insertBtn").addEventListener("click", insertSignature);
  document.getElementById("templateSelect").addEventListener("change", updatePreview);

  loadSavedFields();
  updatePreview();
});

function getFields() {
  return {
    name: document.getElementById("nameInput").value.trim(),
    title: document.getElementById("titleInput").value.trim(),
    workPhone: document.getElementById("workPhoneInput").value.trim(),
    cellPhone: document.getElementById("cellPhoneInput").value.trim()
  };
}

function updatePreview() {
  const templateKey = document.getElementById("templateSelect").value;
  const html = renderSignature(templateKey, getFields());
  const frame = document.getElementById("previewFrame");
  const doc = frame.contentDocument || frame.contentWindow.document;
  doc.open();
  doc.write(html);
  doc.close();
}

function insertSignature() {
  const templateKey = document.getElementById("templateSelect").value;
  const fields = getFields();

  if (!fields.name || !fields.title) {
    showStatus("Please fill in at least your name and title.", true);
    return;
  }

  const html = renderSignature(templateKey, fields);

  Office.context.mailbox.item.body.setSignatureAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        saveFields(fields);
        showStatus("Signature inserted.", false);
      } else {
        showStatus("Could not insert signature: " + asyncResult.error.message, true);
      }
    }
  );
}

function showStatus(message, isError) {
  const el = document.getElementById("statusMsg");
  el.textContent = message;
  el.className = isError ? "error" : "";
}

function saveFields(fields) {
  const settings = Office.context.roamingSettings;
  settings.set("cs_signature_fields", fields);
  settings.saveAsync();
}

function loadSavedFields() {
  const settings = Office.context.roamingSettings;
  const saved = settings.get("cs_signature_fields");
  if (saved) {
    document.getElementById("nameInput").value = saved.name || "";
    document.getElementById("titleInput").value = saved.title || "";
    document.getElementById("workPhoneInput").value = saved.workPhone || "";
    document.getElementById("cellPhoneInput").value = saved.cellPhone || "";
  }
}
