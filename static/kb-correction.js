const kbForm = document.querySelector("#kb-correction-form");
const kbButton = document.querySelector("#kb-correction-button");
const kbResults = document.querySelector("#kb-correction-results");
const kbStatusBox = document.querySelector("#kb-correction-status");
const kbChangesPanel = document.querySelector("#kb-correction-changes-panel");
const kbChangesList = document.querySelector("#kb-correction-changes-list");
const kbNoChanges = document.querySelector("#kb-correction-no-changes");

let kbEditor;
let kbInputEditor;

const KB_MAX_IMAGE_BYTES = 10 * 1024 * 1024;

function kbBase64UploadAdapterPlugin(editor) {
  editor.plugins.get("FileRepository").createUploadAdapter = (loader) => ({
    upload() {
      return new Promise((resolve, reject) => {
        loader.file.then((file) => {
          if (file.size > KB_MAX_IMAGE_BYTES) {
            reject(`Gambar "${file.name}" terlalu besar (${(file.size / (1024 * 1024)).toFixed(1)} MB). Maksimal 10 MB per gambar.`);
            return;
          }
          const reader = new FileReader();
          reader.onload = () => resolve({ default: reader.result });
          reader.onerror = () => reject(reader.error);
          reader.readAsDataURL(file);
        });
      });
    },
  });
}

async function kbGetInputEditor() {
  if (!kbInputEditor) {
    kbInputEditor = await ClassicEditor.create(document.querySelector("#kb-correction-input-editor"), {
      toolbar: [
        "heading", "|", "bold", "italic", "bulletedList", "numberedList", "insertTable", "|",
        "uploadImage", "|", "undo", "redo",
      ],
      extraPlugins: [kbBase64UploadAdapterPlugin],
    });
  }
  return kbInputEditor;
}
kbGetInputEditor();

function kbSetStatus(message, error = false) {
  kbStatusBox.textContent = message;
  kbStatusBox.className = error ? "error" : "status";
}

function kbRenderChanges(changes) {
  kbChangesList.replaceChildren();
  if (!changes || changes.length === 0) {
    kbNoChanges.hidden = false;
  } else {
    kbNoChanges.hidden = true;
    changes.forEach((change) => {
      const li = document.createElement("li");
      li.textContent = change;
      kbChangesList.append(li);
    });
  }
  kbChangesPanel.hidden = false;
}

kbForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  kbButton.disabled = true;
  kbResults.hidden = true;
  kbChangesPanel.hidden = true;

  const inputEditor = await kbGetInputEditor();
  const htmlFragment = inputEditor.getData().trim();
  if (!htmlFragment) {
    kbButton.disabled = false;
    kbSetStatus("Tulis atau tempel konten yang ingin dikoreksi.", true);
    return;
  }

  // Cek tiap gambar base64 inline (menjangkau semua jalur insert: upload,
  // paste langsung, atau drag & drop — paste HTML dari Word/PPT bisa membawa
  // gambar base64 yang sudah menempel dan tidak lewat upload adapter).
  const imageMatches = [...htmlFragment.matchAll(/<img[^>]*\bsrc="data:image\/[a-zA-Z]+;base64,([A-Za-z0-9+/=]+)"[^>]*>/g)];
  for (const match of imageMatches) {
    const sizeMb = (match[1].length * 0.75) / (1024 * 1024);
    if (sizeMb > 10) {
      kbButton.disabled = false;
      kbSetStatus(`Ada gambar berukuran ${sizeMb.toFixed(1)} MB di konten (maksimal 10 MB per gambar) — hapus atau kompres gambar itu dulu.`, true);
      return;
    }
  }
  if (imageMatches.length > 6) {
    kbButton.disabled = false;
    kbSetStatus(`Konten berisi ${imageMatches.length} gambar, maksimal 6 gambar per koreksi.`, true);
    return;
  }

  const htmlSizeMb = new Blob([htmlFragment]).size / (1024 * 1024);
  if (htmlSizeMb > 100) {
    kbButton.disabled = false;
    kbSetStatus(`Konten terlalu besar (${htmlSizeMb.toFixed(0)} MB, maksimal 100 MB) — kurangi jumlah/ukuran gambar.`, true);
    return;
  }

  kbSetStatus("Mencari materi relevan di Knowledge Base dan mengoreksi...");
  try {
    const formData = new FormData(kbForm);
    formData.set("html_fragment", htmlFragment);
    const response = await fetch("/api/correct-from-kb", { method: "POST", body: formData });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Gagal mengoreksi HTML.");

    if (kbEditor) {
      await kbEditor.setData(data.corrected_html || "");
    } else {
      kbEditor = await ClassicEditor.create(document.querySelector("#kb-correction-editor"), {
        toolbar: ["heading", "|", "bold", "italic", "bulletedList", "numberedList", "insertTable", "undo", "redo"],
      });
      await kbEditor.setData(data.corrected_html || "");
    }
    kbRenderChanges(data.changes);
    kbResults.hidden = false;
    kbSetStatus("Koreksi selesai. Tinjau dan edit hasil sebelum digunakan.");
  } catch (error) {
    kbSetStatus(error.message, true);
  } finally {
    kbButton.disabled = false;
  }
});

document.querySelector("#kb-correction-download").addEventListener("click", async () => {
  const content = kbEditor ? await kbEditor.getData() : "";
  const html = (
    "<!DOCTYPE html>\n" +
    '<html lang="id"><head><meta charset="utf-8">\n' +
    "<style>\n" +
    "table { width: 100%; table-layout: fixed; border-collapse: collapse; }\n" +
    "td, th { border: 1px solid #ccc; padding: 8px; vertical-align: top; word-wrap: break-word; }\n" +
    "img { max-width: 100%; height: auto; display: block; }\n" +
    "</style></head><body>\n" +
    content +
    "\n</body></html>"
  );
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "dokumen-terkoreksi.html";
  link.click();
  URL.revokeObjectURL(link.href);
});
