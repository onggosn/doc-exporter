const partForm = document.querySelector("#part-generate-form");
const partReferences = document.querySelector("#part-references");
const partAddReference = document.querySelector("#part-add-reference");
const partStatusBox = document.querySelector("#part-status");
const partGenerateButton = document.querySelector("#part-generate-button");
const partResults = document.querySelector("#part-results");
const partTabs = document.querySelector("#part-chapter-tabs");
const partResultSource = document.querySelector("#part-result-source");
const partProgressPanel = document.querySelector("#part-generate-progress-panel");
const partProgressFill = document.querySelector("#part-generate-progress-fill");
const partProgressPercent = document.querySelector("#part-generate-progress-percent");
const partElapsedTime = document.querySelector("#part-generate-elapsed-time");
const partStepList = document.querySelector("#part-generate-step-list");
const partResultElapsed = document.querySelector("#part-result-elapsed");
const correctSelectionButton = document.querySelector("#correct-selection-button");

let partEditor;
let partChapterResults = [];
let partActiveChapter = 0;
let partDocumentId = null;

function partSetStatus(message, error = false) {
  partStatusBox.textContent = message;
  partStatusBox.className = error ? "error" : "status";
}

function partFormatElapsed(seconds) {
  const total = Math.round(seconds);
  const minutes = Math.floor(total / 60);
  const secs = total % 60;
  return minutes > 0 ? `${minutes}m ${secs}s` : `${secs}s`;
}

function partAddReferencePair() {
  if (partReferences.children.length >= 2) return;
  const row = document.createElement("div");
  row.className = "reference-row";
  row.innerHTML = `
    <label class="field">
      <span>PPT contoh</span>
      <input name="reference_ppt" type="file" accept=".pptx">
    </label>
    <label class="field">
      <span>Dokumen final (.html atau .txt)</span>
      <input name="reference_doc" type="file" accept=".html,.htm,.txt">
    </label>
    <button type="button" class="remove-reference" aria-label="Hapus pasangan contoh">Hapus</button>`;
  row.querySelector(".remove-reference").addEventListener("click", () => {
    row.remove();
    partAddReference.disabled = partReferences.children.length >= 2;
  });
  partReferences.append(row);
  partAddReference.disabled = partReferences.children.length >= 2;
}

partAddReference.addEventListener("click", partAddReferencePair);

function partRenderTabs() {
  partTabs.replaceChildren();
  partChapterResults.forEach((chapter, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `tab${index === partActiveChapter ? " active" : ""}`;
    button.textContent = chapter.title;
    button.setAttribute("role", "tab");
    button.setAttribute("aria-selected", String(index === partActiveChapter));
    button.addEventListener("click", async () => {
      if (partEditor) partChapterResults[partActiveChapter].html = await partEditor.getData();
      partActiveChapter = index;
      partRenderTabs();
      await partEditor.setData(partChapterResults[index].html || "");
    });
    partTabs.append(button);
  });
}

async function partShowResults(data) {
  partDocumentId = data.document_id;
  partChapterResults = data.chapters;
  partActiveChapter = partChapterResults.findIndex((chapter) => chapter.html);
  if (partActiveChapter < 0) partActiveChapter = 0;
  partResultSource.textContent = data.filename;
  partResultElapsed.textContent =
    typeof data.elapsed_seconds === "number" ? `Waktu proses: ${partFormatElapsed(data.elapsed_seconds)}` : "";
  partRenderTabs();
  if (partEditor) {
    await partEditor.setData(partChapterResults[partActiveChapter]?.html || "");
  } else {
    partEditor = await ClassicEditor.create(document.querySelector("#part-editor"), {
      toolbar: ["heading", "|", "bold", "italic", "bulletedList", "numberedList", "insertTable", "undo", "redo"],
    });
    await partEditor.setData(partChapterResults[partActiveChapter]?.html || "");
  }
  partResults.hidden = false;
}

function partRenderSteps(stepDefs, stepsState) {
  partStepList.replaceChildren();
  let fractionSum = 0;
  stepDefs.forEach(({ key, label }) => {
    const info = stepsState[key] || { status: "pending", detail: "", current: 0, total: 0 };
    const li = document.createElement("li");
    li.className = `step-row ${info.status}`;
    const icon = document.createElement("span");
    icon.className = "step-icon";
    const text = document.createElement("span");
    text.className = "step-label";
    text.textContent = label;
    if (info.total > 0) {
      const count = document.createElement("span");
      count.className = "step-count";
      count.textContent = `${info.current}/${info.total}`;
      text.append(count);
    }
    if (info.detail) {
      const detail = document.createElement("span");
      detail.className = "step-detail";
      detail.textContent = info.detail;
      text.append(detail);
    }
    li.append(icon, text);
    partStepList.append(li);

    if (info.status === "done" || info.status === "skipped") {
      fractionSum += 1;
    } else if (info.total > 0) {
      fractionSum += info.current / info.total;
    }
  });
  const percent = Math.round((fractionSum / stepDefs.length) * 100);
  partProgressFill.style.width = `${percent}%`;
  partProgressPercent.textContent = `${percent}%`;
}

async function partPollStatus(sessionId, stepDefs) {
  while (true) {
    const response = await fetch(`/api/generate-status/${sessionId}`);
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Gagal memeriksa status proses.");
    partRenderSteps(stepDefs, data.steps);
    if (typeof data.elapsed_seconds === "number") {
      partElapsedTime.textContent = partFormatElapsed(data.elapsed_seconds);
    }
    if (data.status === "done") {
      data.result.elapsed_seconds = data.elapsed_seconds;
      return data.result;
    }
    if (data.status === "error") throw new Error(data.error || "Pembuatan dokumen gagal.");
    await new Promise((resolve) => setTimeout(resolve, 1200));
  }
}

partForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  partGenerateButton.disabled = true;
  partResults.hidden = true;
  partSetStatus("Memproses presentasi dan membuat isi per bab. Proses ini dapat memerlukan beberapa menit...");
  try {
    const response = await fetch("/api/generate-start", { method: "POST", body: new FormData(partForm) });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Gagal memulai pembuatan dokumen.");
    partRenderSteps(data.steps, data.initial_steps || {});
    partProgressPanel.hidden = false;
    const result = await partPollStatus(data.session_id, data.steps);
    await partShowResults(result);
    partProgressPanel.hidden = true;
    partSetStatus('Dokumen selesai dibuat. Seleksi bagian yang ingin dikoreksi, lalu klik "Koreksi seleksi".');
  } catch (error) {
    partSetStatus(error.message, true);
  } finally {
    partGenerateButton.disabled = false;
  }
});

correctSelectionButton.addEventListener("click", async () => {
  if (!partEditor || !partDocumentId) {
    partSetStatus("Buat dokumen terlebih dahulu sebelum mengoreksi.", true);
    return;
  }
  const selection = partEditor.model.document.selection;
  if (selection.isCollapsed) {
    partSetStatus("Seleksi (highlight) bagian yang ingin dikoreksi di editor terlebih dahulu.", true);
    return;
  }
  const selectedFragment = partEditor.model.getSelectedContent(selection);
  const selectedHtml = partEditor.data.stringify(selectedFragment);
  if (!selectedHtml.trim()) {
    partSetStatus("Seleksi tidak berisi konten yang bisa dikoreksi.", true);
    return;
  }

  correctSelectionButton.disabled = true;
  partSetStatus("Mengoreksi bagian yang diseleksi...");
  try {
    const chapterTitle = partChapterResults[partActiveChapter]?.title || "";
    const response = await fetch("/api/correct-selection", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        document_id: partDocumentId,
        chapter: chapterTitle,
        selected_html: selectedHtml,
      }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Gagal mengoreksi seleksi.");

    const viewFragment = partEditor.data.processor.toView(data.corrected_html);
    const modelFragment = partEditor.data.toModel(viewFragment);
    partEditor.model.insertContent(modelFragment);
    partChapterResults[partActiveChapter].html = await partEditor.getData();
    partSetStatus("Bagian yang diseleksi sudah dikoreksi.");
  } catch (error) {
    partSetStatus(error.message, true);
  } finally {
    correctSelectionButton.disabled = false;
  }
});

document.querySelector("#part-download-html").addEventListener("click", async () => {
  if (partEditor) partChapterResults[partActiveChapter].html = await partEditor.getData();
  const html = partChapterResults.map((chapter) => chapter.html).filter(Boolean).join("\n");
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "dokumen-hasil.html";
  link.click();
  URL.revokeObjectURL(link.href);
});
