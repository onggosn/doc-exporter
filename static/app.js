const form = document.querySelector("#generate-form");
const references = document.querySelector("#references");
const addReference = document.querySelector("#add-reference");
const statusBox = document.querySelector("#status");
const generateButton = document.querySelector("#generate-button");
const results = document.querySelector("#results");
const tabs = document.querySelector("#chapter-tabs");
const resultSource = document.querySelector("#result-source");
const perSlidePanel = document.querySelector("#per-slide-panel");
const babList = document.querySelector("#bab-list");
const slideProgress = document.querySelector("#slide-progress");
const slideList = document.querySelector("#slide-list");
const finalizeButton = document.querySelector("#finalize-button");
const generateProgressPanel = document.querySelector("#generate-progress-panel");
const generateProgressFill = document.querySelector("#generate-progress-fill");
const generateProgressPercent = document.querySelector("#generate-progress-percent");
const generateElapsedTime = document.querySelector("#generate-elapsed-time");
const generateStepList = document.querySelector("#generate-step-list");
const resultElapsed = document.querySelector("#result-elapsed");
let editor;
let chapterResults = [];
let activeChapter = 0;

let mappingSessionId = null;
let mappingChapters = [];
let mappingSlides = [];
let classifiedByRef = {};

function addReferencePair() {
  if (references.children.length >= 2) return;
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
    addReference.disabled = references.children.length >= 2;
  });
  references.append(row);
  addReference.disabled = references.children.length >= 2;
}

function setStatus(message, error = false) {
  statusBox.textContent = message;
  statusBox.className = error ? "error" : "status";
}

function formatElapsed(seconds) {
  const total = Math.round(seconds);
  const minutes = Math.floor(total / 60);
  const secs = total % 60;
  return minutes > 0 ? `${minutes}m ${secs}s` : `${secs}s`;
}

function renderTabs() {
  tabs.replaceChildren();
  chapterResults.forEach((chapter, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `tab${index === activeChapter ? " active" : ""}`;
    button.textContent = chapter.title;
    button.setAttribute("role", "tab");
    button.setAttribute("aria-selected", String(index === activeChapter));
    button.addEventListener("click", async () => {
      if (editor) chapterResults[activeChapter].html = await editor.getData();
      activeChapter = index;
      renderTabs();
      await editor.setData(chapterResults[index].html || "");
    });
    tabs.append(button);
  });
}

addReference.addEventListener("click", addReferencePair);

async function showResults(data) {
  chapterResults = data.chapters;
  activeChapter = chapterResults.findIndex((chapter) => chapter.html);
  if (activeChapter < 0) activeChapter = 0;
  resultSource.textContent = data.filename;
  resultElapsed.textContent =
    typeof data.elapsed_seconds === "number" ? `Waktu proses: ${formatElapsed(data.elapsed_seconds)}` : "";
  renderTabs();
  if (editor) {
    await editor.setData(chapterResults[activeChapter]?.html || "");
  } else {
    editor = await ClassicEditor.create(document.querySelector("#editor"), {
      toolbar: ["heading", "|", "bold", "italic", "bulletedList", "numberedList", "insertTable", "undo", "redo"],
    });
    await editor.setData(chapterResults[activeChapter]?.html || "");
  }
  results.hidden = false;
}

function renderBabList() {
  babList.replaceChildren();
  mappingChapters.forEach((chapter) => {
    const count = Object.values(classifiedByRef).filter((c) => c === chapter).length;
    const li = document.createElement("li");
    li.textContent = `${chapter} — ${count} slide`;
    babList.append(li);
  });
}

function renderSlideList() {
  slideList.replaceChildren();
  mappingSlides.forEach((slide) => {
    const li = document.createElement("li");
    li.className = "slide-row";

    const info = document.createElement("div");
    info.className = "slide-info";
    const titleEl = document.createElement("span");
    titleEl.className = "slide-title";
    titleEl.textContent = slide.title || "(tanpa judul)";
    const refEl = document.createElement("span");
    refEl.className = "slide-ref";
    refEl.textContent = slide.ref;
    info.append(titleEl, refEl);
    li.append(info);

    const chapter = classifiedByRef[slide.ref];
    if (chapter) {
      const result = document.createElement("span");
      result.className = "slide-result";
      result.textContent = chapter;
      li.append(result);
    } else {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "button secondary small";
      button.textContent = "Proses";
      button.addEventListener("click", () => classifySlide(slide.ref, button));
      li.append(button);
    }
    slideList.append(li);
  });
}

async function classifySlide(ref, button) {
  button.disabled = true;
  button.textContent = "Memproses...";
  setStatus(`Mengklasifikasikan slide ${ref}...`);
  try {
    const response = await fetch("/api/classify-slide", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: mappingSessionId, ref }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Gagal mengklasifikasikan slide.");
    classifiedByRef[data.ref] = data.chapter;
    renderBabList();
    renderSlideList();
    const done = Object.keys(classifiedByRef).length;
    slideProgress.textContent =
      `${done} / ${mappingSlides.length} slide — terakhir: "${data.title || data.ref}" -> ${data.chapter}`;
    if (data.done) {
      finalizeButton.hidden = false;
      setStatus('Semua slide sudah diklasifikasikan. Klik "Buat dokumen" untuk menyelesaikan.');
    } else {
      setStatus('Klik tombol "Proses" pada slide lain untuk lanjut.');
    }
  } catch (error) {
    setStatus(error.message, true);
    button.disabled = false;
    button.textContent = "Proses";
  }
}

function getMappingMode() {
  return form.querySelector('input[name="mapping_mode"]:checked').value;
}

function renderGenerateSteps(stepDefs, stepsState) {
  generateStepList.replaceChildren();
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
    generateStepList.append(li);

    if (info.status === "done" || info.status === "skipped") {
      fractionSum += 1;
    } else if (info.total > 0) {
      fractionSum += info.current / info.total;
    }
  });
  const percent = Math.round((fractionSum / stepDefs.length) * 100);
  generateProgressFill.style.width = `${percent}%`;
  generateProgressPercent.textContent = `${percent}%`;
}

async function pollGenerateStatus(sessionId, stepDefs) {
  while (true) {
    const response = await fetch(`/api/generate-status/${sessionId}`);
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Gagal memeriksa status proses.");
    renderGenerateSteps(stepDefs, data.steps);
    if (typeof data.elapsed_seconds === "number") {
      generateElapsedTime.textContent = formatElapsed(data.elapsed_seconds);
    }
    if (data.status === "done") {
      data.result.elapsed_seconds = data.elapsed_seconds;
      return data.result;
    }
    if (data.status === "error") throw new Error(data.error || "Pembuatan dokumen gagal.");
    await new Promise((resolve) => setTimeout(resolve, 1200));
  }
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  generateButton.disabled = true;
  results.hidden = true;
  perSlidePanel.hidden = true;
  generateProgressPanel.hidden = true;

  if (getMappingMode() === "per_slide") {
    setStatus("Menyiapkan slide untuk klasifikasi manual...");
    try {
      const response = await fetch("/api/prepare-mapping", { method: "POST", body: new FormData(form) });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Gagal menyiapkan slide.");
      mappingSessionId = data.session_id;
      mappingChapters = data.chapters;
      mappingSlides = data.slides;
      classifiedByRef = {};
      renderBabList();
      renderSlideList();
      slideProgress.textContent = `0 / ${mappingSlides.length} slide diklasifikasikan.`;
      finalizeButton.hidden = true;
      perSlidePanel.hidden = false;
      setStatus('Klik tombol "Proses" pada tiap slide untuk mengklasifikasikannya ke bab yang sesuai.');
    } catch (error) {
      setStatus(error.message, true);
    } finally {
      generateButton.disabled = false;
    }
    return;
  }

  setStatus("Memproses presentasi dan membuat isi per bab. Proses ini dapat memerlukan beberapa menit...");
  try {
    const response = await fetch("/api/generate-start", { method: "POST", body: new FormData(form) });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Gagal memulai pembuatan dokumen.");
    renderGenerateSteps(data.steps, data.initial_steps || {});
    generateProgressPanel.hidden = false;
    const result = await pollGenerateStatus(data.session_id, data.steps);
    await showResults(result);
    generateProgressPanel.hidden = true;
    setStatus("Dokumen selesai dibuat. Tinjau dan edit hasil sebelum digunakan.");
  } catch (error) {
    setStatus(error.message, true);
  } finally {
    generateButton.disabled = false;
  }
});

finalizeButton.addEventListener("click", async () => {
  finalizeButton.disabled = true;
  setStatus("Menyusun dokumen akhir...");
  try {
    const response = await fetch("/api/finalize", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: mappingSessionId }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Gagal membuat dokumen.");
    await showResults(data);
    perSlidePanel.hidden = true;
    setStatus("Dokumen selesai dibuat. Tinjau dan edit hasil sebelum digunakan.");
  } catch (error) {
    setStatus(error.message, true);
  } finally {
    finalizeButton.disabled = false;
  }
});

document.querySelector("#download-html").addEventListener("click", async () => {
  if (editor) chapterResults[activeChapter].html = await editor.getData();
  const html = chapterResults.map((chapter) => chapter.html).filter(Boolean).join("\n");
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "dokumen-hasil.html";
  link.click();
  URL.revokeObjectURL(link.href);
});
