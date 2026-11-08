(() => {
  const navButton = document.querySelector("#nav-docx-convert");
  const select = document.querySelector("#docx-profile-select");
  const summary = document.querySelector("#docx-profile-summary");
  const body = document.querySelector("#docx-profile-body");
  const statusBox = document.querySelector("#docx-status");
  const layoutRows = document.querySelector("#docx-layout-table tbody");
  const patternRows = document.querySelector("#docx-pattern-table tbody");
  const pageRows = document.querySelector("#docx-page-table tbody");
  const pageTableWrap = document.querySelector("#docx-page-table-wrap");
  const pageExcluded = document.querySelector("#docx-page-excluded");
  const pageHelp = document.querySelector("#docx-page-help");
  const evidenceRows = document.querySelector("#docx-evidence-table tbody");
  const evidenceEmpty = document.querySelector("#docx-evidence-empty");
  const rulesBox = document.querySelector("#docx-rules");
  const stateText = document.querySelector("#docx-settings-state");
  const addButton = document.querySelector("#docx-rule-add");
  const saveButton = document.querySelector("#docx-settings-save");
  const resetButton = document.querySelector("#docx-settings-reset");
  const captionSelect = document.querySelector("#docx-caption-default");

  const styleStats = document.querySelector("#docx-style-stats tbody");
  const styleStatsEmpty = document.querySelector("#docx-style-stats-empty");
  const styleStale = document.querySelector("#docx-style-stale");
  const styleEmpty = document.querySelector("#docx-style-empty");
  const styleGuide = document.querySelector("#docx-style-guide");
  const styleGuideMeta = document.querySelector("#docx-style-guide-meta");
  const styleGuideRows = document.querySelector("#docx-style-guide-table tbody");
  const styleProhibitions = document.querySelector("#docx-style-prohibitions");
  const styleExamples = document.querySelector("#docx-style-examples");
  const styleModel = document.querySelector("#docx-style-model");
  const styleRegion = document.querySelector("#docx-style-region");
  const styleWarning = document.querySelector("#docx-style-warning");
  const styleButton = document.querySelector("#docx-style-button");

  const GUIDE_LABELS = {
    nada: "Nada",
    sudut_pandang: "Sudut pandang",
    struktur_paragraf: "Struktur paragraf",
    panjang_kalimat: "Panjang kalimat",
    gaya_poin: "Gaya poin",
    istilah_dan_singkatan: "Istilah dan singkatan",
    format_angka: "Format angka",
    istilah_asing: "Istilah asing",
  };
  const SENTENCE_LABELS = { paragraf: "paragraf", poin: "poin" };
  const STATS_ABBREVIATIONS = 5;
  const STATS_ITALIC_EXAMPLES = 6;

  const createForm = document.querySelector("#docx-create-form");
  const createButton = document.querySelector("#docx-create-button");
  const progressPanel = document.querySelector("#docx-progress-panel");
  const progressFill = document.querySelector("#docx-progress-fill");
  const progressPercent = document.querySelector("#docx-percent");
  const elapsedText = document.querySelector("#docx-elapsed");
  const stepList = document.querySelector("#docx-step-list");
  const resultPanel = document.querySelector("#docx-result-panel");
  const resultSummary = document.querySelector("#docx-result-summary");
  const resultReport = document.querySelector("#docx-report");
  const downloadLink = document.querySelector("#docx-download");

  const POLL_MS = 2000;

  const RESET_LABEL = resetButton.textContent;
  const RESET_CONFIRM_MS = 4000;
  const LANJUT_LABELS = {
    per_bagian: "Mulai ulang di tiap bagian",
    berlanjut: "Berlanjut antarbagian",
  };

  let profileName = "";
  let rules = [];
  let captionDefault = "headline";
  let saved = false;
  let dirty = false;
  let resetTimer = null;
  let listLoaded = false;

  function setStatus(message, error = false) {
    statusBox.textContent = message;
    statusBox.className = error ? "error" : "status";
  }

  async function api(url, options) {
    const response = await fetch(url, options);
    let data = {};
    try {
      data = await response.json();
    } catch {
      // Respons bukan JSON; pesan umum di bawah dipakai.
    }
    if (!response.ok) throw new Error(data.error || `Permintaan gagal (${response.status}).`);
    return data;
  }

  function cell(content, className) {
    const td = document.createElement("td");
    if (content instanceof Node) td.append(content);
    else td.textContent = content;
    if (className) td.className = className;
    return td;
  }

  function code(text) {
    const el = document.createElement("code");
    el.textContent = text;
    return el;
  }

  function fillTable(tbody, rows) {
    tbody.replaceChildren(
      ...rows.map((cells) => {
        const tr = document.createElement("tr");
        tr.append(...cells);
        return tr;
      }),
    );
  }

  function joinCounts(counts) {
    return Object.entries(counts)
      .map(([key, value]) => `${key} ${value}`)
      .join(", ");
  }

  function renderState() {
    if (dirty) stateText.textContent = "Ada perubahan yang belum disimpan.";
    else if (saved) stateText.textContent = "Pengaturan tersimpan.";
    else stateText.textContent = "Pengaturan bawaan dari referensi (belum disimpan).";
  }

  function markDirty() {
    dirty = true;
    renderState();
  }

  function ruleInput(label, type, value, onInput) {
    const input = document.createElement("input");
    input.type = type;
    input.setAttribute("aria-label", label);
    input.placeholder = label;
    if (type === "checkbox") input.checked = value;
    else input.value = value;
    input.addEventListener("input", () => {
      onInput(type === "checkbox" ? input.checked : input.value);
      markDirty();
    });
    return input;
  }

  function renderRules() {
    rulesBox.replaceChildren(
      ...rules.map((rule, index) => {
        const row = document.createElement("div");
        row.className = "rule-row";

        const keyword = ruleInput("Kata kunci bagian", "text", rule.bagian_cocok, (v) => (rule.bagian_cocok = v));
        keyword.maxLength = 200;
        const prefix = ruleInput("Awalan", "text", rule.awalan, (v) => (rule.awalan = v));
        prefix.maxLength = 20;
        const start = ruleInput("Nomor awal", "number", rule.mulai, (v) => (rule.mulai = v === "" ? "" : Number(v)));
        start.min = 0;
        start.max = 9999;

        const mode = document.createElement("select");
        mode.setAttribute("aria-label", "Mode penomoran");
        Object.entries(LANJUT_LABELS).forEach(([value, label]) => {
          const option = new Option(label, value, false, value === rule.lanjut);
          mode.append(option);
        });
        mode.addEventListener("change", () => {
          rule.lanjut = mode.value;
          markDirty();
        });

        const activeLabel = document.createElement("label");
        activeLabel.className = "radio";
        activeLabel.append(ruleInput("Aktif", "checkbox", rule.aktif, (v) => (rule.aktif = v)), "Aktif");

        const remove = document.createElement("button");
        remove.type = "button";
        remove.className = "remove-reference";
        remove.textContent = "Hapus";
        remove.addEventListener("click", () => {
          rules.splice(index, 1);
          markDirty();
          renderRules();
        });

        row.append(keyword, prefix, start, mode, activeLabel, remove);
        return row;
      }),
    );
    if (rules.length === 0) {
      const empty = document.createElement("p");
      empty.className = "help";
      empty.textContent = "Belum ada aturan: caption tidak diberi nomor otomatis.";
      rulesBox.append(empty);
    }
    renderState();
  }

  function renderPage(page) {
    const hasPdf = Boolean(page && page.pdf.length);
    pageTableWrap.hidden = !hasPdf;
    pageExcluded.replaceChildren();
    if (!hasPdf) {
      pageHelp.textContent = "Profil ini dibuat tanpa PDF referensi. Bangun ulang dengan opsi --pdf untuk mengukur properti halaman.";
      return;
    }
    pageHelp.textContent = page.catatan;

    fillTable(
      pageRows,
      page.pdf.map((pdf) => {
        const margin = pdf.margin_pt;
        const number = pdf.nomor_halaman;
        const mark = pdf.watermark;
        const scale = page.skala_cetak.find((item) => item.pdf === pdf.file);
        return [
          cell(pdf.file),
          cell(`${pdf.ukuran.nama || "ukuran khusus"} (${pdf.ukuran.lebar_pt}×${pdf.ukuran.tinggi_pt})`),
          cell(`kiri ${margin.kiri}, kanan ${margin.kanan}, atas ${margin.atas}, bawah ${margin.bawah}`),
          cell(number
            ? `${number.perataan} bawah, ${number.font} ${number.ukuran_pt} pt (${number.halaman_bernomor}/${number.dari_halaman} halaman)`
            : "tidak ada"),
          cell(mark
            ? `"${mark.teks}" ${mark.sudut_derajat}°, ${mark.warna}, ±${mark.ukuran_pt_perkiraan.join("–")} pt`
            : "tidak ada"),
          cell(scale ? `${scale.skala} (HTML ${scale.ukuran_html_pt} pt, PDF ${scale.ukuran_pdf_pt} pt)` : "-"),
        ];
      }),
    );

    const excluded = new Set();
    page.pdf.forEach((pdf) =>
      pdf.footer_berulang
        .filter((footer) => footer.dikecualikan)
        .forEach((footer) => excluded.add(`${footer.pola} — ${footer.alasan}`)),
    );
    pageExcluded.replaceChildren(
      ...[...excluded].map((text) => {
        const li = document.createElement("li");
        li.textContent = text;
        return li;
      }),
    );
    if (excluded.size === 0) {
      const li = document.createElement("li");
      li.textContent = "Tidak ada.";
      pageExcluded.append(li);
    }
  }

  function listItems(list, texts) {
    list.replaceChildren(
      ...texts.map((text) => {
        const li = document.createElement("li");
        li.textContent = text;
        return li;
      }),
    );
  }

  function statsRows(stats) {
    const sentences = Object.entries(stats.kalimat)
      .filter(([, value]) => value.jumlah > 0)
      .map(([kind, value]) => `${SENTENCE_LABELS[kind] || kind}: rata-rata ${value.rata_kata} kata (median ${value.median_kata}, ${value.jumlah} kalimat)`);
    const numbers = stats.angka;
    const abbreviations = stats.singkatan.slice(0, STATS_ABBREVIATIONS).map((item) => {
      const short = Object.entries(item.singkatan)
        .filter(([, n]) => n > 0)
        .map(([word, n]) => `${word} ${n}`)
        .join("/");
      return `${short || "-"} lawan ${item.lengkap} ${item.jumlah_lengkap}`;
    });
    const technical = Object.entries(stats.istilah_teknis).map(([term, n]) => `${term} ${n}`);
    return [
      ["Panjang kalimat", sentences.join("; ") || "-"],
      ["Desimal", `koma ${numbers.desimal_koma}, titik ${numbers.desimal_titik}`],
      ["Persen", `rapat (4,5%) ${numbers.persen_rapat}, berspasi (4,5 %) ${numbers.persen_spasi}`],
      ["Rupiah", `rapat (Rp16.550) ${numbers.rupiah_rapat}, berspasi (Rp 16.550) ${numbers.rupiah_spasi}`],
      ["Singkatan informal", abbreviations.join("; ") || "-"],
      ["Istilah teknis", technical.join(", ") || "-"],
      ["Istilah asing bercetak miring", `${stats.miring.jumlah} (mis. ${stats.miring.teratas.slice(0, STATS_ITALIC_EXAMPLES).join(", ")})`],
    ];
  }

  function renderStyleGuide(saved, stale) {
    styleStale.hidden = !(saved && stale);
    styleEmpty.hidden = Boolean(saved);
    styleGuide.hidden = !saved;
    if (!saved) return;

    const guide = saved.panduan;
    styleGuideMeta.textContent =
      `Dibuat ${saved.dibuat.slice(0, 10)} dengan ${saved.model_id}.` +
      (saved.contoh_dibuang ? ` ${saved.contoh_dibuang} contoh dibuang karena bukan kutipan persis.` : "");
    fillTable(styleGuideRows, [
      ...Object.entries(GUIDE_LABELS).map(([key, label]) => [cell(label), cell(guide[key])]),
      [cell("Caption headline"), cell(guide.gaya_caption.headline)],
      [cell("Caption label"), cell(guide.gaya_caption.label)],
    ]);
    listItems(styleProhibitions, guide.larangan.length ? guide.larangan : ["Tidak ada."]);
    listItems(
      styleExamples,
      guide.contoh.length ? guide.contoh.map((item) => `[${item.konteks}] ${item.teks}`) : ["Tidak ada."],
    );
  }

  function renderStyle(style) {
    const stats = style.statistik;
    styleStatsEmpty.hidden = Boolean(stats);
    fillTable(styleStats, stats ? statsRows(stats).map(([label, value]) => [cell(label), cell(value)]) : []);
    styleButton.disabled = !style.sampel;
    styleWarning.textContent = style.sampel
      ? `${style.sampel.blok} blok teks referensi (${style.sampel.karakter} karakter) akan dikirim ke Amazon Bedrock. ` +
        "Pastikan dokumen referensi boleh dikirim ke layanan model."
      : "Profil ini belum memiliki sampel gaya, jadi panduan belum bisa dibuat.";
    renderStyleGuide(style.panduan, style.usang);
  }

  function renderSteps(stepDefs, state) {
    let fraction = 0;
    stepList.replaceChildren(
      ...stepDefs.map(({ key, label }) => {
        const info = state[key] || { status: "pending", detail: "", current: 0, total: 0 };
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
        if (info.status === "done" || info.status === "skipped") fraction += 1;
        else if (info.total > 0) fraction += info.current / info.total;
        return li;
      }),
    );
    const percent = Math.round((fraction / stepDefs.length) * 100);
    progressFill.style.width = `${percent}%`;
    progressPercent.textContent = `${percent}%`;
  }

  function renderResult(result) {
    const report = result.laporan;
    resultSummary.textContent = `${result.filename}: ${report.gambar} gambar dari ${report.slide} slide, narasi untuk ${report.narasi} slide.`;
    const lines = [`Setiap gambar slide terpakai ${report.gambar_ganda.length ? "dan ada yang ganda: " + report.gambar_ganda.length : "tepat satu kali"}.`];
    if (report.angka_tidak_didukung.length) {
      lines.push(
        "Angka di narasi yang tidak ada di slide asalnya (periksa): " +
          report.angka_tidak_didukung.map((item) => `${item.ref}: ${item.angka.join(", ")}`).join("; "),
      );
    } else {
      lines.push("Semua angka di narasi ada di slide asalnya.");
    }
    const matched = Object.entries(report.aturan_bab).filter(([, prefixes]) => prefixes.length);
    const unmatched = Object.entries(report.aturan_bab).filter(([, prefixes]) => !prefixes.length).map(([chapter]) => chapter);
    if (matched.length) lines.push("Penomoran caption: " + matched.map(([chapter, prefixes]) => `${chapter} → ${prefixes.join(", ")}`).join("; ") + ".");
    if (unmatched.length) {
      lines.push(`Tidak ada aturan penomoran yang cocok dengan judul bab: ${unmatched.join(", ")}. Caption mengikuti pengaturan bawaan; sesuaikan kata kunci aturan di panel Penomoran caption bila perlu.`);
    }
    if (report.bab_cadangan.length) {
      lines.push(
        "Narasi model gagal untuk bab ini, jadi isinya memakai teks slide apa adanya (tidak ditulis ulang, periksa gayanya): " +
          report.bab_cadangan.map((item) => `${item.bab} (${item.alasan})`).join("; "),
      );
    }
    if (report.slide_dilewati.length) lines.push(`Slide tanpa isi atau slide pembatas (sampul, agenda, penutup) yang dilewati: ${report.slide_dilewati.join(", ")}.`);
    if (report.slide_dikecualikan.length) {
      lines.push(`Slide yang dikecualikan dari semua proses (${report.slide_dikecualikan.length}): ${report.slide_dikecualikan.join(", ")}.`);
    }
    if (report.slide_teks_asli.length) {
      lines.push(`Slide yang teksnya dibawa apa adanya (bab kesimpulan/arahan, atau blok Hasil Pembahasan Dewan Gubernur dan Key Takeaways): ${report.slide_teks_asli.join(", ")}.`);
    }
    if (report.slide_dipecah.length) {
      lines.push(`Slide yang dipecah menjadi panel, tiap panel dengan caption dari judulnya (${report.slide_dipecah.length}): ${report.slide_dipecah.join(", ")}.`);
    }
    if (report.slide_teks.length) {
      lines.push(`Slide padat teks yang ditulis sebagai teks, tanpa gambar (${report.slide_teks.length}): ${report.slide_teks.join(", ")}. Aturannya di doc_composer.py (TEXT_DOMINANT_CHARS).`);
    }
    if (report.bab_kosong.length) lines.push(`Bab tanpa slide (tidak dibuat): ${report.bab_kosong.join(", ")}.`);
    listItems(resultReport, lines);
    downloadLink.href = result.download_url;
    downloadLink.setAttribute("download", result.filename);
    resultPanel.hidden = false;
  }

  async function pollDocx(sessionId, stepDefs) {
    for (;;) {
      const data = await api(`/api/docx/status/${encodeURIComponent(sessionId)}`);
      renderSteps(stepDefs, data.steps);
      elapsedText.textContent = `${Math.round(data.elapsed_seconds)}s`;
      if (data.status === "done") return data.result;
      if (data.status === "error") throw new Error(data.error);
      await new Promise((resolve) => setTimeout(resolve, POLL_MS));
    }
  }

  function renderProfile(profile) {
    const created = profile.meta.created_at ? profile.meta.created_at.slice(0, 10) : "-";
    summary.textContent =
      `${profile.layouts.length} layout dari ${profile.meta.sources.length} dokumen referensi ` +
      `(dibuat ${created}). Caption: ${joinCounts(profile.caption_counts)}.`;

    fillTable(
      layoutRows,
      profile.layouts.map((layout) => {
        const name = cell(layout.name);
        name.title = layout.id;
        return [
          name,
          cell(String(layout.count)),
          cell(layout.width_bucket.length ? layout.width_bucket.join(" / ") : "-"),
          cell(joinCounts(layout.caption.counts)),
          cell(String(layout.caption.bernomor)),
        ];
      }),
    );
    fillTable(
      patternRows,
      profile.row_patterns.map((pattern) => [
        cell(code(pattern.tokens)),
        cell(pattern.role),
        cell(String(pattern.rows)),
        cell(String(pattern.tables)),
        cell(Object.entries(pattern.width_buckets).map(([width, n]) => `${width} ×${n}`).join(", ")),
      ]),
    );
    fillTable(
      evidenceRows,
      profile.evidence.map((item) => [
        cell(item.bagian),
        cell(item.awalan),
        cell(String(item.jumlah)),
        cell(`${item.pertama}–${item.terakhir}`),
        cell(item.contoh),
      ]),
    );
    evidenceEmpty.hidden = profile.evidence.length > 0;
    renderPage(profile.page);
    renderStyle(profile.gaya);

    rules = profile.settings.penomoran.map((rule) => ({ ...rule }));
    captionDefault = profile.settings.caption_default || "headline";
    captionSelect.value = captionDefault;
    saved = profile.settings_saved;
    dirty = false;
    renderRules();
    body.hidden = false;
  }

  async function loadProfile(name) {
    profileName = name;
    body.hidden = true;
    setStatus("Memuat profil…");
    try {
      renderProfile(await api(`/api/profiles/${encodeURIComponent(name)}`));
      setStatus("");
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  async function loadList() {
    setStatus("Memuat daftar profil…");
    try {
      const { profiles } = await api("/api/profiles");
      listLoaded = true;
      select.replaceChildren(...profiles.map((p) => new Option(p.title, p.name)));
      select.disabled = profiles.length === 0;
      createButton.disabled = profiles.length === 0;
      if (profiles.length === 0) {
        summary.textContent = "Belum ada profil. Buat dengan: python profile_builder.py <berkas.html> --name <nama>.";
        setStatus("");
        return;
      }
      await loadProfile(profiles[0].name);
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  function cancelResetConfirm() {
    clearTimeout(resetTimer);
    resetButton.textContent = RESET_LABEL;
    resetButton.dataset.confirm = "";
  }

  select.addEventListener("change", () => {
    if (dirty && !window.confirm("Perubahan yang belum disimpan akan hilang. Lanjutkan?")) {
      select.value = profileName;
      return;
    }
    cancelResetConfirm();
    loadProfile(select.value);
  });

  createForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    createButton.disabled = true;
    resultPanel.hidden = true;
    const formData = new FormData(createForm);
    formData.append("profile", profileName);
    setStatus("");
    try {
      const started = await api("/api/docx/start", { method: "POST", body: formData });
      progressPanel.hidden = false;
      renderSteps(started.steps, started.initial_steps);
      renderResult(await pollDocx(started.session_id, started.steps));
    } catch (error) {
      setStatus(error.message, true);
    } finally {
      createButton.disabled = false;
    }
  });

  captionSelect.addEventListener("change", () => {
    captionDefault = captionSelect.value;
    markDirty();
  });

  addButton.addEventListener("click", () => {
    rules.push({ bagian_cocok: "", awalan: "", mulai: 1, lanjut: "per_bagian", aktif: true });
    markDirty();
    renderRules();
  });

  saveButton.addEventListener("click", async () => {
    cancelResetConfirm();
    saveButton.disabled = true;
    try {
      const result = await api(`/api/profiles/${encodeURIComponent(profileName)}/settings`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ penomoran: rules, caption_default: captionDefault }),
      });
      rules = result.settings.penomoran.map((rule) => ({ ...rule }));
      captionDefault = result.settings.caption_default;
      captionSelect.value = captionDefault;
      saved = true;
      dirty = false;
      renderRules();
      setStatus("Pengaturan penomoran disimpan.");
    } catch (error) {
      setStatus(error.message, true);
    } finally {
      saveButton.disabled = false;
    }
  });

  // Dua klik: klik pertama meminta konfirmasi karena menghapus pengaturan tersimpan.
  resetButton.addEventListener("click", async () => {
    if (resetButton.dataset.confirm !== "1") {
      resetButton.dataset.confirm = "1";
      resetButton.textContent = "Klik lagi untuk mengembalikan";
      resetTimer = setTimeout(cancelResetConfirm, RESET_CONFIRM_MS);
      return;
    }
    cancelResetConfirm();
    try {
      const result = await api(`/api/profiles/${encodeURIComponent(profileName)}/settings`, { method: "DELETE" });
      rules = result.settings.penomoran.map((rule) => ({ ...rule }));
      captionDefault = result.settings.caption_default;
      captionSelect.value = captionDefault;
      saved = false;
      dirty = false;
      renderRules();
      setStatus("Pengaturan dikembalikan ke hasil analisis referensi.");
    } catch (error) {
      setStatus(error.message, true);
    }
  });

  styleButton.addEventListener("click", async () => {
    styleButton.disabled = true;
    setStatus("Membuat panduan gaya lewat Bedrock, ini bisa memakan waktu hingga satu menit…");
    try {
      const result = await api(`/api/profiles/${encodeURIComponent(profileName)}/style-guide`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model_id: styleModel.value, region: styleRegion.value }),
      });
      renderStyleGuide(result.panduan, result.usang);
      setStatus("Panduan gaya dibuat.");
    } catch (error) {
      setStatus(error.message, true);
    } finally {
      styleButton.disabled = false;
    }
  });

  navButton.addEventListener("click", () => {
    if (!listLoaded) loadList();
  });
})();
