from __future__ import annotations

import json
import logging
import os
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any

import boto3
from botocore.config import Config
from botocore.exceptions import BotoCoreError, ClientError
from typing import Callable

import doc_composer
import style_guide

logger = logging.getLogger(__name__)

ProgressCallback = Callable[[str, str, str, int, int], None]


def _report(
    callback: ProgressCallback | None,
    step: str,
    status: str,
    detail: str = "",
    current: int = 0,
    total: int = 0,
) -> None:
    if callback:
        callback(step, status, detail, current, total)

BEDROCK_CLIENT_CONFIG = Config(
    read_timeout=300,
    connect_timeout=10,
    retries={"max_attempts": 5, "mode": "adaptive"},  # mode adaptive membatasi laju sendiri saat terkena throttling
    max_pool_connections=16,  # cukup untuk worker paralel
)

_USAGE_LOCK = threading.Lock()
MAX_WORKERS = 8


def run_parallel(items, worker, workers=1, stop_on_error=True, on_done=None):
    """Jalankan worker(item) untuk tiap item dengan sebanyak `workers` thread. Mengembalikan daftar hasil
    berurutan sesuai items: (True, nilai) atau (False, exception). `on_done(indeks, hasil)` dipanggil di thread
    pemanggil, tiap kali satu item selesai. Dengan stop_on_error, item yang belum mulai dibatalkan begitu ada
    yang gagal dan exception pertama (urut item) dilempar; tanpa itu semua item dijalankan dan kegagalan
    dikembalikan sebagai hasil."""
    outcomes: list = [None] * len(items)

    def call(item):
        try:
            return True, worker(item)
        except Exception as exc:  # noqa: BLE001 - dikembalikan ke pemanggil
            return False, exc

    def record(index: int, outcome) -> bool:
        outcomes[index] = outcome
        if on_done:
            on_done(index, outcome)
        return outcome[0] or not stop_on_error

    count = max(1, min(int(workers or 1), MAX_WORKERS, len(items) or 1))
    if count == 1:
        for index, item in enumerate(items):
            if not record(index, call(item)):
                break
    else:
        pool = ThreadPoolExecutor(max_workers=count)
        try:
            futures = {pool.submit(call, item): index for index, item in enumerate(items)}
            for future in as_completed(futures):
                if not record(futures[future], future.result()):
                    break
        finally:
            pool.shutdown(wait=True, cancel_futures=True)
    if stop_on_error:
        for outcome in outcomes:
            if outcome is not None and not outcome[0]:
                raise outcome[1]
    return outcomes


DEFAULT_MODEL_ID = "global.anthropic.claude-opus-4-6-v1"
MAX_IMAGE_BYTES = 10 * 1024 * 1024

class BedrockError(RuntimeError):
    """Bedrock request or response error safe to show in the UI."""


class NarrationPartial(BedrockError):
    """Sebagian slide sebuah bab gagal dinarasikan. `partial` dan `flagged` memuat slide yang berhasil;
    `failed_slides` slide yang gagal."""

    def __init__(self, message: str, partial: dict, flagged: dict, failed_slides: list):
        super().__init__(message)
        self.partial, self.flagged, self.failed_slides = partial, flagged, failed_slides


class BedrockDocumentGenerator:
    workers = 1  # jumlah thread untuk langkah per-slide; diatur pemanggil (alur PPT → DOCX)

    def __init__(
        self,
        model_id: str | None = None,
        region: str | None = None,
    ):
        self.model_id = model_id or os.environ.get("BEDROCK_MODEL_ID") or DEFAULT_MODEL_ID
        if not self.model_id:
            raise BedrockError(
                "Tentukan Model ID Bedrock di formulir atau melalui BEDROCK_MODEL_ID."
            )
        self.region = (
            region
            or os.environ.get("AWS_REGION")
            or os.environ.get("AWS_DEFAULT_REGION")
        )
        try:
            session = boto3.Session(region_name=self.region)
            self.region = self.region or session.region_name
            if not self.region:
                raise BedrockError(
                    "Region AWS belum ditentukan. Isi Region AWS di formulir atau atur "
                    "AWS_REGION/AWS_DEFAULT_REGION maupun region pada AWS profile."
                )
            self.client = session.client(
                "bedrock-runtime", region_name=self.region, config=BEDROCK_CLIENT_CONFIG
            )
        except (BotoCoreError, ClientError) as exc:
            raise BedrockError(f"Tidak dapat menyiapkan koneksi AWS Bedrock: {exc}") from exc
        self.usage_totals: dict[str, dict[str, int]] = {}

    def _track_usage(self, step: str | None, usage: dict[str, Any] | None) -> None:
        if not step or not usage:
            return
        with _USAGE_LOCK:  # dipanggil dari beberapa thread saat worker paralel
            totals = self.usage_totals.setdefault(step, {"input_tokens": 0, "output_tokens": 0})
            totals["input_tokens"] += usage.get("inputTokens", 0) or 0
            totals["output_tokens"] += usage.get("outputTokens", 0) or 0

    def _converse(
        self, content: list[dict[str, Any]], max_tokens: int = 6000, step: str | None = None
    ) -> str:
        text_parts = [block["text"] for block in content if "text" in block]
        has_image = any("image" in block for block in content)
        logger.debug(
            "Converse REQUEST max_tokens=%s has_image=%s prompt=%s",
            max_tokens, has_image, "\n---\n".join(text_parts),
        )
        try:
            response = self.client.converse(
                modelId=self.model_id,
                messages=[{"role": "user", "content": content}],
                inferenceConfig={"maxTokens": max_tokens, "temperature": 0.2},
            )
            blocks = response["output"]["message"]["content"]
            result = "\n".join(block["text"] for block in blocks if "text" in block).strip()
            logger.debug(
                "Converse RESPONSE usage=%s response=%s",
                response.get("usage"), result,
            )
            self._track_usage(step, response.get("usage"))
            return result
        except (BotoCoreError, ClientError, KeyError, TypeError) as exc:
            message = str(exc)
            if isinstance(exc, ClientError):
                message = exc.response.get("Error", {}).get("Message", message)
            logger.error("Converse ERROR: %s", message)
            raise BedrockError(
                f"Bedrock gagal memproses permintaan. Model: {self.model_id}. "
                f"Region: {self.region}. Detail: {message}"
            ) from exc

    describe_split_slides = True  # alur PPT → DOCX mematikannya: slide yang dipecah cukup diringkas per panel

    def describe_visuals(
        self, deck: dict[str, Any], progress_callback: ProgressCallback | None = None
    ) -> None:
        """Deskripsikan slide bergambar; hasilnya menjadi fakta visual untuk pemetaan dan narasi."""
        visual_slides = [
            slide for slide in deck["items"]
            if slide.get("image_path") and (self.describe_split_slides or len(slide.get("panel_images") or []) < 2)
        ]
        if not visual_slides:
            _report(progress_callback, "visuals", "skipped", "Tidak ada slide bervisual", 0, 0)
            return
        total = len(visual_slides)
        logger.info("Describe visuals dimulai: %s slide bervisual dari %s total slide", total, len(deck["items"]))
        _report(progress_callback, "visuals", "running", "", 0, total)

        def describe(slide: dict[str, Any]) -> None:
            image = Path(slide["image_path"]).read_bytes()
            if len(image) > MAX_IMAGE_BYTES:
                logger.warning("Slide %s dilewati: gambar %s bytes > batas %s", slide["ref"], len(image), MAX_IMAGE_BYTES)
                slide["visual_desc"] = (
                    "Gambar slide tidak dikirim untuk analisis visual karena ukurannya melebihi 5 MB."
                )
                return
            logger.debug("Describe visual slide ref=%s", slide["ref"])
            slide["visual_desc"] = self._converse(
                [
                    {
                        "text": (
                            "Jelaskan secara singkat informasi penting yang terlihat pada slide ini "
                            "(grafik, diagram, gambar, atau tata letak). Jangan menebak detail yang tidak terbaca. "
                            f"Teks yang sudah diekstrak: {slide.get('title', '')} "
                            f"{' '.join(slide.get('text', []))}"
                        )
                    },
                    {"image": {"format": "png", "source": {"bytes": image}}},
                ],
                max_tokens=500,
                step="visuals",
            )

        finished = 0

        def advance(_index: int, _outcome: Any) -> None:
            nonlocal finished
            finished += 1
            _report(progress_callback, "visuals", "running", "", finished, total)

        run_parallel(visual_slides, describe, self.workers, on_done=advance)
        logger.info("Describe visuals selesai: %s slide diproses", total)
        _report(progress_callback, "visuals", "done", "", total, total)

    def summarize_images(
        self, deck: dict[str, Any], progress_callback: ProgressCallback | None = None,
        skip_refs: set[str] | frozenset[str] = frozenset(),
    ) -> None:
        """Ringkas tiap gambar yang dipasang di dokumen untuk caption: tiap panel pada slide yang dipecah,
        atau slide utuh. Slide di `skip_refs` (mis. bab kesimpulan/arahan) tidak diringkas."""
        tasks: list[tuple[dict[str, Any], dict[str, Any] | None]] = []
        for slide in deck["items"]:
            if not slide.get("image_path") or slide["ref"] in skip_refs:
                continue
            panels = slide.get("panel_images") or []
            tasks += [(slide, panel) for panel in panels] if len(panels) >= 2 else [(slide, None)]
        if not tasks:
            _report(progress_callback, "captions", "skipped", "Tidak ada gambar untuk diringkas", 0, 0)
            return
        total = len(tasks)
        _report(progress_callback, "captions", "running", "", 0, total)

        def run(task: tuple[dict[str, Any], dict[str, Any] | None]) -> None:
            slide, panel = task
            target = panel if panel is not None else slide
            image = Path(target["image_path"]).read_bytes()
            if len(image) > MAX_IMAGE_BYTES:
                logger.warning("Ringkasan %s dilewati: gambar %s bytes > batas", slide["ref"], len(image))
                return
            target["summary"], target["summary_flagged"] = self.summarize_image(image, slide, (panel or {}).get("header"))

        finished = 0

        def advance(_index: int, _outcome: Any) -> None:
            nonlocal finished
            finished += 1
            _report(progress_callback, "captions", "running", "", finished, total)

        run_parallel(tasks, run, self.workers, on_done=advance)
        _report(progress_callback, "captions", "done", "", total, total)

    def summarize_image(
        self, image: bytes, slide: dict[str, Any], header: str | None
    ) -> tuple[str, list[str]]:
        """Ringkasan 1-2 kalimat dari gambar itu sendiri untuk caption. Angka yang tidak ada di teks slide
        diminta dibuang sekali; yang tersisa dikembalikan sebagai peringatan."""
        facts = doc_composer.all_facts({**slide, "visual_desc": None})
        feedback = ""
        summary, numbers = "", []
        for attempt in range(2):
            prompt = doc_composer.image_summary_prompt(slide.get("title"), header, facts, feedback)
            summary = doc_composer.clean_summary(self._converse(
                [{"text": prompt}, {"image": {"format": "png", "source": {"bytes": image}}}],
                max_tokens=300, step="visuals",
            ))
            numbers = doc_composer.unsupported_numbers([summary], facts)
            if not numbers:
                break
            feedback = (
                f"\n\nJawaban sebelumnya menyebut angka yang tidak ada di teks slide: {', '.join(numbers)}. "
                "Tulis ulang tanpa angka itu; sebut hanya angka yang tertulis di teks slide."
            )
        return summary, numbers

    @staticmethod
    def _slide_text(slide: dict[str, Any]) -> str:
        parts = [
            f"Referensi: {slide['ref']}",
            f"Judul: {slide.get('title') or '(tanpa judul)'}",
            "Teks: " + " | ".join(slide.get("text", [])),
        ]
        if slide.get("tables"):
            parts.append("Tabel: " + json.dumps(slide["tables"], ensure_ascii=False))
        if slide.get("charts"):
            parts.append("Data chart: " + json.dumps(slide["charts"], ensure_ascii=False))
        if slide.get("notes"):
            parts.append("Speaker notes: " + slide["notes"])
        if slide.get("visual_desc"):
            parts.append("Deskripsi visual: " + slide["visual_desc"])
        return "\n".join(parts)

    @staticmethod
    def _last_json_object(text: str) -> dict[str, Any] | None:
        decoder = json.JSONDecoder()
        found = None
        position = text.find("{")
        while position != -1:
            try:
                value, end = decoder.raw_decode(text, position)
            except json.JSONDecodeError:
                position = text.find("{", position + 1)
                continue
            if isinstance(value, dict):
                found = value
            position = text.find("{", end)
        return found

    @staticmethod
    def _parse_json(text: str, context: str = "pemetaan bab") -> dict[str, Any]:
        cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip(), flags=re.IGNORECASE)
        try:
            try:
                data = json.loads(cleaned)
            except json.JSONDecodeError as exc:
                # Model kadang menjawab, menulis koreksi, lalu menjawab lagi: pakai objek lengkap terakhir.
                data = BedrockDocumentGenerator._last_json_object(cleaned)
                if data is None:
                    raise exc
        except json.JSONDecodeError as exc:
            logger.error(
                "JSON parse GAGAL (%s): %s. Panjang teks=%s. Teks mentah:\n%s",
                context, exc, len(text), text,
            )
            raise BedrockError(f"Model mengembalikan {context} yang bukan JSON valid.") from exc
        if not isinstance(data, dict):
            raise BedrockError(f"Format {context} dari model tidak valid.")
        return data

    MAP_SLIDES_MAX_ATTEMPTS = 3

    def _map_slides(
        self,
        deck: dict[str, Any],
        chapters: list[str],
        mode: str = "batch",
        progress_callback: ProgressCallback | None = None,
    ) -> dict[str, list[str]]:
        if mode == "per_slide":
            mapping = self._map_slides_per_slide(deck, chapters, progress_callback)
        else:
            mapping = self._map_slides_batch(deck, chapters, progress_callback)
        return mapping

    def classify_slide(self, slide: dict[str, Any], chapters: list[str]) -> str:
        """Pilih satu bab untuk satu slide. Dipanggil baik oleh mode per-slide otomatis
        maupun endpoint klasifikasi manual (satu langkah = satu pemanggilan model)."""
        chapters_list = "\n".join(f"- {chapter}" for chapter in chapters)
        prompt = (
            "Pilih SATU bab yang paling sesuai untuk slide berikut ini, berdasarkan topiknya "
            "(judul, teks, data tabel/chart, dan deskripsi visual bila ada). "
            "Balas HANYA dengan judul bab yang dipilih, disalin tepat sama persis dari daftar, "
            "tanpa tanda kutip, tanpa penjelasan tambahan.\n"
            f"Daftar bab:\n{chapters_list}\n\n"
            f"Slide:\n{self._slide_text(slide)}"
        )
        feedback = ""
        for attempt in range(1, self.MAP_SLIDES_MAX_ATTEMPTS + 1):
            response = self._converse([{"text": prompt + feedback}], max_tokens=100, step="mapping")
            candidate = response.strip().strip('"').strip("'").strip()
            if candidate in chapters:
                return candidate
            feedback = (
                f"\n\nJawaban sebelumnya ({candidate!r}) tidak sama persis dengan salah satu judul "
                "bab di daftar. Balas hanya dengan salah satu judul bab di atas, disalin tepat sama persis."
            )
        raise BedrockError(
            f"Model gagal memilih bab yang valid untuk slide {slide['ref']} "
            f"setelah {self.MAP_SLIDES_MAX_ATTEMPTS} percobaan."
        )

    def _map_slides_per_slide(
        self,
        deck: dict[str, Any],
        chapters: list[str],
        progress_callback: ProgressCallback | None = None,
    ) -> dict[str, list[str]]:
        """Klasifikasikan tiap slide satu per satu — pembagian 1-ke-1 dijamin oleh kode,
        bukan oleh model, sehingga tidak ada slide yang hilang atau dobel."""
        mapping: dict[str, list[str]] = {chapter: [] for chapter in chapters}
        total = len(deck["items"])
        _report(progress_callback, "mapping", "running", "", 0, total)
        finished = 0

        def advance(_index: int, _outcome: Any) -> None:
            nonlocal finished
            finished += 1
            _report(progress_callback, "mapping", "running", "", finished, total)

        outcomes = run_parallel(
            deck["items"], lambda slide: self.classify_slide(slide, chapters), self.workers, on_done=advance
        )
        for slide, (_ok, chapter) in zip(deck["items"], outcomes):  # urutan slide dalam bab tetap urutan deck
            mapping[chapter].append(slide["ref"])
        _report(progress_callback, "mapping", "done", "", total, total)
        return mapping

    def _map_slides_batch(
        self,
        deck: dict[str, Any],
        chapters: list[str],
        progress_callback: ProgressCallback | None = None,
    ) -> dict[str, list[str]]:
        slide_text = "\n\n".join(self._slide_text(slide) for slide in deck["items"])
        expected = {slide["ref"] for slide in deck["items"]}
        all_refs = [slide["ref"] for slide in deck["items"]]
        empty_mapping = {chapter: [] for chapter in chapters}
        base_prompt = (
            "Petakan setiap slide ke bab yang paling sesuai berdasarkan isi sebenarnya dari slide itu "
            "(judul, teks, data tabel/chart, dan deskripsi visual bila ada) — termasuk slide yang didominasi "
            "gambar/chart. Tentukan bab untuk setiap slide secara independen berdasarkan topiknya sendiri. "
            "JANGAN mengelompokkan slide yang topiknya tidak berkaitan ke bab yang sama hanya karena sulit "
            "diklasifikasikan atau untuk memenuhi syarat 'setiap ref harus muncul' — kalau sebuah slide sulit "
            "ditentukan, pilih bab dengan kecocokan topik terdekat, bukan bab yang kebetulan masih kosong atau "
            "bersifat umum. "
            "Setiap ref slide di bawah ini harus muncul di dalam hasil tepat satu kali saja, "
            "tidak boleh hilang maupun duplikat di lebih dari satu bab. Jangan mengubah nama bab atau ref. "
            "Kembalikan hanya objek JSON dengan setiap judul bab sebagai key dan array ref slide "
            "sebagai value. Bab tanpa materi harus berisi array kosong.\n"
            f"Key bab yang harus dipakai (contoh struktur kosong): {json.dumps(empty_mapping, ensure_ascii=False)}\n"
            f"Semua ref slide yang harus didistribusikan, total {len(all_refs)} ref, masing-masing tepat satu kali: "
            f"{json.dumps(all_refs, ensure_ascii=False)}\n\n"
            f"Isi presentasi:\n{slide_text}"
        )

        logger.info(
            "Mapping batch dimulai: %s slide, %s bab, max_attempts=%s",
            len(all_refs), len(chapters), self.MAP_SLIDES_MAX_ATTEMPTS,
        )
        logger.debug("Mapping expected refs: %s", sorted(expected))

        feedback = ""
        for attempt in range(1, self.MAP_SLIDES_MAX_ATTEMPTS + 1):
            _report(progress_callback, "mapping", "running", "", attempt - 1, self.MAP_SLIDES_MAX_ATTEMPTS)
            response = self._converse([{"text": base_prompt + feedback}], max_tokens=3000, step="mapping")
            logger.info(
                "Mapping percobaan %s/%s: panjang respons=%s karakter",
                attempt, self.MAP_SLIDES_MAX_ATTEMPTS, len(response),
            )
            try:
                parsed = self._parse_json(response)
                mapping: dict[str, list[str]] = {}
                seen: list[str] = []
                for chapter in chapters:
                    refs = parsed.get(chapter, [])
                    if not isinstance(refs, list) or any(not isinstance(ref, str) for ref in refs):
                        raise BedrockError(f"Format pemetaan untuk {chapter} tidak valid.")
                    mapping[chapter] = refs
                    seen.extend(refs)
                if set(seen) != expected or len(seen) != len(expected):
                    missing = sorted(expected - set(seen))
                    duplicated = sorted({ref for ref in seen if seen.count(ref) > 1})
                    logger.warning(
                        "Mapping percobaan %s GAGAL. Hilang (%s): %s. Duplikat (%s): %s. "
                        "Total ref diharapkan=%s, diterima=%s.",
                        attempt, len(missing), missing, len(duplicated), duplicated,
                        len(expected), len(seen),
                    )
                    raise BedrockError(
                        "Pemetaan slide tidak lengkap atau berisi slide duplikat. "
                        f"Hilang: {missing or '-'}. Duplikat: {duplicated or '-'}."
                    )
                logger.info("Mapping percobaan %s BERHASIL. Pemetaan per bab: %s", attempt, {k: len(v) for k, v in mapping.items()})
                _report(progress_callback, "mapping", "done", "", attempt, self.MAP_SLIDES_MAX_ATTEMPTS)
                return mapping
            except BedrockError as exc:
                if attempt == self.MAP_SLIDES_MAX_ATTEMPTS:
                    logger.error("Mapping GAGAL TOTAL setelah %s percobaan.", self.MAP_SLIDES_MAX_ATTEMPTS)
                    raise BedrockError(
                        "Pemetaan slide dari model tidak lengkap atau berisi slide duplikat "
                        f"setelah {self.MAP_SLIDES_MAX_ATTEMPTS} percobaan. Coba lagi."
                    ) from exc
                feedback = f"\n\nPercobaan sebelumnya gagal: {exc}. Perbaiki agar setiap ref muncul tepat satu kali."
        raise AssertionError("unreachable")

    NARRATION_MAX_ATTEMPTS = 2
    NARRATION_BATCH_SLIDES = 5
    NARRATION_BATCH_FACT_CHARS = 5000
    NARRATION_BATCH_ITEMS = 40  # jumlah butir yang boleh diminta sekaligus; membatasi panjang keluaran
    NARRATION_MAX_TOKENS = 8000

    def write_chapter_narration(
        self, chapter: str, slides: list[dict[str, Any]], style_text: str
    ) -> tuple[dict[str, list[str]], dict[str, list[str]]]:
        """Tulis narasi tiap slide sebuah bab. Bab besar dipecah per kelompok slide supaya keluaran model
        tidak terpotong di batas token. Mengembalikan (narasi per ref, angka yang tidak ada di fakta slide)."""
        narration: dict[str, list[str]] = {}
        flagged: dict[str, list[str]] = {}
        failed: list[dict[str, Any]] = []
        reasons: list[str] = []

        def run(batch: list[dict[str, Any]]) -> None:
            try:
                part, part_flagged = self._write_narration_batch(chapter, batch, style_text)
            except BedrockError as exc:  # kelompok yang gagal tidak membatalkan kelompok lain
                failed.extend(batch)
                reasons.append(str(exc))
                return
            narration.update(part)
            flagged.update(part_flagged)

        batch: list[dict[str, Any]] = []
        size = items = 0
        for slide in slides:
            chars = len(doc_composer.fact_text(slide))
            wanted = doc_composer.narration_limit(slide)
            if batch and (
                len(batch) >= self.NARRATION_BATCH_SLIDES
                or size + chars > self.NARRATION_BATCH_FACT_CHARS
                or items + wanted > self.NARRATION_BATCH_ITEMS
            ):
                run(batch)
                batch, size, items = [], 0, 0
            batch.append(slide)
            size += chars
            items += wanted
        if batch:
            run(batch)
        if failed:
            raise NarrationPartial(reasons[0], narration, flagged, failed)
        return narration, flagged

    def _write_narration_batch(
        self, chapter: str, slides: list[dict[str, Any]], style_text: str
    ) -> tuple[dict[str, list[str]], dict[str, list[str]]]:
        """Satu panggilan model untuk sekelompok slide. Bentuk atau angka yang salah diperbaiki sekali;
        sisanya dilaporkan."""
        refs = [slide["ref"] for slide in slides]
        optional = {slide["ref"] for slide in slides if doc_composer.may_skip(slide)}
        limits = {slide["ref"]: doc_composer.narration_limit(slide) for slide in slides}
        prompt = doc_composer.narration_prompt(chapter, slides, style_text)
        feedback = ""
        for attempt in range(1, self.NARRATION_MAX_ATTEMPTS + 1):
            last = attempt == self.NARRATION_MAX_ATTEMPTS
            response = self._converse([{"text": prompt + feedback}], max_tokens=self.NARRATION_MAX_TOKENS, step="narration")
            try:
                narration = doc_composer.parse_narration(
                    self._parse_json(response, context="narasi bab"), refs, optional, limits)
            except (BedrockError, doc_composer.NarrationError) as exc:
                if last:
                    raise BedrockError(f"Narasi bab '{chapter}' tidak valid: {exc}. Coba lagi.") from exc
                feedback = f"\n\nJawaban sebelumnya tidak valid ({exc}). Perbaiki; kembalikan hanya JSON sesuai bentuk."
                continue
            flagged = {}
            for slide in slides:
                numbers = doc_composer.unsupported_numbers(narration[slide["ref"]], doc_composer.all_facts(slide))
                if numbers:
                    flagged[slide["ref"]] = numbers
            if flagged and not last:
                lines = "; ".join(f"{ref}: {', '.join(numbers)}" for ref, numbers in flagged.items())
                feedback = (
                    f"\n\nAngka berikut pada jawaban sebelumnya TIDAK ada di fakta slide: {lines}. "
                    "Tulis ulang seluruh jawaban hanya dengan angka yang tertulis di fakta masing-masing slide."
                )
                continue
            return narration, flagged
        raise BedrockError(f"Narasi bab '{chapter}' gagal dibuat.")  # tidak tercapai

    def generate_style_guide(
        self, samples: list[dict[str, str]], statistics: dict[str, Any]
    ) -> tuple[dict[str, Any], int]:
        """Rangkum gaya penulisan referensi menjadi panduan. Mengembalikan (panduan, jumlah
        contoh yang dibuang karena bukan kutipan persis dari sampel)."""
        if not samples:
            raise BedrockError("Profil tidak memiliki sampel teks untuk dianalisis.")
        response = self._converse(
            [{"text": style_guide.build_prompt(samples, statistics)}], max_tokens=3000, step="style_guide"
        )
        guide = self._parse_json(response, context="panduan gaya")
        try:
            return style_guide.validate_guide(guide, samples)
        except style_guide.StyleGuideError as exc:
            raise BedrockError(f"Panduan gaya dari model tidak valid: {exc}. Coba lagi.") from exc
