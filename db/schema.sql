-- Tabel log waktu dan token per langkah pembuatan DOCX.
-- Nama tabel mengikuti env PROCESSING_LOG_TABLE (default: document_processing_log_v2).
-- Jalankan sekali:  psql -d <database> -f db/schema.sql
CREATE TABLE IF NOT EXISTS document_processing_log_v2 (
    id                     BIGSERIAL PRIMARY KEY,
    created_at             TIMESTAMPTZ NOT NULL DEFAULT now(),
    document_name          TEXT NOT NULL,
    status                 TEXT NOT NULL,           -- done | cancelled | error

    ingest_seconds         NUMERIC(10, 2),
    ingest_at              TIMESTAMPTZ,
    ingest_input_tokens    BIGINT,
    ingest_output_tokens   BIGINT,

    visuals_seconds        NUMERIC(10, 2),          -- deskripsi visual slide
    visuals_at             TIMESTAMPTZ,
    visuals_input_tokens   BIGINT,
    visuals_output_tokens  BIGINT,

    captions_seconds       NUMERIC(10, 2),          -- ringkasan gambar untuk caption
    captions_at            TIMESTAMPTZ,
    captions_input_tokens  BIGINT,
    captions_output_tokens BIGINT,

    mapping_seconds        NUMERIC(10, 2),
    mapping_at             TIMESTAMPTZ,
    mapping_input_tokens   BIGINT,
    mapping_output_tokens  BIGINT,

    writing_seconds        NUMERIC(10, 2),          -- narasi bab (dan panduan gaya)
    writing_at             TIMESTAMPTZ,
    writing_input_tokens   BIGINT,
    writing_output_tokens  BIGINT,

    total_seconds          NUMERIC(10, 2),          -- termasuk langkah build (tanpa kolom sendiri)
    total_input_tokens     BIGINT,
    total_output_tokens    BIGINT
);

CREATE INDEX IF NOT EXISTS document_processing_log_v2_created_at_idx
    ON document_processing_log_v2 (created_at DESC);
