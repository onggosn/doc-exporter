FROM registry.access.redhat.com/ubi9/python-311:latest

USER 0
# LibreOffice (PPTX -> PDF) untuk render slide visual; PDF -> PNG memakai PyMuPDF (pip).
# Jika repo EPEL/Debian diperlukan untuk libreoffice di base image Anda, gunakan
# base image Debian (python:3.11-slim) dengan:
#   apt-get install -y --no-install-recommends libreoffice-impress fonts-liberation
RUN dnf install -y --setopt=install_weak_deps=False \
        libreoffice-impress liberation-fonts dejavu-sans-fonts \
    && dnf clean all

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

# OpenShift menjalankan pod dengan UID acak (group 0): direktori kerja harus
# bisa ditulis oleh group 0, dan HOME diarahkan ke /tmp untuk LibreOffice.
RUN mkdir -p /data/uploads /data/logs \
    && chgrp -R 0 /app /data && chmod -R g=u /app /data
ENV HOME=/tmp \
    HOST=0.0.0.0 \
    PORT=8080 \
    UPLOAD_DIR=/data/uploads \
    LOG_FILE=/data/logs/proses.log \
    PYTHONUNBUFFERED=1
USER 1001
EXPOSE 8080

# PENTING: 1 worker. Sesi mapping dan job DOCX disimpan di memori proses
# (MAPPING_SESSIONS), jadi multi-worker/multi-replica akan memecah sesi.
CMD ["gunicorn", "-w", "1", "--threads", "8", "--timeout", "900", "-b", "0.0.0.0:8080", "app:app"]
