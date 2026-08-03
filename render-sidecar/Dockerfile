# Debian-based, deliberately not Alpine -- LibreOffice has a documented,
# unfixable incompatibility with Alpine's musl libc (crashes with
# com::sun::star::uno::RuntimeException, exit 134; see Alpine's own GitLab
# issue #9488). Isolating that dependency in this single-purpose sidecar
# keeps the main app's Alpine image untouched, and keeps LibreOffice's much
# larger dependency footprint (and CVE surface) confined to a container
# that has no database credentials, no persistent volume, and no published
# port -- see ../docker-compose.yml's `render` service for the network
# isolation.
FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress \
    imagemagick \
    poppler-utils \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py .

# Nothing in this container is worth root -- it has no data of its own.
RUN useradd --create-home --uid 1000 render
USER render

EXPOSE 8090
HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 \
  CMD python3 -c "import urllib.request; urllib.request.urlopen('http://localhost:8090/health', timeout=3)" || exit 1

CMD ["waitress-serve", "--host=0.0.0.0", "--port=8090", "--threads=2", "app:app"]
