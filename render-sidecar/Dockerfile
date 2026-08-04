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

# fonts-liberation covers Arial/Times New Roman/Courier New (metric-
# compatible substitutes for those three specifically). It does NOT cover
# Calibri/Cambria -- the default fonts in every "Office"-themed deck since
# PowerPoint 2007 -- so any title/body text using them had no matching
# substitute at all. Without a metric-compatible substitute, LibreOffice
# falls back to a generic font whose glyphs are narrower than Calibri's,
# and inflates inter-character spacing trying to preserve the original
# layout's line width -- this is what produced the widely-spaced,
# "stretched" text (confirmed by reproducing it locally: a plain Calibri
# title/subtitle textbox render with exactly that spacing artifact, gone
# once these packages are installed). fonts-crosextra-carlito/-caladea are
# Google's metric-compatible substitutes for Calibri/Cambria respectively
# -- same recipe used by Collabora Online and other headless-LO PDF/image
# conversion services, no EULA acceptance needed (unlike ttf-mscorefonts-
# installer, which would need extra non-interactive-debconf plumbing for
# real Arial/Times New Roman -- not worth it here since Liberation already
# covers those two adequately).
# fonts-open-sans: the real thing, not a substitute -- Open Sans/Open Sans
# ExtraBold are the actual theme fonts used across this app's own content
# library source decks (confirmed by grepping every theme1.xml under
# Package Files/StrategyToolbar/BoilerPlates/16x9/ for <a:latin typeface>:
# Calibri/Calibri Light/Open Sans/Open Sans ExtraBold/Arial are the only
# five fonts in use anywhere in that library), so there's no metric-
# compatible-substitute step needed here the way there is for Calibri --
# just install the font itself, freely licensed (Apache 2.0), no EULA.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress \
    imagemagick \
    poppler-utils \
    fonts-liberation \
    fonts-crosextra-carlito \
    fonts-crosextra-caladea \
    fonts-open-sans \
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
