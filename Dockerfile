FROM python:3.11.9-slim-bookworm

WORKDIR /app
ENV PYTHONUNBUFFERED=1 \
    SAL_USE_VCLPLUGIN=svp \
    OMP_NUM_THREADS=1

# Impress renders the actual PPTX; compatible fonts prevent fallback reflow.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress \
    fonts-dejavu-core fonts-liberation2 \
    fonts-crosextra-carlito fonts-crosextra-caladea \
    fonts-noto-core fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

# No network/model calls: prove PDF rendering works before replacing the live bot.
RUN python tools/check_presentation_pdf_runtime.py

CMD ["python", "main.py"]
