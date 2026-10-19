FROM python:3.11.9-slim-bookworm
WORKDIR /app
ENV PYTHONUNBUFFERED=1 \
    SAL_USE_VCLPLUGIN=svp \
    OMP_NUM_THREADS=1
# Render saved office documents locally with compatible fonts.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress libreoffice-writer \
    fonts-dejavu-core fonts-liberation2 \
    fonts-crosextra-carlito fonts-crosextra-caladea \
    fonts-noto-core fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
# No network/model calls: check both export filters before replacing the live bot.
RUN python tools/check_academic_runtime.py && python tools/check_word_field_input.py && python tools/check_presentation_pdf_runtime.py && python tools/check_word_pdf_runtime.py && python tools/check_car_sale_runtime.py && python tools/check_fixed_word_runtime.py
CMD ["python", "main.py"]
