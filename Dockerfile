FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends patch \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

COPY app.py .
COPY deploy/eai_patch_00 deploy/eai_patch_01 deploy/eai_patch_02 deploy/eai_patch_03 deploy/eai_patch_04 ./

RUN sed -i 's/\r$//' app.py \
    && cat eai_patch_00 eai_patch_01 eai_patch_02 eai_patch_03 eai_patch_04 > eai_cloudrun.patch \
    && patch app.py eai_cloudrun.patch \
    && rm eai_patch_00 eai_patch_01 eai_patch_02 eai_patch_03 eai_patch_04 eai_cloudrun.patch

EXPOSE 8080

CMD streamlit run app.py \
    --server.address=0.0.0.0 \
    --server.port=${PORT:-8080} \
    --server.headless=true \
    --server.enableCORS=false \
    --server.enableXsrfProtection=true \
    --browser.gatherUsageStats=false
