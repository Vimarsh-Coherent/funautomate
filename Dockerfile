FROM python:3.11-slim

# Install system dependencies: LibreOffice for PPTX->PDF, poppler for PDF->images,
# curl for healthcheck
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    poppler-utils \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install Playwright Chromium + its system dependencies
RUN playwright install --with-deps chromium

# Copy application code
COPY . .

# Create required directories
RUN mkdir -p data outputs jobs

# Render uses PORT env variable (default 10000)
ENV PORT=10000
EXPOSE ${PORT}

HEALTHCHECK CMD curl --fail http://localhost:${PORT}/_stcore/health || exit 1

ENTRYPOINT streamlit run app.py \
    --server.port=${PORT} \
    --server.address=0.0.0.0 \
    --server.enableCORS=false \
    --server.enableXsrfProtection=false \
    --server.headless=true
