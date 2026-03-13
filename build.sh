#!/usr/bin/env bash
set -o errexit

# Install system packages (LibreOffice for PPTX->PDF, poppler for PDF->images)
apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    poppler-utils

# Install Python dependencies
pip install --upgrade pip
pip install -r requirements.txt

# Install Playwright Chromium with system deps
playwright install --with-deps chromium

# Create required directories
mkdir -p data outputs jobs
