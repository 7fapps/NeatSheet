# Dockerfile for NeatSheet
# This defines the exact environment NeatSheet runs in when deployed,
# including Tesseract (the OCR engine), which can't be installed via
# pip alone since it's a separate system program, not a Python library.

FROM python:3.11-slim

# Install Tesseract OCR (the system-level program) plus its English
# language data. apt-get is Linux's package installer, similar in
# spirit to pip but for non-Python programs.
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        tesseract-ocr \
        tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

# Set the working directory inside the container
WORKDIR /app

# Copy and install Python dependencies first (this lets Docker cache
# this step, so future deploys are faster if requirements.txt hasn't changed)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application code
COPY . .

# Render provides a PORT environment variable - our app must listen on it.
ENV PORT=10000
EXPOSE 10000

# Run the app using gunicorn (a production-grade server),
# instead of Flask's built-in development server.
CMD gunicorn --bind 0.0.0.0:$PORT app:app
