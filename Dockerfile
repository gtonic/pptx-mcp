# Stage 1: Builder
FROM python:3.10-slim AS builder

WORKDIR /app

# Install build dependencies
RUN apt-get update \
    && apt-get install -y --no-install-recommends gcc libxml2-dev libxslt1-dev

# Copy requirements and install Python dependencies into a target directory
COPY requirements.txt .
RUN pip install --no-cache-dir --prefix=/install -r requirements.txt

# Stage 2: Final image
FROM python:3.10-slim

WORKDIR /app

# Copy only the installed dependencies from the builder stage
COPY --from=builder /install /usr/local

# Copy application code
COPY . .

# Clean up unnecessary files to reduce image size
RUN apt-get purge -y --auto-remove gcc libxml2-dev libxslt1-dev || true \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    && find /usr/local/lib/python3.10/site-packages -name "tests" -type d -exec rm -rf {} + \
    && find /usr/local/lib/python3.10/site-packages -name "__pycache__" -type d -exec rm -rf {} + \
    && rm -rf /root/.cache

VOLUME ["/data"]

EXPOSE 8081

CMD ["uvicorn", "server:app", "--host", "0.0.0.0", "--port", "8081"]
