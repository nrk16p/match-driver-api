# Dockerfile
FROM python:3.10-slim

# set working directory
WORKDIR /app

# install system dependencies (if needed)
RUN apt-get update && \
    apt-get install -y --no-install-recommends gcc && \
    rm -rf /var/lib/apt/lists/*

# copy and install python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# copy your app
COPY . .

# tell Vercel to use the PORT env var
CMD ["gunicorn", "app:app", "-b", "0.0.0.0:${PORT:-5000}"]
