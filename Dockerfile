FROM python:3.11-slim

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        poppler-utils \
        fonts-dejavu-core \
        fontconfig \
    && rm -rf /var/lib/apt/lists/*

COPY webappGamilAPI/requirements.txt /app/webappGamilAPI/requirements.txt
RUN python -m pip install --no-cache-dir -r /app/webappGamilAPI/requirements.txt

COPY . /app

RUN mkdir -p /app/webappGamilAPI/output

EXPOSE 8000

CMD ["python", "-m", "uvicorn", "webappGamilAPI.app:app", "--host", "0.0.0.0", "--port", "8000"]
