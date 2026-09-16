FROM python:3.11-slim
WORKDIR /app
RUN apt-get update && apt-get install -y --no-install-recommends \
libpango-1.0-0 libpangoft2-1.0-0 libffi-dev \
&& rm -rf /var/lib/apt/lists/*
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
# FIX (16/09/2026): era "--workers 2 --threads 4 --timeout 120".
# 2 workers = 2 processos Python, cada um com seu heap (~170 MiB no total
# em repouso) -- apertado demais no antigo limite de 512 MiB, contribuiu
# pro OOM que derrubava pedidos grandes. O sistema é I/O-bound (passa o
# tempo esperando o bucket do GCS), então 1 worker com mais threads atende
# igual gastando metade da memoria; o Cloud Run escala por instancia, nao
# por worker. Timeout subiu de 120s pra 600s pra alinhar com o timeout de
# requisicao do Cloud Run (senao o gunicorn mata pedido grande antes, com
# 502 confuso).
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 600 main:app
