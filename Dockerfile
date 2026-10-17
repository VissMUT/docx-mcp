FROM python:3.12-slim

# pandoc — для формул LaTeX <-> OMML (docx_formula) и генерации docx_build
# anymd v8.5.3 — чтение любых файлов в Markdown (file_read); бинарник из
# релиза GitHub (на PyPI пакета нет, npm-версия требует Node)
# ⚠️ Acquire::ForceIPv4 — apt в docker-build виснет на IPv6 (deb.debian.org отдаёт
# AAAA первым, а маршрута v6 нет); с форсом апдейт идёт 2с вместо вечного зависания
RUN echo 'Acquire::ForceIPv4 "true";' > /etc/apt/apt.conf.d/99force-ipv4 \
 && echo 'precedence ::ffff:0:0/96 100' >> /etc/gai.conf \
 && apt-get update \
 && apt-get install -y --no-install-recommends pandoc curl ca-certificates \
 && curl -fsSL -o /tmp/anymd.tar.gz \
      https://github.com/SylphxAI/anymd/releases/download/v8.5.3/anymd-linux-x64-gnu.tar.gz \
 && tar -xzf /tmp/anymd.tar.gz -C /usr/local/bin \
 && chmod +x /usr/local/bin/anymd \
 && rm -f /tmp/anymd.tar.gz \
 && rm -rf /var/lib/apt/lists/* \
  && pip install --no-cache-dir \
      "fastmcp>=4,<5" "python-docx==1.2.0" msoffcrypto-tool docx2pdf python-dotenv \
      boto3 psycopg2-binary docxtpl

# Эталон стилей pandoc (docx_build): TNR 12, поля A4, PAGE в футере
# /assets — НЕ примонтирован томом, генерится только при сборке
COPY tools/make_reference.py /tmp/make_reference.py
RUN mkdir -p /assets \
 && python3 /tmp/make_reference.py /assets/reference.docx \
 && rm -f /tmp/make_reference.py

WORKDIR /app
COPY . .

EXPOSE 9113
CMD ["python3", "word_mcp_server.py"]
