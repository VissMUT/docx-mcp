# docx-mcp

MCP-сервер для создания и редактирования документов Word (.docx) агентами LobeChat.
Работает в Docker на Linux (streamable HTTP, stateless), файлы забирает из вложений
чата (rustfs) и возвращает https-ссылкой в чат.

Основан на [ykarapazar/word-mcp-live](https://github.com/ykarapazar/word-mcp-live)
(в свою очередь — форк [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)),
лицензия MIT. Это самостоятельное репо с собственной историей.

## Наши добавления поверх word-mcp-live

| Тулы | Назначение |
|---|---|
| `docx_open` / `docx_commit` / `docx_undo` | Вложение чата (`file_...`), https-ссылка или имя файла из БЗ → рабочий каталог; результат → rustfs → https-ссылка + sha256; снапшоты/откат |
| `docx_formula` | Нативные формулы OMML: list / insert / replace / replace_cell / delete / clear_cell (LaTeX ↔ OMML через pandoc). Display-формулы — таблицы-врезки 1×2 (по центру + номер справа) |
| `docx_number_formulas` | Сквозная нумерация display-формул простым текстом «(N)» (без SEQ-полей, F9 не нужен); legacy-абзацы мигрируют во врезки |
| `docx_append_block` | Добавление раздела (заголовки + абзацы + формулы) в конец документа одним вызовом |
| `docx_read` / `file_read` | Структурное чтение (индексы согласованы с тулами правок) / любой файл (PDF, XLSX, PPTX, HTML) как Markdown через anymd |
| `docx_create` / `docx_build` / `docx_fill` | Пустой документ / сборка из pandoc-markdown (toc, number_sections, врезки формул) / заполнение шаблона docxtpl |
| СПДС-утилиты | `docx_highlight`, `docx_caption`, `docx_crossref`, `docx_docproperty`, `docx_replace_font`, `docx_replace_table` |

## Развёртывание (Docker, Linux)

```bash
docker build -t docx-mcp .
# systemd-юнит: doc-editor.service (пример из нашего деплоя)
docker run --name docx-mcp --restart unless-stopped \
  -p 127.0.0.1:9113:9113 \
  -e MCP_TRANSPORT=streamable-http -e MCP_HOST=0.0.0.0 -e MCP_PORT=9113 \
  -e MCP_PATH=/sse -e MCP_AUTHOR=... \
  docx-mcp:latest
```

Требования: доступ к БД LobeChat (PostgreSQL, таблица `files`) и rustfs (S3) —
переменные `PGHOST/PGDATABASE/PGUSER/PGPASSWORD` и `RUSTFS_*` (см. `doc-editor.service`).

## Технологии

- python-docx + lxml (OOXML), pandoc (LaTeX ↔ OMML, markdown → docx)
- FastMCP (streamable-http, stateless), anymd, docxtpl
