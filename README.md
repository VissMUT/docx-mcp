# docx-mcp

MCP-сервер для работы с документами Microsoft Word (.docx) в среде Linux:
создание и редактирование документов, нативные формулы OMML, таблицы,
колонтитулы, режим исправлений (track changes), комментарии, сноски,
защита документа.

Транспорт — Streamable HTTP (FastMCP), поддерживается stateless-режим.
Файлы загружаются из S3-совместимого хранилища (rustfs) или по ссылке,
результат сохраняется в хранилище и возвращается в виде HTTPS-ссылки.

Проект основан на [word-mcp-live](https://github.com/ykarapazar/word-mcp-live)
(Yüce Karapazar), который в свою очередь основан на
[Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)
(GongRzhe). Лицензия — MIT, см. [LICENSE](LICENSE).

## Возможности

### Документы

- `docx_open` — открыть документ из вложения, HTTPS-ссылки или по имени
  файла из базы данных LobeChat; возвращает путь, используемый остальными
  инструментами;
- `docx_commit` — сохранить результат в хранилище и вернуть
  HTTPS-ссылку с контрольной суммой sha256;
- `docx_undo` — откат к последнему снимку (снимки создаются при
  открытии и фиксации);
- `docx_read` — структурное чтение: секции и поля, дерево заголовков,
  диапазоны абзацев с индексами, сетка таблиц, колонтитулы, изображения;
- `file_read` — чтение произвольных файлов (PDF, DOCX, XLSX, PPTX,
  EPUB, HTML, изображения) в Markdown;
- `create_document`, `copy_document`, `list_available_documents`,
  `get_document_info`, `get_document_text`, `get_document_outline`,
  `merge_documents`, `convert_to_pdf`.

### Генерация

- `docx_build` — сборка документа из pandoc-markdown (таблицы,
  формулы `$...$` / `$$...$$`, оглавление, нумерация заголовков,
  добавление разделов через `append`);
- `docx_create` — пустой документ с базовым оформлением
  (Times New Roman 12, поля, номер страницы в колонтитуле);
- `docx_fill` — заполнение шаблона docxtpl (переменные `{{var}}`,
  циклы, условия) из JSON;
- `docx_append_block` — добавление блока (заголовки, абзацы, формулы)
  в конец документа одним вызовом;
- `add_heading`, `add_paragraph`, `add_table`, `add_picture`,
  `add_page_break`, списки (`insert_numbered_list_near_text`).

### Формулы

- `docx_formula` — нативные формулы Word (OMML), LaTeX-нотация;
  операции: список, вставка (в тело документа или в ячейку таблицы),
  замена, удаление, очистка ячейки;
- display-формулы оформляются таблицами-врезками 1×2 без границ:
  формула по центру, номер — справа;
- `docx_number_formulas` — сквозная нумерация формул «(N)» обычным
  текстом (без полей SEQ, номер виден сразу, F9 не требуется).

### Текст и правки

- `search_and_replace`, `format_text`, `delete_paragraph`,
  `insert_line_or_paragraph_near_text`, `insert_header_near_text`,
  `replace_paragraph_block_below_header`,
  `replace_block_between_manual_anchors`;
- режим исправлений: `track_replace`, `track_insert`, `track_delete`,
  `list_tracked_changes`, `accept_tracked_changes`,
  `reject_tracked_changes`.

### Таблицы

- `docx_replace_table` — замена таблицы целиком новой из
  markdown-разметки;
- форматирование: `format_table`, `set_table_cell_shading`,
  `apply_table_alternating_rows`, `highlight_table_header`,
  `set_table_cell_alignment`, `format_table_cell_text`,
  `set_table_cell_padding`;
- объединение ячеек: `merge_table_cells`,
  `merge_table_cells_horizontal`, `merge_table_cells_vertical`;
- ширина: `set_table_column_width(s)`, `set_table_width`,
  `auto_fit_table_columns`.

### Оформление

- колонтитулы и нумерация страниц: `add_header_footer`,
  `add_page_numbers`, `set_page_layout`, `add_section_break`,
  `add_watermark`;
- оглавление `add_table_of_contents`, закладки `add_bookmark`;
- подписи и ссылки: `docx_caption` (нумерация таблиц и рисунков),
  `docx_crossref` (поля REF/PAGEREF), `docx_docproperty`
  (DOCPROPERTY-поля);
- выделение цветом `docx_highlight` (поиск и перекраска),
  смена шрифта `docx_replace_font`;
- сноски и концевые сноски, комментарии, гиперссылки,
  пользовательские стили `create_custom_style`.

### Защита

- `protect_document`, `unprotect_document`,
  `add_restricted_editing`, `add_digital_signature`,
  `verify_document`.

### Режим live-редактирования (Windows)

Инструменты `word_live_*` управляют открытым документом в Microsoft
Word через COM-интерфейс: вставка и форматирование текста, таблицы,
закладки, перекрёстные ссылки, ревизии, снимки и откат. Работают
только в среде Windows с установленным Microsoft Word; в Linux
недоступны.

## Примеры

### Сборка документа из markdown

`docx_build` принимает текст в диалекте pandoc: таблицы,
формулы `$...$` (в строке) и `$$...$$` (отдельной строкой),
разрыв страницы `\newpage`.

```markdown
# Пояснительная записка

## Исходные данные

| Параметр | Значение |
|---|---|
| Напряжение питания | 24 В |
| Коэффициент запаса | 1,2 |

## Расчет

Суммарный ток нагрузки:

$$I = \frac{P}{U} = \frac{122}{24} \approx 5{,}1\ \text{А}$$
```

```json
{"ok": true, "filename": "/tmp/doc-editor/pz.docx",
 "info": {"paragraphs": 12, "tables": 2, "formulas": 1, "sections": 1}}
```

Параметры `toc=true` (оглавление) и `number_sections=true`
(нумерация заголовков) включаются при необходимости. Display-формулы
автоматически оформляются врезками: формула по центру, номер справа.

### Добавление раздела с формулами

`docx_append_block` добавляет заголовки, абзацы и формулы в конец
документа одним вызовом:

```json
[
  {"type": "heading", "level": 2,
   "text": "Расчет источника бесперебойного питания"},
  {"type": "p", "text": "Суммарная мощность нагрузки:"},
  {"type": "formula", "latex": "P_n = \\sum_i P_i", "display": true},
  {"type": "p", "text": "Принимаем АКБ 24 В, 9 А·ч."}
]
```

```json
{"ok": true, "paragraphs_added": 3, "formulas_added": 1}
```

После изменения состава формул выполняется `docx_number_formulas()`:
номера «(N)» проставляются обычным текстом в правой ячейке врезки.

### Правка существующего документа

1. `docx_open(file_XXXX)` — открыть вложение чата, HTTPS-ссылку
   или файл по имени из базы LobeChat;
2. `docx_read(scope="structure")` — получить карту документа
   (секции, заголовки, таблицы, поля);
3. `search_and_replace("старый текст", "новый текст")` — внести правку;
4. `docx_commit()` — сохранить результат; в ответе возвращается
   публичная ссылка и контрольная сумма:

```json
{"ok": true, "url": "https://xais.ru/lobe/doc-editor/3f2a1c00-pz.docx",
 "file": "pz.docx", "sha256": "a4a9c67062decfcb8d83054bcf66b7a59de59f628ec4215d45f7dd3b100f1888",
 "size": 12720}
```

### Режим исправлений

Для согласования изменений используются инструменты `track_replace`,
`track_insert`, `track_delete`. Завершение цикла правок:

```json
{"success": true, "accepted": 20}
```

`accept_tracked_changes()` принимает все исправления,
`reject_tracked_changes()` отклоняет их, `list_tracked_changes()`
возвращает перечень с авторами и датами.

## Архитектура

MCP-клиент обращается к серверу по Streamable HTTP (порт 9113, FastMCP
в stateless-режиме). Обработка выполняется в рабочем каталоге
`/tmp/doc-editor`.

- Источник файла: S3-совместимое хранилище (rustfs, ключ из таблицы
  `files` базы LobeChat) или HTTPS-ссылка.
- Сохранение результата (`docx_commit`): загрузка в rustfs (bucket
  `lobe`) по ключу `doc-editor/<uuid>-<имя>`, публичная ссылка имеет
  вид `https://<PUBLIC_BASE_URL>/lobe/doc-editor/<uuid>-<имя>`.

Сервер не хранит состояние между вызовами: открытый файл задаётся
путем в аргументах инструментов либо запоминается для текущей сессии.
Снимки для `docx_undo` хранятся в `/tmp/doc-editor/snapshots`.

## Требования

- Python 3.12;
- pandoc (конвертация LaTeX в OMML и markdown в docx);
- anymd (чтение PDF/XLSX/PPTX/HTML в Markdown);
- PostgreSQL с таблицей `files` LobeChat (поиск файлов по id/имени);
- S3-совместимое хранилище (rustfs).

## Переменные окружения

| Переменная | Назначение |
|---|---|
| `MCP_TRANSPORT` | Транспорт: `streamable-http` (по умолчанию) или `sse` |
| `MCP_HOST`, `MCP_PORT`, `MCP_PATH` | Адрес и путь сервера (по умолчанию `0.0.0.0:9113/sse`) |
| `MCP_AUTHOR`, `MCP_AUTHOR_INITIALS` | Автор изменений (track changes, комментарии) |
| `PGHOST`, `PGDATABASE`, `PGUSER`, `PGPASSWORD` | База данных LobeChat |
| `RUSTFS_HOST`, `RUSTFS_BUCKET` | Хранилище файлов |
| `RUSTFS_ACCESS_KEY`, `RUSTFS_SECRET_KEY` | Доступ к хранилищу |
| `PUBLIC_BASE_URL` | База публичных ссылок (по умолчанию `https://xais.ru`) |
| `S3_KEY_PREFIX` | Префикс ключей при сохранении (по умолчанию `doc-editor`) |
| `WORK_DIR` | Рабочий каталог (по умолчанию `/tmp/doc-editor`) |
| `REFERENCE_DOCX` | Эталон стилей pandoc для `docx_build` |

## Запуск

Сборка образа:

```bash
docker build -t docx-mcp .
```

Запуск контейнера:

```bash
docker run --name docx-mcp --restart unless-stopped \
  -p 127.0.0.1:9113:9113 \
  -e MCP_TRANSPORT=streamable-http \
  -e MCP_HOST=0.0.0.0 -e MCP_PORT=9113 -e MCP_PATH=/sse \
  -e PGHOST=<host> -e PGDATABASE=lobechat -e PGUSER=<user> \
  -e PGPASSWORD=<password> \
  -e RUSTFS_HOST=<host> -e RUSTFS_BUCKET=<bucket> \
  -e RUSTFS_ACCESS_KEY=<key> -e RUSTFS_SECRET_KEY=<secret> \
  docx-mcp:latest
```

Пример systemd-юнита — в файле `doc-editor.service`.

## Подключение к MCP-клиенту

```json
{
  "mcpServers": {
    "docx-mcp": {
      "url": "http://127.0.0.1:9113/sse"
    }
  }
}
```

## Лицензия

MIT. Код основан на проектах
[word-mcp-live](https://github.com/ykarapazar/word-mcp-live) и
[Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)
(MIT); их авторство сохранено в LICENSE.
