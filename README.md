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

| Инструмент | Назначение |
|---|---|
| `docx_open` | Открыть документ из вложения, HTTPS-ссылки или по имени файла из базы LobeChat |
| `docx_commit` | Сохранить результат в хранилище, вернуть ссылку и контрольную сумму sha256 |
| `docx_undo` | Откат к последнему снимку |
| `docx_read` | Структурное чтение: секции, поля, заголовки, абзацы с индексами, таблицы, колонтитулы, изображения |
| `file_read` | Чтение PDF, DOCX, XLSX, PPTX, EPUB, HTML и изображений в Markdown |
| `create_document` | Новый документ с метаданными |
| `copy_document` | Копия документа |
| `list_available_documents` | Список документов в каталоге |
| `get_document_info` | Свойства и статистика документа |
| `get_document_text` | Полный текст документа |
| `get_document_outline` | Структура документа (заголовки) |
| `merge_documents` | Объединение нескольких документов |
| `convert_to_pdf` | Конвертация в PDF |

### Генерация

| Инструмент | Назначение |
|---|---|
| `docx_build` | Сборка документа из pandoc-markdown: таблицы, формулы, оглавление, нумерация заголовков, добавление разделов через `append` |
| `docx_create` | Пустой документ с базовым оформлением (шрифт, поля, номер страницы в колонтитуле) |
| `docx_fill` | Заполнение шаблона docxtpl (переменные, циклы, условия) из JSON |
| `docx_append_block` | Блок (заголовки, абзацы, формулы) в конец документа одним вызовом |
| `add_heading` | Заголовок уровня 1–9 |
| `add_paragraph` | Абзац |
| `add_table` | Таблица из данных |
| `add_picture` | Изображение с масштабированием |
| `add_page_break` | Разрыв страницы |
| `insert_numbered_list_near_text` | Маркированный или нумерованный список |
| `insert_line_or_paragraph_near_text` | Строка или абзац до/после указанного текста |
| `insert_header_near_text` | Заголовок до/после указанного текста |

### Формулы

| Инструмент | Назначение |
|---|---|
| `docx_formula` | Нативные формулы Word (OMML) в LaTeX-нотации: список, вставка (в тело документа или в ячейку таблицы), замена, удаление, очистка ячейки |
| `docx_number_formulas` | Сквозная нумерация формул «(N)» обычным текстом (без полей SEQ, F9 не требуется) |

Display-формулы оформляются таблицами-врезками 1×2 без границ:
формула по центру, номер справа.

### Текст и правки

| Инструмент | Назначение |
|---|---|
| `search_and_replace` | Поиск и замена по всему документу |
| `format_text` | Форматирование фрагмента текста (шрифт, начертание, цвет) |
| `delete_paragraph` | Удаление абзаца по индексу |
| `track_replace` / `track_insert` / `track_delete` | Правки в режиме исправлений |
| `list_tracked_changes` | Список исправлений |
| `accept_tracked_changes` / `reject_tracked_changes` | Принятие и отклонение исправлений |
| `replace_paragraph_block_below_header` | Замена блока абзацев под заголовком |
| `replace_block_between_manual_anchors` | Замена блока между ручными якорями |
| `get_highlighted_text` | Текст с маркерной подсветкой |
| `find_text_in_document` | Поиск с учётом регистра и целых слов |
| `get_paragraph_text_from_document` | Текст абзаца по индексу |
| `create_custom_style` | Пользовательский стиль |

### Таблицы

| Инструмент | Назначение |
|---|---|
| `docx_replace_table` | Замена таблицы целиком новой из markdown-разметки |
| `format_table` | Границы и стиль таблицы |
| `set_table_cell_shading` | Заливка ячейки |
| `apply_table_alternating_rows` | Чередование цветов строк |
| `highlight_table_header` | Выделение строки заголовка |
| `set_table_cell_alignment` / `set_table_alignment_all` | Горизонтальное и вертикальное выравнивание |
| `format_table_cell_text` | Форматирование текста ячейки |
| `set_table_cell_padding` | Внутренние отступы ячейки |
| `merge_table_cells` / `merge_table_cells_horizontal` / `merge_table_cells_vertical` | Объединение ячеек |
| `set_table_column_width` / `set_table_column_widths` | Ширина колонок |
| `set_table_width` / `auto_fit_table_columns` | Ширина таблицы и автоподбор |

### Оформление

| Инструмент | Назначение |
|---|---|
| `add_header_footer` | Колонтитулы |
| `add_page_numbers` | Номера страниц |
| `add_table_of_contents` | Оглавление |
| `add_bookmark` | Закладка |
| `docx_caption` | Подписи таблиц и рисунков с автонумерацией |
| `docx_crossref` | Перекрёстные ссылки (поля REF/PAGEREF) |
| `docx_docproperty` | DOCPROPERTY-поля |
| `docx_highlight` | Выделение текста цветом (поиск и перекраска) |
| `docx_replace_font` | Смена шрифта по документу |
| `set_page_layout` / `add_section_break` | Параметры страницы и разрывы разделов |
| `add_watermark` | Водяной знак |
| `add_footnote_*` / `add_endnote_*` | Сноски и концевые сноски |
| `add_comment` / `get_all_comments` / `get_comments_by_author` / `get_comments_for_paragraph` | Комментарии |
| `manage_hyperlinks` | Гиперссылки |

### Защита

| Инструмент | Назначение |
|---|---|
| `protect_document` / `unprotect_document` | Парольная защита документа |
| `add_restricted_editing` | Ограничение редактирования отдельными разделами |
| `add_digital_signature` | Цифровая подпись |
| `verify_document` | Проверка защиты и подписи |

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
# Отчет

## Исходные данные

| Показатель | Значение |
|---|---|
| Объем выборки | 120 |
| Период наблюдения | январь-март |

## Расчет

Среднее значение:

$$\bar{x} = \frac{1}{n} \sum_{i=1}^{n} x_i$$
```

```json
{"ok": true, "filename": "/tmp/doc-editor/report.docx",
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
  {"type": "heading", "level": 2, "text": "Анализ данных"},
  {"type": "p", "text": "Среднее значение по выборке:"},
  {"type": "formula", "latex": "\\bar{x} = \\frac{1}{n} \\sum_{i=1}^{n} x_i", "display": true},
  {"type": "p", "text": "Результат округляется до двух знаков после запятой."}
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
{"ok": true, "url": "https://<host>/<bucket>/doc-editor/<uuid>-report.docx",
 "file": "report.docx", "sha256": "<sha256>", "size": 12720}
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
| `PUBLIC_BASE_URL` | База публичных ссылок (адрес, с которого доступно хранилище) |
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
