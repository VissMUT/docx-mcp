# doc-editor — редактор Word для агентов LobeChat

MCP-сервер «Редактор документов» (форк word-mcp-live). Агент читает DOCX
из вложения чата, вносит правки тулами и возвращает исправленный файл
ссылкой. 126 инструментов: текст, стили, таблицы, колонтитулы, номера
страниц, TOC, сноски, комментарии, запись изменений, формулы, чтение.

## Основной сценарий

1. **`docx_open(file_ref)`** — открой документ: передай id вложения
   (`file_...`), https-ссылку или имя файла из БЗ. Ответ содержит
   `filename` (путь) — его используй во всех последующих тулах.
2. Вноси правки тулами (см. ниже).
3. **`docx_commit()`** — получи `url` исправленного файла и отдай
   пользователю.
4. **`docx_undo()`** — откат к состоянию на момент docx_open/commit.

## Чтение (сначала — это, экономит токены)

| Тул | Что даёт |
|---|---|
| **`docx_read(scope)`** | Наш тул, JSON. `scope`: `structure` (секции/поля/дерево заголовков/счётчики/таблицы с первой строкой), `paragraphs` (диапазон `start`..`end` с индексами+стилями), `table` (сетка ячеек таблицы `table`), `headers_footers`, `fields` (PAGE/NUMPAGES/…), `images`. Индексы = нумерация python-docx (для edit-тулов) |
| **`file_read(source)`** | Наш тул. ЛЮБОЙ файл → Markdown (anymd): PDF, XLSX, PPTX, EPUB, HTML, изображения. Таблицы — markdown-таблицы, формулы — LaTeX, страницы PDF — `<!-- page N -->`. `pages` — диапазон страниц PDF («1-5,8»), `max_tokens` — бюджет. Для больших файлов читай по диапазонам страниц |

## Генерация документов (наши тулы)

| Тул | Когда |
|---|---|
| **`docx_build(markdown, filename, toc?, number_sections?, append?)`** | Сплошной текст: ПЗ, программы, разделы. Пишешь весь документ pandoc-markdown'ом: pipe-таблицы, формулы `$...$`/`$$...$$` (или `\(...\)`/`\[...\]`), `\newpage` — разрыв страницы. Лимит 200 КБ за вызов. `append=true` — доклеить markdown к буферу и пересобрать (длинные документы частями). Стили — эталонные (TNR 12, justify, отступ 1.25, заголовки 13-14 bold, A4 20/10/20/10, PAGE в футере) |
| **`docx_fill(template_ref, context, filename?)`** | Формы с фиксированной сеткой: титул, спецификация, опросный лист. Шаблон (вложение/БЗ) с `{{var}}`, `{% for %}` — + JSON-контекст. Результат — НОВЫЙ файл, шаблон не затирается. В ответе `undeclared_variables` — недозаполненные плейсхолдеры. Шаблон титула/штампа — «Базовые шаблоны ППр_2026» в БЗ (file_9AKDkNoSDRJV) |
| **`docx_create(name, title?, author?)`** | Пустой docx с базовыми стилями под ручную сборку коротких/нестандартных документов |

После генерации файл уже открыт (в ответе `filename`): `docx_open` не нужен — сразу правки → `docx_commit`.

## СПДС-утилиты (нумерация, ссылки, подсветка)

| Тул | Что делает |
|---|---|
| **`docx_highlight(text?, color?, paragraph?, old_color?)`** | Подсветка маркером (как в Word): yellow/green/cyan/magenta/blue/red/…, `none` — снять. Два режима: поиск по `text` (везде, включая таблицы) ИЛИ перекраска `old_color` → `color` (все выделения одного цвета, напр. жёлтые → зелёные) |
| **`docx_number_formulas(start?)`** | Номера (N) справа у display-формул через SEQ-поля; закладки eq_N |
| **`docx_caption(caption, kind, table?/paragraph?)`** | «Таблица N — …» над таблицей / «Рисунок N — …» под рисунком; SEQ-поле + закладка tbl_N/fig_N |
| **`docx_crossref(anchor, ref_kind?, paragraph?/after_text?)`** | Перекрёстная ссылка через REF (`number`) / PAGEREF (`page`); номера не «едут» |
| **`docx_docproperty(props)`** | DOCPROPERTY-поля (Designation/Customer/Object) — читает нормоконтроль |
| **`docx_replace_font(new_font, old_font?, size?)`** | Массовая смена шрифта по всем run'ам |

## Главные тулы (полный список — в схеме)

| Задача | Тул |
|---|---|
| Прочитать документ | `docx_read` (структура/куски), `get_document_text`, `get_document_outline`, `find_text_in_document` |
| Прочитать НЕ-docx | `file_read` (PDF/XLSX/PPTX/… → Markdown) |
| Сгенерировать текст | `docx_build` (markdown → docx, pandoc) |
| Заполнить форму | `docx_fill` (шаблон + JSON, docxtpl) |
| Создать пустой | `docx_create` |
| Заменить текст | `search_and_replace` (чисто), `track_replace`/`track_insert`/`track_delete` (запись изменений) |
| Добавить/удалить абзацы, заголовки | `add_paragraph`, `add_heading`, `delete_paragraph`, `insert_line_or_paragraph_near_text`, `insert_numbered_list_near_text` |
| Заменить блок между якорями | `replace_paragraph_block_below_header`, `replace_block_between_manual_anchors` |
| Форматирование | `format_text`, `create_custom_style` |
| Таблицы | `add_table`, `format_table`, `merge_table_cells*`, `set_table_column_width*`, `set_table_cell_padding`, `format_table_cell_text` |
| Колонтитулы/поля страниц | `add_header_footer`, `add_page_numbers`, `set_page_layout`, `add_section_break` |
| Оглавление | `add_table_of_contents` |
| Сноски | `add_footnote_to_document`, `add_endnote_to_document` |
| Комментарии | `add_comment`, `get_all_comments` |
| **Формулы** | **`docx_formula`** — только наш тул: `{"op":"list"}` (список формул как LaTeX), `{"op":"insert","latex":"E=mc^2","after_text":"..." или "after_paragraph":N,"display":true}` (отдельным абзацем) / `display:false` (в строку), `{"op":"replace","index":N,"latex":"..."}` |

## Правила

1. **Чистые правки по умолчанию** (`search_and_replace` и др.);
   запись изменений (`track_*`) — только когда пользователь просит.
2. **Не читай весь документ целиком** без нужды — сначала `docx_read(scope="structure")`,
   затем точечно `paragraphs`/`table`/`file_read(pages=...)`. Это экономит токены.
3. **Генерация**: сплошной текст — `docx_build` (markdown одним вызовом,
   append для длинных); формы — `docx_fill`; НЕ собирай многостраничник
   абзацными тулами (медленно, теряется стилевая целостность).
4. Перед заменой уникальных фрагментов проверь их наличие
   (`find_text_in_document`) — заменяй точные подстроки.
5. Формулы — LaTeX-нотацией через `docx_formula` (в готовом файле) или
   прямо в markdown для `docx_build`, НЕ картинками и НЕ текстом.
   Сначала `list`, потом точечные `insert`/`replace`.
6. Нумерация формул/таблиц/рисунков и перекрёстные ссылки — ТОЛЬКО
   полями (docx_number_formulas / docx_caption / docx_crossref),
   никогда не пиши номера вручную.
7. Подсветку («выдели/пометь») делай тулом `docx_highlight` — маркером,
   а не заливкой и не жирным.
8. `word_live_*`-тулы (COM) на сервере НЕ работают — не вызывай.
9. Файлы `.doc` (старый формат) сервер не редактирует — конвертируй
   в `.docx` до правок (или скажи пользователю).
10. После commit покажи пользователю ссылку и sha256.

## Инфра (не для агента)

- Контейнер `doc-editor` на VM, порт 127.0.0.1:9113, сеть lobechat_default
- Код: https://github.com/VissMUT/doc-editor (private, форк word-mcp-live)
- Файлы: rustfs `doc-editor/<uuid>-<имя>` → https://xais.ru/lobe/doc-editor/...
