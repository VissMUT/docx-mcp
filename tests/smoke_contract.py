"""Регрессионный смоук-тест контрактного набора тулов doc-editor.

Запуск (внутри контейнера doc-editor):
    python3 /tmp/smoke_contract.py

Проверяет на pandoc-генерированном документе (docx_build) именно те тулы,
которые использует контракт субагента в промпте Технаря. Ловит класс багов
типа CT_P.tag (python-docx 1.2.0): инструмент «всегда not found» при живом файле.
"""
import asyncio
import json
import os
import sys
import traceback

sys.path.insert(0, "/app")

FAILS = []
WARN = []


def check(name, ok, detail=""):
    if ok:
        print(f"  OK    {name} {detail}")
    else:
        FAILS.append(name)
        print(f"  FAIL  {name} {detail}")


def run(name, coro_fn):
    try:
        r = asyncio.run(coro_fn())
        return r
    except Exception:
        check(name, False, "EXCEPTION: " + str(traceback.format_exc(limit=2)))
        return None


MD = """# Раздел первый

Вводный текст с опечаткой «расчот».

## Энергетический расчёт резерва электропитания

Старый текст блока один.

* пункт старый

## Обоснование применения радиоканала

Текст после блока.

| Колонка А | Колонка Б |
|---|---|
| яч1 | яч2 |

Формула мощности: $P = U \\cdot I$.
"""


def main():
    from word_document_server.tools import storage_tools as st
    from word_document_server.tools import content_tools as ct
    from word_document_server.tools import tracked_changes_tools as tc
    from word_document_server.tools import extended_document_tools as ext
    from word_document_server.utils.document_utils import (
        replace_block_between_manual_anchors,
    )

    # 1. Генерация pandoc-документа (docx_build)
    r = st.docx_build(MD, "smoke", toc=True, number_sections=True)
    check("docx_build", isinstance(r, dict) and r.get("ok") and r.get("filename"), str(r)[:120])
    fname = r.get("filename") if isinstance(r, dict) else None

    # 2. Чтение (docx_read: structure/paragraphs/table)
    r = st.docx_read(scope="structure", filename=fname)
    check("docx_read structure", isinstance(r, dict) and "sections" in r, str(r)[:80])
    r = st.docx_read(scope="paragraphs", filename=fname, start=0, end=10)
    ok = isinstance(r, dict) and r.get("total", 0) > 0
    check("docx_read paragraphs", ok, f"total={r.get('total') if isinstance(r, dict) else '?'}")

    # 3. РЕГРЕССИЯ CT_P.tag: замена блока по якорям
    r = replace_block_between_manual_anchors(
        fname,
        "Энергетический расчёт резерва электропитания",
        ["Новый абзац блока один.", "Новый абзац блока два."],
        "Обоснование применения радиоканала",
    )
    check("replace_block_between_manual_anchors", "not found" not in r, r[:120])

    # 4. Замена блока под заголовком (второй блоковый тул)
    r2 = run(
        "replace_paragraph_block_below_header",
        lambda: ct.replace_paragraph_block_below_header_tool(
            fname, "Обоснование применения радиоканала", ["Новый текст после заголовка."]
        ),
    )
    check("replace_paragraph_block_below_header", r2 is not None and "not found" not in str(r2), str(r2)[:120])

    # 5. search_and_replace
    r3 = run(
        "search_and_replace",
        lambda: ct.search_and_replace(fname, "расчот", "расчёт"),
    )
    check("search_and_replace", r3 is not None and "Replaced" in str(r3), str(r3)[:80])

    # 6. track_replace + accept
    r4 = run(
        "track_replace",
        lambda: tc.track_replace(fname, "Новый абзац блока один", "Абзац блока один (tracked)"),
    )
    check("track_replace", r4 is not None and "error" not in str(r4).lower(), str(r4)[:80])
    r5 = run("accept_tracked_changes", lambda: tc.accept_tracked_changes(fname))
    check("accept_tracked_changes", r5 is not None and "error" not in str(r5).lower(), str(r5)[:80])

    # 7. docx_formula: insert + list (врезка 1x2)
    r6 = st.docx_formula(
        [{"op": "insert", "latex": "R=\\rho\\frac{L}{S}", "display": True, "after_text": "Новый абзац блока два."}]
    )
    check("docx_formula insert", isinstance(r6, dict) and r6.get("ok"), str(r6)[:120])
    r7 = st.docx_formula([{"op": "list"}])
    ok7 = isinstance(r7, dict) and r7.get("ok") and r7.get("results", [{}])[0].get("op") == "list" and len(r7["results"][0].get("formulas", [])) >= 1
    check("docx_formula list", ok7, str(r7)[:120])

    # 8. docx_number_formulas
    r8 = st.docx_number_formulas(filename=fname)
    check("docx_number_formulas", isinstance(r8, dict) and r8.get("ok"), str(r8)[:120])

    # 9. docx_highlight: поиск + перекраска
    r9 = st.docx_highlight(text="Новый текст после заголовка", color="yellow", filename=fname)
    check("docx_highlight search", isinstance(r9, dict) and r9.get("ok"), str(r9)[:120])
    r10 = st.docx_highlight(color="green", old_color="yellow", filename=fname)
    check("docx_highlight recolor", isinstance(r10, dict) and r10.get("ok"), str(r10)[:120])

    # 10. get_highlighted_text
    r11 = run("get_highlighted_text", lambda: ext.get_highlighted_text_from_document(fname))
    check("get_highlighted_text", r11 is not None, str(r11)[:80])

    # 11. docx_caption + docx_crossref
    r12 = st.docx_caption("Сводная таблица", kind="table", filename=fname, table=0)
    check("docx_caption", isinstance(r12, dict) and r12.get("ok"), str(r12)[:120])
    r13 = st.docx_crossref("table_1", ref_kind="number", filename=fname)
    check("docx_crossref", isinstance(r13, dict) and r13.get("ok"), str(r13)[:120])

    # 12. docx_append_block (новый раздел в конец одним вызовом)
    r14 = st.docx_append_block(
        [
            {"type": "heading", "level": 1, "text": "Приложение. Смоук-раздел"},
            {"type": "p", "text": "Абзац приложения."},
            {"type": "formula", "latex": "E=mc^2", "display": True},
        ],
        filename=fname,
    )
    check("docx_append_block", isinstance(r14, dict) and r14.get("ok"), str(r14)[:120])

    # 13. docx_replace_table
    r15 = st.docx_replace_table(0, "| А | Б |\n|---|---|\n| 1 | 2 |")
    check("docx_replace_table", isinstance(r15, dict) and r15.get("ok"), str(r15)[:120])

    # 14. docx_docproperty + docx_replace_font
    r16 = st.docx_docproperty({"Designation": "TEST-001", "Customer": "Смоук"}, filename=fname)
    check("docx_docproperty", isinstance(r16, dict) and r16.get("ok"), str(r16)[:120])
    r17 = st.docx_replace_font("Arial", old_font="Times New Roman", filename=fname)
    check("docx_replace_font", isinstance(r17, dict) and r17.get("ok"), str(r17)[:120])

    # 15. docx_commit + docx_undo
    r18 = st.docx_commit(filename=fname)
    check("docx_commit", isinstance(r18, dict) and r18.get("ok") and r18.get("url"), str(r18)[:120])
    r19 = st.docx_undo()
    check("docx_undo", isinstance(r19, dict) and r19.get("ok"), str(r19)[:120])

    print()
    if FAILS:
        print(f"SMOKE FAILED: {len(FAILS)} fail(s): {FAILS}")
        sys.exit(1)
    print("SMOKE OK: все проверки пройдены")
    sys.exit(0)


if __name__ == "__main__":
    main()
