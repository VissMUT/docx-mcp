# -*- coding: utf-8 -*-
"""doc-editor: интеграция с хранилищем LobeChat + формулы OMML + снапшоты.

Наши инструменты поверх word-mcp-live:
  docx_open(file)     — вложение чата (file_...), https-ссылка или имя файла
                        из БД LobeChat -> рабочий каталог контейнера;
  docx_commit(name)   — загрузка результата в rustfs -> https-ссылка;
  docx_formula(ops)   — формулы: list / insert (абзац, текст или ячейка
                        таблицы table/row/col) / replace / replace_cell
                        (замена содержимого ячейки на формулу; LaTeX <-> OMML
                        через pandoc, нативные формулы Word);
  docx_undo()         — откат к последнему снапшоту (состояние на docx_open
                        или после последнего docx_commit).
"""

import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
import urllib.request
import uuid
import zipfile
from copy import deepcopy

from docx import Document
from docx.oxml import OxmlElement
from lxml import etree

WORK_DIR = os.environ.get('WORK_DIR', '/tmp/doc-editor')
SNAP_DIR = os.path.join(WORK_DIR, 'snapshots')
S3_HOST = os.environ.get('RUSTFS_HOST', 'lobe-rustfs')
S3_BUCKET = os.environ.get('RUSTFS_BUCKET', 'lobe')
S3_ACCESS_KEY = os.environ.get('RUSTFS_ACCESS_KEY', 'lobeadmin')
S3_SECRET_KEY = os.environ.get('RUSTFS_SECRET_KEY', '')
PUBLIC_BASE = os.environ.get('PUBLIC_BASE_URL', 'https://xais.ru')
S3_KEY_PREFIX = os.environ.get('S3_KEY_PREFIX', 'doc-editor')
MAX_BYTES = 300 * 1024 * 1024

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

_lock = threading.Lock()
_state = {'filename': None}


def _tool_json(fn, *args, **kw):
    try:
        with _lock:
            return json.dumps(fn(*args, **kw), ensure_ascii=False)
    except Exception as e:  # noqa: BLE001
        return json.dumps({'error': '%s: %s' % (type(e).__name__, e)},
                          ensure_ascii=False)


def _safe(name):
    name = os.path.basename(
        re.sub(r'[^A-Za-z0-9._-]', '_', name or 'document.docx'))
    return name or 'document.docx'


def _rustfs_url(key):
    return 'http://%s:9000/%s/%s' % (S3_HOST, S3_BUCKET, key.lstrip('/'))


def _download(url, path):
    req = urllib.request.Request(url, headers={'User-Agent': 'doc-editor/1.0'})
    with urllib.request.urlopen(req, timeout=300) as resp:
        total = 0
        with open(path, 'wb') as f:
            while True:
                chunk = resp.read(1 << 20)
                if not chunk:
                    break
                total += len(chunk)
                if total > MAX_BYTES:
                    raise ValueError('файл слишком большой (лимит 300 МБ)')
                f.write(chunk)


def _find_file(q):
    import psycopg2
    conn = psycopg2.connect(
        host=os.environ.get('PGHOST', 'lobe-postgres'),
        dbname=os.environ.get('PGDATABASE', 'lobechat'),
        user=os.environ.get('PGUSER', 'postgres'),
        password=os.environ.get('PGPASSWORD', ''),
        connect_timeout=10)
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT name, url FROM files"
            " WHERE (is_deleted IS NOT TRUE)"
            "   AND url IS NOT NULL AND url NOT LIKE 'internal://%%'"
            "   AND (name ILIKE %s OR name ILIKE %s)"
            " ORDER BY created_at DESC LIMIT 10",
            (q, '%' + q + '%'))
        rows = cur.fetchall()
    finally:
        conn.close()
    if not rows:
        return []
    exact = [r for r in rows if r[0].strip().lower() == q.lower()]
    return exact + [r for r in rows if r not in exact]


def _find_by_id(fid):
    import psycopg2
    conn = psycopg2.connect(
        host=os.environ.get('PGHOST', 'lobe-postgres'),
        dbname=os.environ.get('PGDATABASE', 'lobechat'),
        user=os.environ.get('PGUSER', 'postgres'),
        password=os.environ.get('PGPASSWORD', ''),
        connect_timeout=10)
    try:
        cur = conn.cursor()
        cur.execute("SELECT name, url FROM files"
                    " WHERE id = %s AND (is_deleted IS NOT TRUE)", (fid,))
        row = cur.fetchone()
    finally:
        conn.close()
    return row


def _upload(local_path, key):
    import boto3
    from botocore.config import Config
    s3 = boto3.client(
        's3', endpoint_url='http://%s:9000' % S3_HOST,
        aws_access_key_id=S3_ACCESS_KEY,
        aws_secret_access_key=S3_SECRET_KEY,
        config=Config(signature_version='s3v4',
                      s3={'addressing_style': 'path'}),
        region_name='us-east-1')
    s3.upload_file(local_path, S3_BUCKET, key)


def _snapshot(filename):
    os.makedirs(SNAP_DIR, exist_ok=True)
    src = os.path.join(WORK_DIR, filename)
    if os.path.exists(src):
        shutil.copy2(src, os.path.join(SNAP_DIR, filename + '.last.zip'))


def _restore_snapshot(filename):
    snap = os.path.join(SNAP_DIR, filename + '.last.zip')
    if not os.path.exists(snap):
        raise ValueError('снапшот не найден — сделай docx_open заново')
    shutil.copy2(snap, os.path.join(WORK_DIR, filename))


def _path_of(filename):
    f = (filename or _state.get('filename') or '').strip()
    if not f:
        raise ValueError('файл не открыт: вызови docx_open или укажи '
                         'filename из ответа docx_open')
    f = os.path.basename(f)
    p = os.path.join(WORK_DIR, f)
    if not os.path.exists(p):
        raise ValueError('файл «%s» не найден в рабочем каталоге' % f)
    return f, p


# --------------------------------------------------------------------------
# Формулы: LaTeX <-> OMML (pandoc)
# --------------------------------------------------------------------------

def _pandoc(args, cwd):
    r = subprocess.run(['pandoc'] + args, cwd=cwd,
                       capture_output=True, text=True, timeout=120)
    if r.returncode != 0:
        raise RuntimeError('pandoc: %s' % r.stderr[:300])
    return r


def _latex_to_omml(latex, display):
    """LaTeX -> lxml-элемент m:oMathPara (display) или m:oMath (inline)."""
    tmp = tempfile.mkdtemp(prefix='doced_')
    try:
        body = ('$$\n%s\n$$' % latex) if display else '$%s$' % latex
        md = os.path.join(tmp, 'eq.md')
        out = os.path.join(tmp, 'eq.docx')
        with open(md, 'w', encoding='utf-8') as f:
            f.write(body + '\n')
        _pandoc(['-f', 'markdown+tex_math_dollars', '-t', 'docx',
                 '-o', out, md], tmp)
        with zipfile.ZipFile(out) as z:
            xml = z.read('word/document.xml')
        root = etree.fromstring(xml)
        tag = '{%s}oMathPara' % M_NS if display else '{%s}oMath' % M_NS
        el = root.find('.//' + tag)
        if el is None:
            raise ValueError('pandoc не сгенерировал формулу (проверь LaTeX)')
        return el
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _omml_to_latex(el, display):
    """m:oMathPara/m:oMath -> LaTeX (строка без ограничителей)."""
    tmp = tempfile.mkdtemp(prefix='doced_')
    try:
        d = Document()
        p = d.add_paragraph()
        p._p.append(deepcopy(el))
        out = os.path.join(tmp, 'eq.docx')
        d.save(out)
        r = _pandoc(['-f', 'docx', '-t', 'latex', out], tmp)
        tex = r.stdout.strip()
        if display:
            m = re.search(r'\\\[(.*?)\\\]', tex, re.S)
        else:
            m = re.search(r'\\\((.*?)\\\)', tex, re.S)
        if not m:
            # pandoc может выдать $...$ в markdown-режиме — fallback
            m = re.search(r'\$\$(.*?)\$\$|\$(.*?)\$', tex, re.S)
        return (m.group(1) if m and m.group(1) is not None
                else (m.group(2) if m else tex)).strip()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _iter_omath(doc):
    """[(type, lxml-el, loc)] формул документа.

    loc — индекс body-абзаца (как в doc.paragraphs / docx_read) или
    dict {'table': T, 'row': R, 'col': C} для формул в ячейках таблиц.
    Индексы согласованы с doc.paragraphs — абзацы таблиц НЕ считаются.
    """
    out = []
    for pidx, p in enumerate(doc.paragraphs):
        for child in p._p:
            if child.tag == '{%s}oMath' % M_NS:
                out.append(('inline', child, pidx))
            elif child.tag == '{%s}oMathPara' % M_NS:
                out.append(('display', child, pidx))
    seen = set()
    for ti, tbl in enumerate(doc.tables):
        for ri, row in enumerate(tbl.rows):
            for ci, cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    if p._p in seen:
                        continue
                    seen.add(p._p)
                    for child in p._p:
                        if child.tag == '{%s}oMath' % M_NS:
                            out.append(('inline', child,
                                        {'table': ti, 'row': ri,
                                         'col': ci}))
                        elif child.tag == '{%s}oMathPara' % M_NS:
                            out.append(('display', child,
                                        {'table': ti, 'row': ri,
                                         'col': ci}))
    return out


def _strip_formula_tail(p, math_el):
    """Убрать из абзаца хвост нумерации display-формулы (элементы ПОСЛЕ
    math_el): tab, «(», «)», поля (SEQ eq и др.), закладки eq_*.

    Обычный текст в хвосте не трогаем. Используется delete (перед
    удалением формулы) и number_formulas (перед перенумерацией)."""
    from docx.oxml.ns import qn
    follow = False
    for child in list(p._p):
        if child is math_el:
            follow = True
            continue
        if not follow:
            continue
        tag = child.tag
        if tag in ('{%s}bookmarkStart' % W_NS, '{%s}bookmarkEnd' % W_NS):
            if (child.get(qn('w:name')) or '').startswith('eq_'):
                child.getparent().remove(child)
            continue
        if tag != '{%s}r' % W_NS:
            continue
        txt = ''.join(t.text or ''
                      for t in child.findall('.//{%s}t' % W_NS))
        has_tab = bool(child.findall('.//{%s}tab' % W_NS))
        has_fld = bool(child.findall('.//{%s}fldChar' % W_NS))
        if has_fld or (has_tab and not txt.strip()) or \
                re.match(r'^[\s\(\)\d]+$', txt.strip() or ''):
            child.getparent().remove(child)


# --------------------------------------------------------------------------
# Врезки формул: таблицы 1x2 без границ (формула по центру | номер справа)
# --------------------------------------------------------------------------

def _text_width_twips(doc):
    """Ширина текстовой области (dxa/twips) по первой секции."""
    sec = doc.sections[0]
    w = sec.page_width - sec.left_margin - sec.right_margin
    return max(1, int(w / 635))


def _clear_cell_runs(cell):
    """Очистить ячейку от всего содержимого, оставив один пустой абзац."""
    target = cell.paragraphs[0]
    for child in list(target._p):
        if child.tag != '{%s}pPr' % W_NS:
            target._p.remove(child)
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)


def _formula_table(om, doc):
    """Собрать w:tbl — врезку формулы 1x2 БЕЗ границ на всю ширину текста.

    Ячейка 0 (~92%): абзац jc=center с математикой (display oMathPara);
    ячейка 1 (~8%): пустой абзац jc=right — сюда ставится «(N)».
    Вертикальное центрирование обеих ячеек."""
    from docx.oxml.ns import qn
    W = _text_width_twips(doc)
    w0 = int(W * 0.92)
    w1 = W - w0
    tbl = OxmlElement('w:tbl')
    tblPr = OxmlElement('w:tblPr')
    tblStyle = OxmlElement('w:tblStyle')
    tblStyle.set(qn('w:val'), 'TableNormal')
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(W))
    tblW.set(qn('w:type'), 'dxa')
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:%s' % edge)
        e.set(qn('w:val'), 'nil')
        e.set(qn('w:sz'), '0')
        e.set(qn('w:space'), '0')
        e.set(qn('w:color'), 'auto')
        borders.append(e)
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    cellmar = OxmlElement('w:tblCellMar')
    for side, w in (('top', '40'), ('left', '115'), ('bottom', '40'),
                    ('right', '115')):
        m = OxmlElement('w:%s' % side)
        m.set(qn('w:w'), w)
        m.set(qn('w:type'), 'dxa')
        cellmar.append(m)
    look = OxmlElement('w:tblLook')
    look.set(qn('w:val'), '0000')
    look.set(qn('w:firstRow'), '0')
    look.set(qn('w:lastRow'), '0')
    look.set(qn('w:firstColumn'), '0')
    look.set(qn('w:lastColumn'), '0')
    look.set(qn('w:noHBand'), '0')
    look.set(qn('w:noVBand'), '0')
    for el in (tblStyle, tblW, borders, layout, cellmar, look):
        tblPr.append(el)
    tbl.append(tblPr)
    grid = OxmlElement('w:tblGrid')
    for w in (w0, w1):
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(w))
        grid.append(gc)
    tbl.append(grid)
    tr = OxmlElement('w:tr')
    for col_w, jc_val, content in ((w0, 'center', deepcopy(om)),
                                   (w1, 'right', None)):
        tc = OxmlElement('w:tc')
        tcPr = OxmlElement('w:tcPr')
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), str(col_w))
        tcW.set(qn('w:type'), 'dxa')
        va = OxmlElement('w:vAlign')
        va.set(qn('w:val'), 'center')
        tcPr.append(tcW)
        tcPr.append(va)
        tc.append(tcPr)
        p = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), jc_val)
        pPr.append(jc)
        p.append(pPr)
        if content is not None:
            p.append(content)
        tc.append(p)
        tr.append(tc)
    tbl.append(tr)
    return tbl


def _is_borderless_formula_table(tbl, col_has_math=True):
    """True, если таблица — врезка формулы: 1 строка x 2 колонки, без
    границ, в ячейке 0 математика (если col_has_math), в ячейке 1 только
    «(N)» или пусто."""
    from docx.oxml.ns import qn
    if len(tbl.rows) != 1 or len(tbl.columns) != 2:
        return False
    borders = tbl._tbl.tblPr.find(qn('w:tblBorders'))
    if borders is not None:
        for edge in borders:
            if edge.get(qn('w:val')) not in (None, 'nil', 'none'):
                return False
    if col_has_math and not tbl.cell(0, 0)._tc.findall(
            './/{%s}oMath' % M_NS) and not tbl.cell(0, 0)._tc.findall(
            './/{%s}oMathPara' % M_NS):
        return False
    t1 = tbl.cell(0, 1).text.strip()
    if t1 and not re.match(r'^\(\d+\)$', t1):
        return False
    return True


def _formulas_to_tables(doc):
    """Конвертация display-формул pandoc (абзацы тела с oMathPara) в
    таблицы-врезки 1x2 на том же месте."""
    for p in list(doc.paragraphs):
        omp = p._p.find('{%s}oMathPara' % M_NS)
        if omp is None:
            continue
        tbl = _formula_table(omp, doc)
        p._p.addnext(tbl)
        p._p.getparent().remove(p._p)


# --------------------------------------------------------------------------
# Инструменты
# --------------------------------------------------------------------------

def _resolve_to_workdir(file_ref, force_ext=None):
    """file_xxx / https-URL / имя из БД -> (fname, path, src_name).

    Скачивает файл в WORK_DIR; force_ext — принудительное расширение
    (для docx_open: .docx)."""
    s = (file_ref or '').strip()
    if not s:
        raise ValueError('передай id вложения (file_...), https-ссылку '
                         'или имя файла из БЗ')
    os.makedirs(WORK_DIR, exist_ok=True)
    src_name = None
    if s.startswith('file_'):
        row = _find_by_id(s)
        if not row:
            raise ValueError('файл «%s» не найден в хранилище LobeChat' % s)
        src_name, key = row
        url = _rustfs_url(key)
    elif re.match(r'^https?://', s):
        url = s
    else:
        rows = _find_file(s)
        if not rows:
            raise ValueError('файл «%s» не найден в хранилище (проверь имя '
                             'вложения)' % s)
        src_name, key = rows[0]
        url = _rustfs_url(key)
    fname = _safe(src_name or url.rsplit('/', 1)[-1])
    if force_ext and not fname.lower().endswith(force_ext):
        fname += force_ext
    path = os.path.join(WORK_DIR, fname)
    _download(url, path)
    return fname, path, (src_name or os.path.basename(url))


def docx_open(file_ref):
    """Скачать документ в рабочий каталог и вернуть filename для тулов."""
    fname, path, src_name = _resolve_to_workdir(file_ref,
                                                force_ext=('.docx',))
    doc = Document(path)  # проверка открываемости
    formulas = _iter_omath(doc)
    _snapshot(fname)
    _state['filename'] = fname
    return {
        'ok': True,
        'filename': path,          # этот путь передавай в остальные тулы
        'source_name': src_name,
        'info': {
            'paragraphs': len(doc.paragraphs),
            'tables': len(doc.tables),
            'sections': len(doc.sections),
            'images': len(doc.inline_shapes),
            'formulas': len(formulas),
        },
        'hint': 'все правки вноси тулами (search_and_replace, add_heading, '
                'track_replace, add_header_footer, docx_formula...), затем '
                'вызови docx_commit для получения ссылки.',
    }


def docx_commit(filename=None):
    """Сохранить результат в хранилище и вернуть https-ссылку."""
    fname, path = _path_of(filename)
    Document(path)  # валидация
    key = '%s/%s-%s' % (S3_KEY_PREFIX, uuid.uuid4().hex[:8], fname)
    _upload(path, key)
    with open(path, 'rb') as f:
        sha = hashlib.sha256(f.read()).hexdigest()
    _snapshot(fname)
    return {
        'ok': True,
        'url': '%s/lobe/%s' % (PUBLIC_BASE, key),
        'file': fname,
        'sha256': sha,
        'size': os.path.getsize(path),
    }


def docx_undo():
    """Откатить к последнему снапшоту (на docx_open / последний commit)."""
    fname, path = _path_of(None)
    _restore_snapshot(fname)
    return {'ok': True, 'file': fname,
            'note': 'файл возвращён к последнему снапшоту'}


def docx_replace_table(table, markdown):
    """Снести таблицу целиком и собрать новую из markdown (pipe-таблица)."""
    fname, path = _path_of(None)
    doc = Document(path)
    ti = int(table)
    if ti < 0 or ti >= len(doc.tables):
        raise ValueError('table %d вне диапазона (всего таблиц %d)'
                         % (ti, len(doc.tables)))
    md = (markdown or '').strip()
    if not md:
        raise ValueError('markdown пуст — передай pipe-таблицу')
    tmp = tempfile.mkdtemp(prefix='doced_')
    try:
        mdf = os.path.join(tmp, 't.md')
        out = os.path.join(tmp, 't.docx')
        with open(mdf, 'w', encoding='utf-8') as f:
            f.write(md + '\n')
        _pandoc(['-f', 'markdown', '-t', 'docx', '-o', out, mdf], tmp)
        with zipfile.ZipFile(out) as z:
            xml = z.read('word/document.xml')
        root = etree.fromstring(xml)
        new_tbl = root.find('.//{%s}tbl' % W_NS)
        if new_tbl is None:
            raise ValueError('pandoc не сгенерировал таблицу из markdown '
                             '(проверь синтаксис pipe-таблицы)')
        old_tbl = doc.tables[ti]._tbl
        old_tbl.getparent().replace(old_tbl, new_tbl)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    doc.save(path)
    return {'ok': True, 'table': ti,
            'note': 'таблица заменена (оформление стандартное, pandoc)'}


def docx_formula(ops):
    """Формулы: list / insert (абзац или ячейка таблицы) / replace /
    replace_cell (замена содержимого ячейки таблицы на формулу) /
    replace_text (абзац-маркер find → формула на его месте) /
    delete (удалить формулу по index из list) / clear_cell (очистить ячейку)."""
    if not isinstance(ops, list):
        ops = [ops]
    fname, path = _path_of(None)
    results = []
    doc = Document(path)

    for op in ops:
        kind = (op or {}).get('op')
        if kind == 'list':
            items = []
            for t, el, loc in _iter_omath(doc):
                items.append({'index': len(items), 'type': t,
                              'paragraph': loc,
                              'latex': _omml_to_latex(el, t == 'display')})
            results.append({'op': 'list', 'formulas': items})
            continue

        if kind == 'insert':
            latex = op.get('latex', '')
            display = bool(op.get('display', True))
            om = _latex_to_omml(latex, display)
            target_idx = None
            if op.get('after_paragraph') is not None:
                target_idx = int(op['after_paragraph'])
                if target_idx < 0 or target_idx >= len(doc.paragraphs):
                    raise ValueError('after_paragraph %d вне диапазона '
                                     '(всего %d)'
                                     % (target_idx, len(doc.paragraphs)))
                target = doc.paragraphs[target_idx]
            elif op.get('after_text'):
                t = op['after_text']
                target = None
                for i, p in enumerate(doc.paragraphs):
                    if t in p.text:
                        target, target_idx = p, i
                if target is None:
                    raise ValueError('текст «%s» не найден' % t)
            elif op.get('table') is not None:
                ti = int(op['table'])
                if ti < 0 or ti >= len(doc.tables):
                    raise ValueError('table %d вне диапазона (всего %d)'
                                     % (ti, len(doc.tables)))
                ri = int(op.get('row', 0))
                ci = int(op.get('col', 0))
                try:
                    cell = doc.tables[ti].cell(ri, ci)
                except (IndexError, KeyError) as e:
                    raise ValueError('row/col вне диапазона таблицы %d: %s'
                                     % (ti, e))
                if display:
                    para = cell.paragraphs[-1]
                    if para.text.strip() or len(list(para._p)) > 1:
                        para = cell.add_paragraph()
                    para._p.append(deepcopy(om))
                else:
                    para = cell.paragraphs[-1]
                    para._p.append(deepcopy(om))
                results.append({'op': 'insert', 'ok': True, 'mode': 'cell',
                                'table': ti, 'row': ri, 'col': ci})
                continue
            else:
                target_idx = len(doc.paragraphs) - 1
                target = doc.paragraphs[target_idx]
            if display:
                tbl = _formula_table(om, doc)
                target._p.addnext(tbl)
                results.append({'op': 'insert', 'ok': True,
                                'after_paragraph': target_idx,
                                'mode': 'table',
                                'note': 'формула во врезке 1x2 без границ '
                                        '(по центру); номер — '
                                        'docx_number_formulas'})
            else:
                new_p = OxmlElement('w:p')
                new_p.append(deepcopy(om))
                target._p.addnext(new_p)
                results.append({'op': 'insert', 'ok': True,
                                'after_paragraph': target_idx,
                                'mode': 'paragraph'})
            continue

        if kind == 'replace':
            idx = int(op.get('index', 0))
            latex = op.get('latex', '')
            items = _iter_omath(doc)
            if idx < 0 or idx >= len(items):
                raise ValueError('index %d вне диапазона (всего формул %d)'
                                 % (idx, len(items)))
            t, old_el, pidx = items[idx]
            om = _latex_to_omml(latex, t == 'display')
            old_el.getparent().replace(old_el, deepcopy(om))
            results.append({'op': 'replace', 'index': idx, 'ok': True})
            continue

        if kind == 'replace_cell':
            latex = op.get('latex', '')
            display = bool(op.get('display', True))
            om = _latex_to_omml(latex, display)
            if op.get('table') is not None:
                ti = int(op['table'])
                ri = int(op.get('row', 0))
                ci = int(op.get('col', 0))
            elif op.get('find_text'):
                txt = op['find_text']
                found = None
                for t_i, tbl in enumerate(doc.tables):
                    for r_i, row in enumerate(tbl.rows):
                        for c_i, cell in enumerate(row.cells):
                            if cell.text.strip() == txt:
                                found = (t_i, r_i, c_i)
                                break
                        if found:
                            break
                    if found:
                        break
                if not found:
                    raise ValueError('ячейка с текстом «%s» не найдена' % txt)
                ti, ri, ci = found
            else:
                raise ValueError('replace_cell: укажи table/row/col или '
                                 'find_text')
            if ti < 0 or ti >= len(doc.tables):
                raise ValueError('table %d вне диапазона (всего %d)'
                                 % (ti, len(doc.tables)))
            try:
                cell = doc.tables[ti].cell(ri, ci)
            except (IndexError, KeyError) as e:
                raise ValueError('row/col вне диапазона таблицы %d: %s'
                                 % (ti, e))
            target = cell.paragraphs[0]
            for child in list(target._p):
                if child.tag != '{%s}pPr' % W_NS:
                    target._p.remove(child)
            for extra in cell.paragraphs[1:]:
                extra._p.getparent().remove(extra._p)
            target._p.append(deepcopy(om))
            results.append({'op': 'replace_cell', 'ok': True,
                            'table': ti, 'row': ri, 'col': ci})
            continue

        if kind == 'replace_text':
            # заменить АБЗАЦ-МАРКЕР (find = его точный текст) на формулу.
            # display=true — абзац заменяется врезкой 1x2 на его месте;
            # false — математика встаёт в сам абзац (inline).
            latex = op.get('latex', '')
            display = bool(op.get('display', True))
            find = (op.get('find') or '').strip()
            if not find:
                raise ValueError('replace_text: укажи find (текст абзаца-'
                                 'маркера)')
            om = _latex_to_omml(latex, display)
            target = None
            for p in doc.paragraphs:
                if p.text.strip() == find:
                    target = p
                    break
            if target is None:
                for p in doc.paragraphs:
                    if find in p.text:
                        target = p
                        break
            if target is None:
                raise ValueError('абзац с текстом «%s» не найден' % find)
            if display:
                tbl = _formula_table(om, doc)
                target._p.addprevious(tbl)
                target._p.getparent().remove(target._p)
                mode = 'table'
            else:
                for child in list(target._p):
                    if child.tag != '{%s}pPr' % W_NS:
                        target._p.remove(child)
                target._p.append(deepcopy(om))
                mode = 'inline'
            results.append({'op': 'replace_text', 'find': find, 'ok': True,
                            'mode': mode})
            continue

        if kind == 'delete':
            from docx.text.paragraph import Paragraph
            idx = int(op.get('index', 0))
            items = _iter_omath(doc)
            if idx < 0 or idx >= len(items):
                raise ValueError('index %d вне диапазона (всего формул %d)'
                                 % (idx, len(items)))
            t, old_el, loc = items[idx]
            parent_p = old_el.getparent()
            res = {'op': 'delete', 'index': idx, 'ok': True}
            if isinstance(loc, int):
                _strip_formula_tail(Paragraph(parent_p, doc), old_el)
                old_el.getparent().remove(old_el)
                p = Paragraph(parent_p, doc)
                if t == 'display' and p.text.strip() == '' and \
                        not p._p.findall('.//{%s}drawing' % W_NS):
                    parent_p.getparent().remove(parent_p)
                    res['paragraph_removed'] = loc
                else:
                    res['removed_from_paragraph'] = loc
            else:
                old_el.getparent().remove(old_el)
                tbl = doc.tables[loc['table']]
                if t == 'display' and \
                        _is_borderless_formula_table(tbl, col_has_math=False):
                    tbl._tbl.getparent().remove(tbl._tbl)
                    res['formula_table_removed'] = loc['table']
                else:
                    res['removed_from_cell'] = loc
            results.append(res)
            continue

        if kind == 'clear_cell':
            if op.get('table') is not None:
                ti = int(op['table'])
                ri = int(op.get('row', 0))
                ci = int(op.get('col', 0))
            elif op.get('find_text'):
                txt = op['find_text']
                found = None
                for t_i, tbl in enumerate(doc.tables):
                    for r_i, row in enumerate(tbl.rows):
                        for c_i, cell in enumerate(row.cells):
                            if cell.text.strip() == txt:
                                found = (t_i, r_i, c_i)
                                break
                        if found:
                            break
                    if found:
                        break
                if not found:
                    raise ValueError('ячейка с текстом «%s» не найдена' % txt)
                ti, ri, ci = found
            else:
                raise ValueError('clear_cell: укажи table/row/col или '
                                 'find_text')
            if ti < 0 or ti >= len(doc.tables):
                raise ValueError('table %d вне диапазона (всего %d)'
                                 % (ti, len(doc.tables)))
            try:
                cell = doc.tables[ti].cell(ri, ci)
            except (IndexError, KeyError) as e:
                raise ValueError('row/col вне диапазона таблицы %d: %s'
                                  % (ti, e))
            _clear_cell_runs(cell)
            results.append({'op': 'clear_cell', 'ok': True,
                            'table': ti, 'row': ri, 'col': ci})
            continue

        raise ValueError('неизвестная операция формулы: %r (list/insert/'
                         'replace/replace_cell/delete/clear_cell)' % kind)

    doc.save(path)
    return {'ok': True, 'results': results}


def _heading_style(doc, level):
    """Стиль заголовка уровня level (1..4) — по styleId или имени.
    В pandoc-доках w:name в нижнем регистре, add_heading() по имени
    не находит — ищем по styleId 'Heading%d'."""
    sid = 'Heading%d' % level
    for st in doc.styles:
        try:
            if st.style_id == sid:
                return st
        except Exception:  # noqa: BLE001
            pass
    for st in doc.styles:
        if (st.name or '') == 'Heading %d' % level:
            return st
    raise ValueError('стиль заголовка уровня %d не найден в документе'
                     % level)


def docx_append_block(items, filename=None):
    """Добавить блок контента (заголовки, абзацы, формулы) В КОНЕЦ
    документа одним вызовом. Порядок гарантирован, существующее
    содержимое не трогается.

    items: список элементов, добавляются последовательно:
      {"type":"heading", "level":1..4, "text":"..."} — заголовок;
      {"type":"p", "text":"..."} — обычный абзац;
      {"type":"formula", "latex":"...", "display":true} — формула
        (display=true — отдельным абзацем ПО ЦЕНТРУ; false — в конец
        предыдущего добавленного абзаца).

    Для новых разделов с формулами используй ЭТОТ тул одним вызовом
    вместо цепочки add_heading/add_paragraph/docx_formula. Номера формул
    потом: docx_number_formulas()."""
    if not isinstance(items, list) or not items:
        raise ValueError('items: непустой список элементов')
    fname, path = _path_of(filename)
    doc = Document(path)
    added_p = 0
    added_f = 0
    last_p = None
    for item in items:
        if not isinstance(item, dict):
            raise ValueError('каждый элемент — dict {type, ...}')
        kind = item.get('type')
        if kind == 'heading':
            lvl = int(item.get('level', 1))
            if lvl < 1 or lvl > 4:
                raise ValueError('level %d вне 1..4' % lvl)
            p = doc.add_paragraph(item.get('text', ''))
            p.style = _heading_style(doc, lvl)
            added_p += 1
            last_p = p
        elif kind == 'p':
            p = doc.add_paragraph(item.get('text', ''))
            added_p += 1
            last_p = p
        elif kind == 'formula':
            latex = item.get('latex', '')
            display = bool(item.get('display', True))
            om = _latex_to_omml(latex, display)
            if display:
                tbl = _formula_table(om, doc)
                body = doc.element.body
                sect = body.find('{%s}sectPr' % W_NS)
                if sect is not None:
                    sect.addprevious(tbl)
                else:
                    body.append(tbl)
                added_p += 1
                last_p = None
            else:
                if last_p is None:
                    p = doc.add_paragraph()
                    added_p += 1
                    last_p = p
                last_p._p.append(deepcopy(om))
            added_f += 1
        else:
            raise ValueError('неизвестный type %r (heading/p/formula)'
                             % kind)
    doc.save(path)
    return {'ok': True, 'paragraphs_added': added_p,
            'formulas_added': added_f,
            'note': 'добавлено в конец документа; display-формулы — во '
                    'врезках 1x2 по центру; для номеров вызови '
                    'docx_number_formulas'} 


# --------------------------------------------------------------------------
# Чтение: docx_read (структура с индексами python-docx) и file_read (anymd)
# --------------------------------------------------------------------------

RESPONSE_CAP = 60000
EMU_PER_CM = 360000.0


def _clip_size(obj):
    """Гарантия размера ответа: при превышении — подсказка сузить scope."""
    s = json.dumps(obj, ensure_ascii=False)
    if len(s) <= RESPONSE_CAP:
        return obj
    if isinstance(obj, dict) and 'paragraphs' in obj:
        while len(json.dumps(obj, ensure_ascii=False)) > RESPONSE_CAP:
            obj['paragraphs'] = obj['paragraphs'][:max(1, len(obj['paragraphs']) // 2)]
            obj['truncated'] = True
        return obj
    return {'error': 'ответ слишком большой (%d б) — сузь диапазон '
                     '(параметры from/to, pages, max_tokens)' % len(s)}


def _heading_tree(doc):
    tree, stack = [], []
    for i, p in enumerate(doc.paragraphs):
        st = (p.style.name or '') if p.style else ''
        if st.startswith('Heading'):
            try:
                lvl = int(st.rsplit(None, 1)[-1])
            except (ValueError, IndexError):
                lvl = 2
            node = {'level': lvl, 'text': p.text[:120], 'paragraph': i,
                    'children': []}
            while stack and stack[-1]['level'] >= lvl:
                stack.pop()
            if stack:
                stack[-1]['children'].append(node)
            else:
                tree.append(node)
            stack.append(node)
    return tree


def _fields_scan(doc):
    """Поля (PAGE/NUMPAGES/DATE/SEQ/...) с индексами абзацев."""
    body = doc.element.body
    out, pidx = [], 0
    for p in body.iter('{%s}p' % W_NS):
        instr, active = '', False
        for node in p.iter():
            tag = etree.QName(node).localname
            if tag == 'fldChar':
                t = node.get('{%s}fldCharType' % W_NS)
                if t == 'begin':
                    active, instr = True, ''
                elif t == 'end' and active:
                    m = re.match(r'\s*([A-Z]+)', instr or '')
                    out.append({'field': m.group(1) if m else (instr or '?')[:30],
                                'paragraph': pidx})
                    active = False
            elif active and tag == 'instrText':
                instr += node.text or ''
            elif tag == 'fldSimple':
                instr_s = node.get('{%s}instr' % W_NS) or ''
                m = re.match(r'\s*([A-Z]+)', instr_s)
                out.append({'field': m.group(1) if m else instr_s[:30],
                            'paragraph': pidx})
        pidx += 1
    return out


def _docx_read_structure(doc, fname):
    return {
        'file': fname,
        'sections': [{
            'index': i,
            'size_mm': [round(s.page_width / EMU_PER_CM * 10, 1),
                        round(s.page_height / EMU_PER_CM * 10, 1)],
            'orientation': 'landscape'
                           if s.page_width > s.page_height else 'portrait',
            'margins_mm': {
                'left': round(s.left_margin / EMU_PER_CM * 10, 1),
                'right': round(s.right_margin / EMU_PER_CM * 10, 1),
                'top': round(s.top_margin / EMU_PER_CM * 10, 1),
                'bottom': round(s.bottom_margin / EMU_PER_CM * 10, 1)},
        } for i, s in enumerate(doc.sections)],
        'headings': _heading_tree(doc),
        'counts': {
            'paragraphs': len(doc.paragraphs),
            'tables': len(doc.tables),
            'images': len(doc.inline_shapes),
            'formulas': len(_iter_omath(doc)),
            'fields': len(_fields_scan(doc)),
            'comments': _read_comments_part(doc),
        },
        'tables': [{
            'index': i, 'rows': len(t.rows), 'cols': len(t.columns),
            'first_row': [c.text[:40] for c in t.rows[0].cells[:6]],
        } for i, t in enumerate(doc.tables)],
    }


def _read_comments_part(doc):
    """Число комментариев (word/comments.xml), без разбора содержимого."""
    try:
        for rel in doc.part.rels.values():
            if rel.reltype.endswith('/comments'):
                xml = rel.target_part.blob
                return len(etree.fromstring(xml).findall(
                    './/{%s}comment' % W_NS))
    except Exception:  # noqa: BLE001
        pass
    return 0


def docx_read(scope='structure', filename=None, start=None, end=None,
              table=None):
    """Структурное чтение открытого документа (индексы — как в python-docx).

    scope:
      structure       — карта: секции (формат/поля), дерево заголовков,
                        счётчики, таблицы (индекс/размер/первая строка);
      paragraphs      — абзацы start..end (по умолч. 0..199): index/style/text;
      table           — сетка таблицы с индексом table;
      headers_footers — колонтитулы секций (default/first/even);
      fields          — поля PAGE/NUMPAGES/DATE/SEQ с абзацами;
      images          — картинки с размерами и абзацами.
    """
    fname, path = _path_of(filename)
    doc = Document(path)
    if scope == 'structure':
        return _docx_read_structure(doc, fname)
    if scope == 'paragraphs':
        total = len(doc.paragraphs)
        a = max(0, int(start if start is not None else 0))
        b = min(total, int(end if end is not None else a + 200))
        return _clip_size({
            'total': total, 'from': a, 'to': b, 'truncated': False,
            'paragraphs': [{'index': i,
                            'style': (p.style.name if p.style else None),
                            'text': p.text[:500]}
                           for i, p in enumerate(doc.paragraphs[a:b], start=a)]})
    if scope == 'table':
        idx = int(table if table is not None else 0)
        if idx < 0 or idx >= len(doc.tables):
            raise ValueError('таблица %d не найдена (всего %d)'
                             % (idx, len(doc.tables)))
        t = doc.tables[idx]
        return {
            'index': idx, 'rows': len(t.rows), 'cols': len(t.columns),
            'grid': [[c.text[:200] for c in row.cells] for row in t.rows]}
    if scope == 'headers_footers':
        out = []
        for i, s in enumerate(doc.sections):
            item = {'section': i}
            for name, part in (('header', s.header),
                               ('first_header', s.first_page_header),
                               ('even_header', s.even_page_header),
                               ('footer', s.footer),
                               ('first_footer', s.first_page_footer),
                               ('even_footer', s.even_page_footer)):
                try:
                    item[name] = ('\n'.join(p.text for p in part.paragraphs)
                                  or None)
                except Exception:  # noqa: BLE001
                    item[name] = None
            out.append(item)
        return {'headers_footers': out}
    if scope == 'fields':
        return {'fields': _fields_scan(doc)}
    if scope == 'images':
        return {'images': [{
            'index': i,
            'size_cm': [round(sh.width / EMU_PER_CM, 2),
                        round(sh.height / EMU_PER_CM, 2)],
        } for i, sh in enumerate(doc.inline_shapes)]}
    raise ValueError('неизвестный scope: %r (structure/paragraphs/table/'
                     'headers_footers/fields/images)' % scope)


ANYMD_BIN = os.environ.get('ANYMD_BIN', 'anymd')


def file_read(source, pages=None, max_tokens=12000, scope=None):
    """Прочитать ЛЮБОЙ файл как Markdown (anymd): PDF/DOCX/XLSX/PPTX/EPUB/
    HTML/изображения.

    source: id вложения (file_...), https-ссылка или имя файла из БЗ.
    pages: диапазон страниц PDF, напр. "1-5,8" (для многостраничных).
    max_tokens: бюджет (по умолч. 12000) — для больших файлов указывай
    меньше или используй pages.

    Возвращает {file, markdown, truncated, chars}. Таблицы — markdown-
    таблицами, формулы — LaTeX, страницы PDF — маркерами <!-- page N -->.
    Не влияет на файл, открытый через docx_open.
    """
    fname, path, src_name = _resolve_to_workdir(source)
    cmd = [ANYMD_BIN, path]
    if pages:
        cmd += ['--pages', str(pages)]
    if max_tokens:
        cmd += ['--max-tokens', str(int(max_tokens))]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    except subprocess.TimeoutExpired:
        raise RuntimeError('anymd не уложился в 120с — уменьши pages')
    if r.returncode != 0:
        raise RuntimeError('anymd: %s' % (r.stderr or '')[:300])
    cap = 200000
    out = r.stdout or ''
    return {'ok': True, 'file': fname, 'markdown': out[:cap],
            'truncated': len(out) > cap, 'chars': len(out),
            'note': ('продолжение большого документа — повторный вызов с '
                     'параметром pages (следующий диапазон)')}


# --------------------------------------------------------------------------
# Генерация: docx_create, docx_build (pandoc), docx_fill (docxtpl)
# --------------------------------------------------------------------------

MD_CAP = 200 * 1024  # лимит markdown-входа docx_build
REFERENCE_DOCX = os.environ.get('REFERENCE_DOCX',
                                '/assets/reference.docx')
_build_buffers = {}  # fname -> накопленный markdown (append-сборка)


def _ensure_docx_name(name):
    fname = _safe(name or 'document')
    if not fname.lower().endswith('.docx'):
        fname += '.docx'
    return fname


def _gen_contract(fname, path, hint, **info):
    """Единый контракт генераторов: как docx_open, но «уже открыт»."""
    _snapshot(fname)
    _state['filename'] = fname
    return {'ok': True, 'filename': path, 'info': info, 'hint': hint}


def _apply_base_styles(doc):
    """Базовые стили: TNR 12 justify отступ 1.25, заголовки 13-14 bold,
    A4 20/10/20/10, PAGE в футере (правила Технаря)."""
    from docx.shared import Pt, Cm, Mm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    def set_font(style, size, bold=None, align=None, indent=None):
        style.font.name = 'Times New Roman'
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor(0, 0, 0)
        rpr = style.element.get_or_add_rPr()
        rf = rpr.get_or_add_rFonts()
        for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
            rf.set(qn(attr), 'Times New Roman')
        if bold is not None:
            style.font.bold = bold
        if align is not None:
            style.paragraph_format.alignment = align
        if indent is not None:
            style.paragraph_format.first_line_indent = Cm(indent)

    set_font(doc.styles['Normal'], 12, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             indent=1.25)
    doc.styles['Normal'].paragraph_format.space_after = Pt(0)
    set_font(doc.styles['Heading 1'], 14, bold=True,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=0)
    set_font(doc.styles['Heading 2'], 13, bold=True,
             align=WD_ALIGN_PARAGRAPH.LEFT, indent=0)

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(1.0)
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)

    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve'); instr.text = ' PAGE '
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
    for el in (fld1, instr, fld2):
        r._r.append(el)


def docx_create(name, title=None, author=None):
    """Создать НОВЫЙ документ DOCX в рабочем каталоге.

    Стили: Times New Roman 12, выравнивание по ширине, отступ первой
    строки 1.25 см, заголовки 13-14 pt bold, A4, поля 20/10/20/10 мм,
    номер страницы в нижнем колонтитуле.
    """
    fname = _ensure_docx_name(name)
    os.makedirs(WORK_DIR, exist_ok=True)
    path = os.path.join(WORK_DIR, fname)
    doc = Document()
    if title:
        doc.core_properties.title = title
    if author:
        doc.core_properties.author = author
    _apply_base_styles(doc)
    doc.save(path)
    return _gen_contract(
        fname, path,
        hint=('документ создан и уже открыт: наполняй тулами '
              '(add_heading, add_paragraph, add_table, docx_formula...), '
              'затем docx_commit() для ссылки'),
        paragraphs=len(doc.paragraphs), tables=len(doc.tables),
        sections=len(doc.sections))


_PAGE_BREAK_BLOCK = ('```{=openxml}\n<w:p><w:r><w:br w:type="page"/>'
                     '</w:r></w:p>\n```\n')


def _normalize_markdown(md):
    """\\newpage/\\pagebreak -> raw-openxml разрыв страницы."""
    return re.sub(r'\\newpage\b|\\pagebreak\b', _PAGE_BREAK_BLOCK, md)


def _polish_tables(doc):
    """Границы всем таблицам + чистка ячеек (убрать отступы и табы)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    for tbl in doc.tables:
        tblPr = tbl._tbl.tblPr
        old = tblPr.find(qn('w:tblBorders'))
        if old is not None:
            tblPr.remove(old)
        borders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            e = OxmlElement('w:%s' % edge)
            e.set(qn('w:val'), 'single')
            e.set(qn('w:sz'), '4')
            e.set(qn('w:space'), '0')
            e.set(qn('w:color'), 'auto')
            borders.append(e)
        tblPr.append(borders)
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    pPr = p._p.get_or_add_pPr()
                    ind = pPr.find(qn('w:ind'))
                    if ind is not None:
                        pPr.remove(ind)
                    pstyle = pPr.find(qn('w:pStyle'))
                    if pstyle is not None:
                        pPr.remove(pstyle)
                    tabs = pPr.find(qn('w:tabs'))
                    if tabs is not None:
                        pPr.remove(tabs)
                    for r in p._p.findall(qn('w:r')):
                        for tab in r.findall(qn('w:tab')):
                            r.remove(tab)


def docx_build(markdown, filename, toc=None, number_sections=None,
               append=None):
    """Собрать DOCX из markdown (pandoc + эталонные стили).

    markdown: текст в pandoc-диалекте: pipe-таблицы, формулы $...$ / $$...$$
    (или \\(...\\) / \\[...\\]), разрыв страницы — \\newpage, картинки —
    ссылками (SVG не поддерживается). Лимит 200 КБ за вызов.
    append: true — доклеить markdown к буферу документа и пересобрать
    заново (для длинных документов частями).
    """
    if not isinstance(markdown, str) or not markdown.strip():
        raise ValueError('markdown пуст')
    if len(markdown) > MD_CAP:
        raise ValueError('markdown больше лимита 200 КБ — разбей на части '
                         'и доклеивай через append=true')
    fname = _ensure_docx_name(filename)
    os.makedirs(WORK_DIR, exist_ok=True)
    path = os.path.join(WORK_DIR, fname)
    if append:
        if fname not in _build_buffers:
            raise ValueError('буфер пуст: первый вызов делай без append=true')
        _build_buffers[fname] = _build_buffers[fname] + '\n\n' + markdown
    else:
        _build_buffers[fname] = markdown
    full_md = _normalize_markdown(_build_buffers[fname])
    tmp = tempfile.mkdtemp(prefix='doced_')
    try:
        md_path = os.path.join(tmp, 'doc.md')
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(full_md)
        args = ['-f', 'markdown+tex_math_dollars+tex_math_single_backslash'
                      '+tex_math_double_backslash+raw_attribute',
                '-t', 'docx']
        if os.path.exists(REFERENCE_DOCX):
            args += ['--reference-doc', REFERENCE_DOCX]
        if toc:
            args.append('--toc')
        if number_sections:
            args.append('--number-sections')
        args += ['-o', path, md_path]
        _pandoc(args, tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    doc = Document(path)
    _polish_tables(doc)
    _formulas_to_tables(doc)
    doc.save(path)
    doc = Document(path)  # валидация
    return _gen_contract(
        fname, path,
        hint=('документ собран и уже открыт: точечные правки тулами '
              '(search_and_replace, add_header_footer...), затем '
              'docx_commit() для ссылки. Продолжение длинного документа: '
              'docx_build(markdown, filename, append=true).'),
        paragraphs=len(doc.paragraphs), tables=len(doc.tables),
        formulas=len(_iter_omath(doc)), sections=len(doc.sections),
        buffer_chars=len(_build_buffers[fname]))


def docx_fill(template_ref, context, filename=None):
    """Заполнить DOCX-шаблон данными (docxtpl/jinja2).

    template_ref: id вложения (file_...), https-ссылка или имя файла из БЗ
    с плейсхолдерами {{var}}, {% for %}, {% if %}.
    context: JSON-объект со значениями.
    Результат пишется в НОВЫЙ файл (шаблон не затирается).
    """
    try:
        from docxtpl import DocxTemplate
        from jinja2.sandbox import SandboxedEnvironment
    except ImportError:
        raise RuntimeError('docxtpl не установлен в образе')
    if not isinstance(context, dict):
        raise ValueError('context должен быть JSON-объектом (словарём)')
    tname, tpath, _ = _resolve_to_workdir(template_ref,
                                          force_ext=('.docx',))
    base = os.path.splitext(tname)[0]
    fname = _ensure_docx_name(filename or (base + '_filled'))
    if fname == tname:
        fname = _ensure_docx_name(base + '_filled')
    path = os.path.join(WORK_DIR, fname)
    tpl = DocxTemplate(tpath)
    tpl.render(context, jinja_env=SandboxedEnvironment())
    tpl.save(path)
    undeclared = []
    try:
        undeclared = list(tpl.undeclared_template_variables)
    except Exception:  # noqa: BLE001
        pass
    doc = Document(path)  # валидация
    return _gen_contract(
        fname, path,
        hint=('шаблон заполнен и результат уже открыт: правки тулами, '
              'затем docx_commit() для ссылки'),
        paragraphs=len(doc.paragraphs), tables=len(doc.tables),
        undeclared_variables=undeclared)


# --------------------------------------------------------------------------
# СПДС-утилиты: подсветка, перекрёстные ссылки, подписи, нумерация формул,
# DOCPROPERTY, шрифты
# --------------------------------------------------------------------------

HL_COLORS = ('yellow', 'green', 'cyan', 'magenta', 'blue', 'red',
             'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta',
             'darkRed', 'darkYellow', 'darkGray', 'lightGray', 'black',
             'none')

_bookmark_seq = [1000]


def _next_bm_id():
    _bookmark_seq[0] += 1
    return _bookmark_seq[0]


def _field_run(paragraph, instr, bookmark=None, text=None):
    """Вставить поле Word (SEQ/REF/PAGEREF/DOCPROPERTY) в абзац.

    Каноническая структура OOXML в отдельных runs:
    begin / instrText / separate / [bookmarkStart] кэш [bookmarkEnd] / end.
    Закладка — внутри результата (между separate и end), чтобы Word
    не терял её при обновлении полей (F9) и не выдавал
    «Ошибка! Закладка не определена».
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    r_begin = paragraph.add_run()
    f_begin = OxmlElement('w:fldChar')
    f_begin.set(qn('w:fldCharType'), 'begin')
    r_begin._r.append(f_begin)
    r_instr = paragraph.add_run()
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = instr
    r_instr._r.append(it)
    r_sep = paragraph.add_run()
    f_sep = OxmlElement('w:fldChar')
    f_sep.set(qn('w:fldCharType'), 'separate')
    r_sep._r.append(f_sep)
    if bookmark:
        bm_s = OxmlElement('w:bookmarkStart')
        bm_s.set(qn('w:id'), str(_next_bm_id()))
        bm_s.set(qn('w:name'), bookmark)
        r_sep._r.addnext(bm_s)
    r_text = None
    if text:
        r_text = paragraph.add_run()
        t = OxmlElement('w:t')
        t.text = text
        r_text._r.append(t)
    if bookmark:
        bm_e = OxmlElement('w:bookmarkEnd')
        bm_e.set(qn('w:id'), str(_bookmark_seq[0]))
        if r_text is not None:
            r_text._r.addnext(bm_e)
        else:
            r_sep._r.addnext(bm_e)
    r_end = paragraph.add_run()
    f_end = OxmlElement('w:fldChar')
    f_end.set(qn('w:fldCharType'), 'end')
    r_end._r.append(f_end)
    return r_end


def _iter_all_paragraphs(doc):
    """Абзацы тела + ячейки таблиц + колонтитулы."""
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for s in doc.sections:
        for part in (s.header, s.footer, s.first_page_header,
                     s.first_page_footer, s.even_page_header,
                     s.even_page_footer):
            try:
                for p in part.paragraphs:
                    yield p
            except Exception:  # noqa: BLE001
                continue


def docx_highlight(text=None, color='yellow', filename=None, paragraph=None,
                   old_color=None):
    """Подсветить (маркером) текст или перекрасить существующую подсветку.

    Режим 1 (поиск текста): text — подстрока; подсвечиваются run'ы, её
    содержащие. paragraph: индекс абзаца или None — везде, включая таблицы.
    Режим 2 (перекраска): old_color — заменить ВСЕ выделения этого цвета
    на color (text не нужен). Напр. старые жёлтые → зелёные:
    docx_highlight(color='green', old_color='yellow').
    color='none' снимает подсветку.
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    if color not in HL_COLORS:
        raise ValueError('неизвестный цвет %r (варианты: %s)'
                         % (color, ', '.join(HL_COLORS)))
    if old_color is not None and old_color not in HL_COLORS:
        raise ValueError('неизвестный old_color %r (варианты: %s)'
                         % (old_color, ', '.join(HL_COLORS)))
    if text is None and old_color is None:
        raise ValueError('передай text (поиск) или old_color (перекраска)')
    texts = list(text) if isinstance(text, (list, tuple)) else [text]
    fname, path = _path_of(filename)
    doc = Document(path)
    done = 0
    if old_color is not None:
        targets = _iter_all_paragraphs(doc)
        for p in targets:
            for r in p.runs:
                rpr = r._r.get_or_add_rPr()
                hl = rpr.find(qn('w:highlight'))
                if hl is None or hl.get(qn('w:val')) != old_color:
                    continue
                if color == 'none':
                    rpr.remove(hl)
                else:
                    hl.set(qn('w:val'), color)
                done += 1
    else:
        targets = ([doc.paragraphs[int(paragraph)]]
                   if paragraph is not None else _iter_all_paragraphs(doc))
        for p in targets:
            for t in texts:
                if t not in p.text:
                    continue
                for r in p.runs:
                    if t not in r.text:
                        continue
                    rpr = r._r.get_or_add_rPr()
                    hl = rpr.find(qn('w:highlight'))
                    if color == 'none':
                        if hl is not None:
                            rpr.remove(hl)
                            done += 1
                        continue
                    if hl is None:
                        hl = OxmlElement('w:highlight')
                        rpr.append(hl)
                    hl.set(qn('w:val'), color)
                    done += 1
    doc.save(path)
    return {'ok': True, 'highlighted_runs': done,
            'note': 'подсветка видна как маркер в Word'}


def docx_crossref(anchor, ref_kind='number', filename=None,
                  paragraph=None, after_text=None):
    """Вставить перекрёстную ссылку (REF/PAGEREF поле) на закладку.

    anchor: имя закладки (add_bookmark, docx_caption, docx_number_formulas
    ставят их автоматически).
    ref_kind: 'number' (номер, REF \\h) или 'page' (страница, PAGEREF).
    Позиция: paragraph (индекс) или after_text (в конец найденного абзаца).
    """
    if ref_kind not in ('number', 'page'):
        raise ValueError("ref_kind: 'number' или 'page'")
    fname, path = _path_of(filename)
    doc = Document(path)
    if paragraph is not None:
        if int(paragraph) < 0 or int(paragraph) >= len(doc.paragraphs):
            raise ValueError('paragraph %s вне диапазона (всего %d)'
                             % (paragraph, len(doc.paragraphs)))
        p = doc.paragraphs[int(paragraph)]
    elif after_text:
        p = None
        for cand in doc.paragraphs:
            if after_text in cand.text:
                p = cand
                break
        if p is None:
            raise ValueError('текст «%s» не найден' % after_text)
    else:
        p = doc.paragraphs[-1]
    instr = (' PAGEREF %s \\h ' % anchor if ref_kind == 'page'
             else ' REF %s \\h ' % anchor)
    _field_run(p, instr)
    doc.save(path)
    return {'ok': True, 'ref_kind': ref_kind, 'anchor': anchor,
            'note': 'ссылка обновится в Word (F9 / открытие)'}


def docx_caption(caption, kind='table', seq_name=None, filename=None,
                 table=None, paragraph=None):
    """Подпись таблицы (над ней) или рисунка (под ним) с автонумерацией.

    kind: 'table' | 'figure'. caption: текст подписи без номера.
    table: индекс таблицы для kind='table' (подпись вставляется над ней).
    paragraph: индекс абзаца с рисунком для kind='figure' (подпись под ним;
    по умолчанию — последний абзац с картинкой).
    Ставит закладку tbl_N / fig_N для docx_crossref.
    """
    if kind not in ('table', 'figure'):
        raise ValueError("kind: 'table' или 'figure'")
    fname, path = _path_of(filename)
    doc = Document(path)
    seq = seq_name or ('Table' if kind == 'table' else 'Figure')
    label = ('Таблица' if kind == 'table' else 'Рисунок')
    p = doc.add_paragraph()
    anchor = '%s_%d' % (seq.lower(), _next_bm_id())
    _field_run(p, ' SEQ %s \\* ARABIC ' % seq, text='1',
               bookmark=anchor)
    p.add_run(' — %s' % caption)
    try:
        p.style = doc.styles['Caption']
    except KeyError:
        pass
    if kind == 'table':
        idx = int(table) if table is not None else 0
        if idx < 0 or idx >= len(doc.tables):
            raise ValueError('таблица %d не найдена (всего %d)'
                             % (idx, len(doc.tables)))
        doc.tables[idx]._tbl.addprevious(p._p)
    else:
        if paragraph is not None:
            idx = int(paragraph)
        else:
            idx = None
            for i in range(len(doc.paragraphs) - 1, -1, -1):
                body = doc.paragraphs[i]._p
                if body.findall('.//{*}drawing') or body.findall('.//{*}pict'):
                    idx = i
                    break
            if idx is None:
                raise ValueError('рисунок не найден — передай paragraph '
                                 'явно')
        if idx < 0 or idx >= len(doc.paragraphs):
            raise ValueError('paragraph %d вне диапазона (всего %d)'
                             % (idx, len(doc.paragraphs)))
        doc.paragraphs[idx]._p.addnext(p._p)
    doc.save(path)
    return {'ok': True, 'label': '%s N — %s' % (label, caption),
            'seq': seq, 'anchor': anchor,
            'note': 'номер обновится в Word (F9)'}


def docx_number_formulas(filename=None, start=1):
    """Пронумеровать display-формулы ПРОСТЫМ текстом «(N)».

    Номер ставится в ПРАВУЮ ячейку врезки (таблица 1x2): формула в
    ячейке 0 по центру, «(N)» — в ячейке 1 справа. Для формул в рабочих
    таблицах (2 колонки, формула в col 0) номер кладётся в соседнюю
    ячейку. Legacy-абзацы с oMathPara в теле документа мигрируют во
    врезки. Без SEQ-полей и закладок: номера видны сразу, F9 не нужен.
    Перевызов перезаписывает номера (идемпотентно).

    start: стартовый номер (обычно 1).
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    fname, path = _path_of(filename)
    doc = Document(path)
    _formulas_to_tables(doc)
    items = _iter_omath(doc)
    done = 0
    skipped = []
    for t, el, loc in items:
        if t != 'display':
            continue
        n = int(start) + done
        if isinstance(loc, dict):
            tbl = doc.tables[loc['table']]
            if len(tbl.columns) == 2 and loc['col'] == 0:
                num_cell = tbl.cell(loc['row'], 1)
            else:
                skipped.append({'loc': loc, 'reason': 'нет соседней '
                                'ячейки под номер'})
                continue
        else:
            skipped.append({'paragraph': loc,
                            'reason': 'не мигрировало во врезку'})
            continue
        _clear_cell_runs(num_cell)
        p = num_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run('(%d)' % n)
        done += 1
    doc.save(path)
    return {'ok': True, 'numbered': done, 'start': int(start),
            'skipped': skipped,
            'note': 'номера — простой текст «(N)» в правой ячейке врезки; '
                    'при изменении состава формул вызови заново'}


def docx_docproperty(props, filename=None):
    """Управление DOCPROPERTY-полями (обозначение тома, заказчик...).

    props: {key: value} — вставляет поля в конец документа. Спец-ключи:
    'Designation' (обозначение тома), 'Customer' (заказчик), 'Object'
    (объект). Нормоконтроль читает эти поля.
    """
    if not isinstance(props, dict) or not props:
        raise ValueError('props: непустой {key: value}')
    fname, path = _path_of(filename)
    doc = Document(path)
    for k, v in props.items():
        p = doc.add_paragraph()
        p.add_run('%s: ' % k)
        _field_run(p, ' DOCPROPERTY "%s" ' % str(k), text=str(v))
    doc.save(path)
    return {'ok': True, 'props': list(props.keys()),
            'note': 'поля DOCPROPERTY добавлены в конец документа'}


def docx_replace_font(new_font, old_font=None, size=None, filename=None):
    """Массовая замена шрифта по всем run'ам (тело, таблицы, колонтитулы).

    new_font: новый шрифт (напр. 'Times New Roman', 'GOST type A').
    old_font: менять только этот шрифт (None — все).
    size: необязательный кегль в пунктах.
    """
    from docx.shared import Pt
    from docx.oxml.ns import qn
    fname, path = _path_of(filename)
    doc = Document(path)
    count = 0
    for p in _iter_all_paragraphs(doc):
        for r in p.runs:
            if old_font and (r.font.name or '').lower() != old_font.lower():
                continue
            r.font.name = new_font
            rpr = r._r.get_or_add_rPr()
            rf = rpr.get_or_add_rFonts()
            for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
                rf.set(qn(attr), new_font)
            if size is not None:
                r.font.size = Pt(float(size))
            count += 1
    doc.save(path)
    return {'ok': True, 'runs_changed': count, 'font': new_font}
