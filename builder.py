"""
Builder — takes structured CV data and fills the DIGITALL DOCX template.
Works by unzipping the template, doing XML-level manipulation, and rezipping.
"""
import zipfile
import io
import copy
import re
import os
from lxml import etree

# ---------------------------------------------------------------------------
# Namespace helpers
# ---------------------------------------------------------------------------

W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

def _w(tag):
    return f"{{{W}}}{tag}"

def _get_text(element):
    """Concatenate all w:t text under element."""
    return "".join((t.text or "") for t in element.iter(_w("t")))

def _set_run_text(run, text):
    """Replace text in a w:r element, keeping its w:rPr."""
    for t in run.findall(_w("t")):
        run.remove(t)
    t_el = etree.SubElement(run, _w("t"))
    t_el.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t_el.set(XML_SPACE, "preserve")

def _clear_runs(para):
    """Remove all w:r elements from a paragraph."""
    for r in para.findall(_w("r")):
        para.remove(r)

def _first_rpr(para):
    """Return a deep copy of the first run's rPr, or None."""
    r = para.find(_w("r"))
    if r is None:
        return None
    rpr = r.find(_w("rPr"))
    return copy.deepcopy(rpr) if rpr is not None else None

def _add_run(para, text, rpr=None):
    """Append a new w:r with optional rPr to para."""
    r = etree.SubElement(para, _w("r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, _w("t"))
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set(XML_SPACE, "preserve")
    return r


# ---------------------------------------------------------------------------
# Paragraph finders
# ---------------------------------------------------------------------------

def _find_para_by_text(body, target, exact=False):
    """Find the first w:p whose concatenated text matches target."""
    for p in body.iter(_w("p")):
        txt = _get_text(p).strip()
        if exact:
            if txt == target:
                return p
        else:
            if target in txt:
                return p
    return None

def _paragraphs_between(body, start_text, end_text):
    """
    Return all w:p elements that come AFTER the paragraph containing start_text
    and BEFORE the paragraph containing end_text.
    Works across direct children and inside table cells.
    """
    all_paras = list(body.iter(_w("p")))
    start_idx = end_idx = None
    for i, p in enumerate(all_paras):
        txt = _get_text(p).strip()
        if start_idx is None and start_text in txt:
            start_idx = i
        elif start_idx is not None and end_idx is None and end_text in txt:
            end_idx = i
            break
    if start_idx is None:
        return []
    if end_idx is None:
        end_idx = len(all_paras)
    return all_paras[start_idx + 1 : end_idx]


# ---------------------------------------------------------------------------
# Section replacers
# ---------------------------------------------------------------------------

def _replace_job_title(body, title):
    """Replace the job-title paragraph (large grey text, sz=40, color=869CAD)."""
    for p in body.iter(_w("p")):
        txt = _get_text(p).strip()
        if not txt:
            continue
        # Check for the big grey style (sz=40, color 869CAD)
        rpr = p.find(f".//{_w('rPr')}")
        sz = p.find(f".//{_w('sz')}")
        col = p.find(f".//{_w('color')}")
        has_big = False
        for r in p.iter(_w("r")):
            rp = r.find(_w("rPr"))
            if rp is not None:
                sz_el = rp.find(_w("sz"))
                col_el = rp.find(_w("color"))
                if sz_el is not None and sz_el.get(_w("val")) == "40":
                    if col_el is not None and "869CAD" in (col_el.get(_w("val")) or "").upper():
                        has_big = True
                        break
        # Skip section headings
        if has_big and txt not in ("Summary", "Technologies", "Education", "Languages",
                                    "Courses/Certificates/Trainings", "Work Experience"):
            rpr_copy = _first_rpr(p)
            _clear_runs(p)
            _add_run(p, title, rpr_copy)
            return


def _replace_summary(body, summary_text):
    """Replace the summary paragraphs (between Summary header and Technologies header)."""
    paras = _paragraphs_between(body, "Summary", "Technologies")
    if not paras:
        return
    # Keep the first content paragraph's formatting as reference
    ref_rpr = _first_rpr(paras[0]) if paras else None
    # Remove all content from existing paragraphs and clear all but the first
    parent_map = {c: p for p in body.iter() for c in p}
    # Replace first para with new text, remove others
    for i, p in enumerate(paras):
        if i == 0:
            _clear_runs(p)
            _add_run(p, summary_text, ref_rpr)
        else:
            par = parent_map.get(p)
            if par is not None:
                try:
                    par.remove(p)
                except ValueError:
                    pass


def _find_tech_para(body, label_text):
    """Find the 'Aufzhlung' paragraph that has label_text in bold."""
    for p in body.iter(_w("p")):
        pstyle = p.find(f".//{_w('pStyle')}")
        if pstyle is None or pstyle.get(_w("val")) != "Aufzhlung":
            continue
        txt = _get_text(p)
        if label_text.rstrip(":").lower() in txt.lower():
            return p
    return None


def _replace_tech_line(body, label_text, new_value):
    """
    Replace the value part of a Technologies bullet like:
      [bold] Primary expertise[/bold]: [regular] old value
    If new_value is empty, the paragraph is left as-is and will be cleaned
    by _remove_empty_tech_bullets() afterwards.
    """
    p = _find_tech_para(body, label_text)
    if p is None:
        return

    runs = p.findall(_w("r"))
    if not runs:
        return

    # Find the last non-bold run (that's the value run) and update it
    value_runs = []
    for r in runs:
        rpr = r.find(_w("rPr"))
        # Only w:b (not w:bCs) makes Latin text visually bold
        bold = (rpr is not None and rpr.find(_w("b")) is not None)
        if not bold:
            value_runs.append(r)

    if value_runs:
        # Replace text in last value run, remove others
        for r in value_runs[:-1]:
            p.remove(r)
        _set_run_text(value_runs[-1], new_value)
    else:
        # All runs are bold; append a new value run
        last_r = runs[-1]
        rpr_copy = _first_rpr(p)
        # Remove bold from the copy
        if rpr_copy is not None:
            for b in rpr_copy.findall(_w("b")):
                rpr_copy.remove(b)
            for b in rpr_copy.findall(_w("bCs")):
                rpr_copy.remove(b)
        _add_run(p, new_value, rpr_copy)


def _remove_empty_tech_bullets(body):
    """
    After filling tech lines, remove ALL empty Aufzhlung paragraphs in the
    Technologies section (both the spacers and any empty field lines).
    """
    all_paras = list(body.iter(_w("p")))
    start_idx = end_idx = None
    for i, p in enumerate(all_paras):
        txt = _get_text(p).strip()
        if start_idx is None and txt == "Technologies":
            start_idx = i
        elif start_idx is not None and txt == "Education":
            end_idx = i
            break

    if start_idx is None:
        return

    end_idx = end_idx or len(all_paras)
    parent_map = {c: p for p in body.iter() for c in p}

    for p in all_paras[start_idx + 1 : end_idx]:
        pstyle = p.find(f".//{_w('pStyle')}")
        if pstyle is None or pstyle.get(_w("val")) != "Aufzhlung":
            continue
        txt = _get_text(p).strip()
        # Remove if empty OR if it's a label-only line with no value (e.g. "Programming languages: ")
        is_empty = not txt
        is_label_only = txt.endswith(":") or (": " in txt and txt.split(": ", 1)[1].strip() == "")
        if is_empty or is_label_only:
            parent = parent_map.get(p)
            if parent is not None:
                try:
                    parent.remove(p)
                except ValueError:
                    pass


def _replace_education(body, edu_list):
    """Replace education fields (Degree, Institute, Location, Year lines).
    Note: the DIGITALL template has a single education block.
    Only the first entry is written; extras are ignored (visible in the form)."""
    if not edu_list:
        return
    edu = edu_list[0]  # Template only has one education slot

    _replace_inline_field(body, "Degree:", edu.get("degree", ""))
    _replace_inline_field(body, "Institute:", edu.get("institute", ""))
    _replace_inline_field(body, "Location:", edu.get("location", ""))
    _replace_inline_field(body, "Year:", edu.get("year", ""))


def _replace_inline_field(body, field_label, new_value):
    """
    Find a paragraph that contains field_label (e.g. 'Degree:') and replace
    the text AFTER the label with new_value.
    """
    for p in body.iter(_w("p")):
        txt = _get_text(p)
        if field_label in txt:
            # Strategy: keep all runs that form the label, replace/add value run
            runs = p.findall(_w("r"))
            # Find which run contains the end of the label
            accumulated = ""
            label_end_run_idx = -1
            for i, r in enumerate(runs):
                t = _get_text(r)
                accumulated += t
                if field_label in accumulated:
                    label_end_run_idx = i
                    break
            if label_end_run_idx < 0:
                continue
            # Remove all runs after the label run
            for r in runs[label_end_run_idx + 1:]:
                p.remove(r)
            # Trim label run to just the label portion if it has extra value text
            label_run = runs[label_end_run_idx]
            label_run_text = _get_text(label_run)
            if field_label in label_run_text:
                colon_pos = label_run_text.index(field_label) + len(field_label)
                _set_run_text(label_run, label_run_text[:colon_pos] + " ")
            # Add a new run with the value
            rpr = _first_rpr(p)
            _add_run(p, new_value, rpr)
            return


def _replace_language_list(body, languages):
    """Replace language entries after the Languages header."""
    paras = _paragraphs_between(body, "Languages", "Courses/Certificates/Trainings")
    if not paras:
        return
    ref_rpr = _first_rpr(paras[0]) if paras else None
    parent_map = {c: p for p in body.iter() for c in p}

    for i, p in enumerate(paras):
        if i < len(languages):
            lang = languages[i]
            _clear_runs(p)
            _add_run(p, f"{lang['language']}: {lang['level']}", ref_rpr)
        else:
            par = parent_map.get(p)
            if par is not None:
                try:
                    par.remove(p)
                except ValueError:
                    pass

    # Add extra paragraphs if we have more languages than template slots
    if len(languages) > len(paras):
        last_para = paras[-1] if paras else None
        if last_para is not None:
            par = parent_map.get(last_para)
            if par is not None:
                insert_idx = list(par).index(last_para) + 1
                for lang in languages[len(paras):]:
                    new_p = copy.deepcopy(last_para)
                    _clear_runs(new_p)
                    _add_run(new_p, f"{lang['language']}: {lang['level']}", ref_rpr)
                    par.insert(insert_idx, new_p)
                    insert_idx += 1


def _replace_certifications(body, certs):
    """Replace certification bullet entries."""
    # Find the section by looking for Aufzhlung paragraphs after the certs header
    header_para = _find_para_by_text(body, "Courses/Certificates/Trainings")
    if header_para is None:
        return

    all_paras = list(body.iter(_w("p")))
    try:
        start_idx = all_paras.index(header_para) + 1
    except ValueError:
        return

    # Collect existing cert paragraphs (Aufzhlung style after the header)
    cert_paras = []
    for p in all_paras[start_idx:]:
        pstyle = p.find(f".//{_w('pStyle')}")
        if pstyle is not None and pstyle.get(_w("val")) == "Aufzhlung":
            cert_paras.append(p)
        else:
            # Stop at the next major section (first non-empty, non-Aufzhlung paragraph)
            if _get_text(p).strip():
                break

    if not cert_paras:
        return

    ref_rpr = _first_rpr(cert_paras[0])
    ref_p = cert_paras[0]
    parent_map = {c: p for p in body.iter() for c in p}

    for i, p in enumerate(cert_paras):
        if i < len(certs):
            _clear_runs(p)
            _add_run(p, certs[i], ref_rpr)
        else:
            par = parent_map.get(p)
            if par is not None:
                try:
                    par.remove(p)
                except ValueError:
                    pass

    # Add extra if needed
    if len(certs) > len(cert_paras):
        last_p = cert_paras[-1]
        par = parent_map.get(last_p)
        if par is not None:
            idx = list(par).index(last_p) + 1
            for cert in certs[len(cert_paras):]:
                new_p = copy.deepcopy(ref_p)
                _clear_runs(new_p)
                _add_run(new_p, cert, ref_rpr)
                par.insert(idx, new_p)
                idx += 1


# ---------------------------------------------------------------------------
# Work Experience table builder
# ---------------------------------------------------------------------------

def _make_label_cell(label, color="4CC0FF", border_spec=None, width=2972):
    """Build the left 'label' table cell."""
    tc = etree.Element(_w("tc"))
    tcp = etree.SubElement(tc, _w("tcPr"))
    etree.SubElement(tcp, _w("tcW"), {_w("w"): str(width), _w("type"): "dxa"})
    if border_spec:
        borders = etree.SubElement(tcp, _w("tcBorders"))
        for side, spec in border_spec.items():
            el = etree.SubElement(borders, _w(side))
            for k, v in spec.items():
                el.set(_w(k), v)
    etree.SubElement(tcp, _w("shd"), {_w("val"): "clear", _w("color"): "auto", _w("fill"): "auto"})

    p = etree.SubElement(tc, _w("p"))
    ppr = etree.SubElement(p, _w("pPr"))
    etree.SubElement(ppr, _w("jc"), {_w("val"): "left"})
    rpr_el = etree.SubElement(ppr, _w("rPr"))
    etree.SubElement(rpr_el, _w("rFonts"), {_w("cs"): "Arial"})
    etree.SubElement(rpr_el, _w("b"))
    etree.SubElement(rpr_el, _w("bCs"))
    etree.SubElement(rpr_el, _w("color"), {_w("val"): color})
    etree.SubElement(rpr_el, _w("sz"), {_w("val"): "24"})
    etree.SubElement(rpr_el, _w("szCs"), {_w("val"): "24"})

    r = etree.SubElement(p, _w("r"))
    rpr2 = etree.SubElement(r, _w("rPr"))
    etree.SubElement(rpr2, _w("rFonts"), {_w("cs"): "Arial"})
    etree.SubElement(rpr2, _w("b"))
    etree.SubElement(rpr2, _w("bCs"))
    etree.SubElement(rpr2, _w("color"), {_w("val"): color})
    etree.SubElement(rpr2, _w("sz"), {_w("val"): "24"})
    etree.SubElement(rpr2, _w("szCs"), {_w("val"): "24"})
    t = etree.SubElement(r, _w("t"))
    t.text = label
    return tc


def _make_value_cell(text, bold=False, border_spec=None, width=6095):
    """Build the right 'value' table cell with a single paragraph."""
    tc = etree.Element(_w("tc"))
    tcp = etree.SubElement(tc, _w("tcPr"))
    etree.SubElement(tcp, _w("tcW"), {_w("w"): str(width), _w("type"): "dxa"})
    if border_spec:
        borders = etree.SubElement(tcp, _w("tcBorders"))
        for side, spec in border_spec.items():
            el = etree.SubElement(borders, _w(side))
            for k, v in spec.items():
                el.set(_w(k), v)
    etree.SubElement(tcp, _w("shd"), {_w("val"): "clear", _w("color"): "auto", _w("fill"): "auto"})

    p = etree.SubElement(tc, _w("p"))
    ppr = etree.SubElement(p, _w("pPr"))
    rpr_ppr = etree.SubElement(ppr, _w("rPr"))
    etree.SubElement(rpr_ppr, _w("rFonts"), {_w("cs"): "Arial"})
    if bold:
        etree.SubElement(rpr_ppr, _w("b"))
        etree.SubElement(rpr_ppr, _w("bCs"))
    etree.SubElement(rpr_ppr, _w("sz"), {_w("val"): "24"})
    etree.SubElement(rpr_ppr, _w("szCs"), {_w("val"): "24"})

    if text:
        r = etree.SubElement(p, _w("r"))
        rpr2 = etree.SubElement(r, _w("rPr"))
        etree.SubElement(rpr2, _w("rFonts"), {_w("cs"): "Arial"})
        if bold:
            etree.SubElement(rpr2, _w("b"))
            etree.SubElement(rpr2, _w("bCs"))
        etree.SubElement(rpr2, _w("sz"), {_w("val"): "24"})
        etree.SubElement(rpr2, _w("szCs"), {_w("val"): "24"})
        t = etree.SubElement(r, _w("t"))
        t.text = text
        if text.startswith(" ") or text.endswith(" "):
            t.set(XML_SPACE, "preserve")
    return tc


def _make_responsibilities_cell(responsibilities, border_spec=None, width=6095):
    """Build the right cell for responsibilities (bullet list)."""
    tc = etree.Element(_w("tc"))
    tcp = etree.SubElement(tc, _w("tcPr"))
    etree.SubElement(tcp, _w("tcW"), {_w("w"): str(width), _w("type"): "dxa"})
    if border_spec:
        borders = etree.SubElement(tcp, _w("tcBorders"))
        for side, spec in border_spec.items():
            el = etree.SubElement(borders, _w(side))
            for k, v in spec.items():
                el.set(_w(k), v)
    etree.SubElement(tcp, _w("shd"), {_w("val"): "clear", _w("color"): "auto", _w("fill"): "auto"})

    # Always have at least one paragraph so the cell isn't empty (Word requires it)
    # but don't render a visible blank bullet — use a truly empty paragraph instead.
    if not responsibilities:
        p = etree.SubElement(tc, _w("p"))
        etree.SubElement(p, _w("pPr"))
        return tc

    for item in responsibilities:
        p = etree.SubElement(tc, _w("p"))
        ppr = etree.SubElement(p, _w("pPr"))
        etree.SubElement(ppr, _w("pStyle"), {_w("val"): "Aufzhlung"})
        rpr_ppr = etree.SubElement(ppr, _w("rPr"))
        etree.SubElement(rpr_ppr, _w("lang"), {_w("val"): "en-US"})

        if item:
            r = etree.SubElement(p, _w("r"))
            rpr2 = etree.SubElement(r, _w("rPr"))
            etree.SubElement(rpr2, _w("lang"), {_w("val"): "en-US"})
            t = etree.SubElement(r, _w("t"))
            t.text = item
            if item.endswith(" "):
                t.set(XML_SPACE, "preserve")
    return tc


_BLUE_TOP = {"val": "single", "sz": "24", "space": "0", "color": "21408C"}
_THIN = {"val": "single", "sz": "4", "space": "0", "color": "auto"}
_BLUE_BOTTOM = {"val": "single", "sz": "4", "space": "0", "color": "21408C"}


def _make_job_rows(job):
    """Generate the 6 table rows for one work experience entry."""
    rows = []

    def make_tr(cnf_odd=False):
        tr = etree.Element(_w("tr"))
        trpr = etree.SubElement(tr, _w("trPr"))
        cnf_val = "000000100000" if cnf_odd else "000000000000"
        etree.SubElement(trpr, _w("cnfStyle"), {
            _w("val"): cnf_val,
            _w("firstRow"): "0", _w("lastRow"): "0",
            _w("firstColumn"): "0", _w("lastColumn"): "0",
            _w("oddVBand"): "0", _w("evenVBand"): "0",
            _w("oddHBand"): "1" if cnf_odd else "0",
            _w("evenHBand"): "0",
            _w("firstRowFirstColumn"): "0", _w("firstRowLastColumn"): "0",
            _w("lastRowFirstColumn"): "0", _w("lastRowLastColumn"): "0",
        })
        return tr

    # Row 1: Project Name (header)
    tr1 = make_tr(cnf_odd=True)
    label1 = _make_label_cell("Project Name", color="21408C", border_spec={
        "top": _BLUE_TOP, "bottom": _THIN, "right": _THIN,
    })
    # add cnfStyle to tcPr
    tcp1 = label1.find(_w("tcPr"))
    cnf_tc = etree.Element(_w("cnfStyle"))
    cnf_tc.set(_w("val"), "000010000000")
    for k in ["firstRow","lastRow","firstColumn","lastColumn","oddVBand","evenVBand","oddHBand","evenHBand",
              "firstRowFirstColumn","firstRowLastColumn","lastRowFirstColumn","lastRowLastColumn"]:
        cnf_tc.set(_w(k), "1" if k == "oddVBand" else "0")
    tcp1.insert(0, cnf_tc)

    val1 = _make_value_cell(job.get("project_name", ""), bold=False, border_spec={
        "top": _BLUE_TOP, "left": _THIN, "bottom": _THIN,
    })
    tr1.append(label1)
    tr1.append(val1)
    rows.append(tr1)

    # Row 2: Employer
    tr2 = make_tr(cnf_odd=False)
    tr2.append(_make_label_cell("Employer", color="4CC0FF", border_spec={
        "top": _THIN, "right": _THIN,
    }))
    tr2.append(_make_value_cell(job.get("employer", "Confidential"), bold=True, border_spec={
        "top": _THIN, "left": _THIN,
    }))
    rows.append(tr2)

    # Row 3: Duration
    tr3 = make_tr(cnf_odd=True)
    trpr3 = tr3.find(_w("trPr"))
    etree.SubElement(trpr3, _w("trHeight"), {_w("val"): "260"})
    tr3.append(_make_label_cell("Duration", color="4CC0FF", border_spec={
        "right": _THIN,
    }))
    tr3.append(_make_value_cell(job.get("duration", ""), bold=True, border_spec={
        "left": _THIN,
    }))
    rows.append(tr3)

    # Row 4: Responsibilities
    tr4 = make_tr(cnf_odd=False)
    tr4.append(_make_label_cell("Responsibilities", color="4CC0FF", border_spec={
        "right": _THIN,
    }))
    tr4.append(_make_responsibilities_cell(job.get("responsibilities", []), border_spec={
        "left": _THIN,
    }))
    rows.append(tr4)

    # Row 5: Position
    tr5 = make_tr(cnf_odd=True)
    trpr5 = tr5.find(_w("trPr"))
    etree.SubElement(trpr5, _w("trHeight"), {_w("val"): "401"})
    tr5.append(_make_label_cell("Position", color="4CC0FF", border_spec={
        "right": _THIN,
    }))
    tr5.append(_make_value_cell(job.get("position", ""), bold=True, border_spec={
        "left": _THIN,
    }))
    rows.append(tr5)

    # Row 6: Location (bottom border blue)
    tr6 = make_tr(cnf_odd=False)
    tr6.append(_make_label_cell("Location", color="4CC0FF", border_spec={
        "bottom": _BLUE_BOTTOM, "right": _THIN,
    }))
    tr6.append(_make_value_cell(job.get("location", ""), bold=True, border_spec={
        "left": _THIN,
    }))
    rows.append(tr6)

    return rows


def _rebuild_work_experience_table(body, jobs):
    """Find the work experience table and replace its content rows."""
    tables = body.findall(f".//{_w('tbl')}")
    # The work experience table is the second table (index 1)
    if len(tables) < 2:
        return
    tbl = tables[1]

    # Remove all existing rows
    for tr in tbl.findall(_w("tr")):
        tbl.remove(tr)

    # Generate rows for each job
    for job in jobs:
        for row in _make_job_rows(job):
            tbl.append(row)


# ---------------------------------------------------------------------------
# Name helpers
# ---------------------------------------------------------------------------

def _initials(full_name):
    """'Bozhidar Gavrilov' → 'B.G.'"""
    parts = full_name.strip().split()
    return ".".join(p[0].upper() for p in parts if p) + "." if parts else "N.N."


def _replace_name_in_header(header_xml_bytes, display_name):
    """
    Replace the candidate name run in a header XML.
    The name is the grey (color=7F7F7F) paragraph in header2.xml.
    """
    root = etree.fromstring(header_xml_bytes)
    for p in root.iter(_w("p")):
        for r in p.findall(_w("r")):
            rpr = r.find(_w("rPr"))
            if rpr is None:
                continue
            col = rpr.find(_w("color"))
            if col is not None and (col.get(_w("val")) or "").upper() in ("7F7F7F", "808080"):
                # This is the name run — replace its text
                for t in r.findall(_w("t")):
                    r.remove(t)
                t_el = etree.SubElement(r, _w("t"))
                t_el.text = display_name
                if display_name and (display_name[0] == " " or display_name[-1] == " "):
                    t_el.set(XML_SPACE, "preserve")
                return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    # No name run found — return unchanged
    return header_xml_bytes


# ---------------------------------------------------------------------------
# Confidential text scrubber
# ---------------------------------------------------------------------------

def _scrub_text(text, full_name, initials, employers=None):
    """
    Replace the candidate's name and employer names in free-text fields
    so no identifying information leaks into a confidential CV.

    Replaces:
      - Full name  (e.g. "Dimitar Draganchev")  → initials (e.g. "D.D.")
      - First name alone  (e.g. "Dimitar")        → initials
      - Each employer name                         → "Confidential"
    """
    if not text:
        return text

    if full_name:
        # Full name — most specific, replace first
        text = re.sub(re.escape(full_name), initials, text, flags=re.I)
        # First name alone (only if long enough to avoid false positives)
        parts = full_name.split()
        if parts and len(parts[0]) > 3:
            text = re.sub(r"\b" + re.escape(parts[0]) + r"\b", initials, text, flags=re.I)
        # Last name alone (only if long enough)
        if len(parts) > 1 and len(parts[-1]) > 3:
            text = re.sub(r"\b" + re.escape(parts[-1]) + r"\b", initials, text, flags=re.I)

    for employer in (employers or []):
        if employer and len(employer) > 3:
            text = re.sub(r"\b" + re.escape(employer) + r"\b", "Confidential", text, flags=re.I)

    return text


# ---------------------------------------------------------------------------
# Main build function
# ---------------------------------------------------------------------------

def build_cv(template_path, data, output_path, confidential=True):
    """
    Fill the DIGITALL template with candidate data and save to output_path.

    confidential=True  → name shown as initials (e.g. B.G.), employers shown as 'Confidential'
    confidential=False → full name and real employer names shown

    data dict keys:
        name, job_title, summary,
        technologies: {primary_expertise, programming_languages, methodologies, other_skills},
        education: [{degree, institute, location, year}, ...],
        languages: [{language, level}, ...],
        certifications: [str, ...],
        work_experience: [{project_name, employer, duration, responsibilities, position, location}, ...]
    """
    full_name = data.get("name", "").strip()
    display_name = _initials(full_name) if confidential else full_name

    # Collect original employer names before masking (needed for text scrubbing)
    original_employers = [
        j.get("employer", "").strip()
        for j in data.get("work_experience", [])
        if j.get("employer", "").strip().lower() not in ("", "confidential")
    ]

    # Deep-copy work experience and mask employers when confidential
    jobs = copy.deepcopy(data.get("work_experience", []))
    if confidential:
        for job in jobs:
            job["employer"] = "Confidential"

    # Scrub name & employer names from all free-text fields
    summary = data.get("summary", "")
    if confidential:
        summary = _scrub_text(summary, full_name, display_name, original_employers)
        for job in jobs:
            job["responsibilities"] = [
                _scrub_text(r, full_name, display_name, original_employers)
                for r in job.get("responsibilities", [])
            ]

    # --- Read template ZIP in memory ---
    with open(template_path, "rb") as f:
        template_bytes = f.read()

    in_buf = io.BytesIO(template_bytes)
    out_buf = io.BytesIO()

    with zipfile.ZipFile(in_buf, "r") as zin, zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            raw = zin.read(item.filename)

            if item.filename == "word/document.xml":
                root = etree.fromstring(raw)
                body = root.find(_w("body"))

                tech  = data.get("technologies", {})
                edu   = data.get("education", [])
                langs = data.get("languages", [])
                certs = data.get("certifications", [])

                _replace_job_title(body, data.get("job_title", ""))
                _replace_summary(body, summary)
                _replace_tech_line(body, "Primary expertise",    tech.get("primary_expertise", ""))
                _replace_tech_line(body, "Programming languages", tech.get("programming_languages", ""))
                _replace_tech_line(body, "Methodologies",        tech.get("methodologies", ""))
                _replace_tech_line(body, "Other Skills",         tech.get("other_skills", ""))
                _remove_empty_tech_bullets(body)   # clean leftover empty/spacer bullets
                _replace_education(body, edu)
                _replace_language_list(body, langs)
                _replace_certifications(body, certs)
                _rebuild_work_experience_table(body, jobs)

                raw = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

            elif item.filename == "word/header2.xml":
                raw = _replace_name_in_header(raw, display_name)

            zout.writestr(item, raw)

    with open(output_path, "wb") as f:
        f.write(out_buf.getvalue())
