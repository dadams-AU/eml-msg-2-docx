import io, os, re, shutil, subprocess, tempfile, zipfile
from email import policy
import email
import streamlit as st

# Optional .msg support
try:
    import extract_msg  # pip install extract-msg olefile rtfde compressed-rtf tzlocal chardet beautifulsoup4
    HAS_EXTRACT_MSG = True
except Exception:
    HAS_EXTRACT_MSG = False

CID_RE = re.compile(r'cid:(?P<cid>[^"\'>\s)]+)', re.IGNORECASE)

def ensure_text(x):
    return x.decode("utf-8", "replace") if isinstance(x, (bytes, bytearray)) else x

def have_prog(name):
    try:
        subprocess.run([name, "-v"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        return True
    except Exception:
        return False

def decode_eml_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = (part.get("Content-Disposition") or "").lower()
            if ctype == "text/html" and "attachment" not in disp:
                return "text/html", part.get_content()
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = (part.get("Content-Disposition") or "").lower()
            if ctype == "text/plain" and "attachment" not in disp:
                return "text/plain", part.get_content()
        return None, None
    else:
        ctype = msg.get_content_type()
        if ctype in ("text/html", "text/plain"):
            return ctype, msg.get_content()
        return None, None

def collect_eml_cids(msg, outdir):
    cid_map, counter = {}, 0
    for part in msg.walk():
        cid = part.get("Content-ID")
        if not cid:
            continue
        cid = cid.strip("<> \t\r\n")
        ctype = part.get_content_type()
        if not (ctype.startswith("image/") or ctype == "application/octet-stream"):
            continue
        maintype, subtype = ctype.split("/", 1) if "/" in ctype else ("application", "bin")
        ext = {"jpeg":"jpg","pjpeg":"jpg","svg+xml":"svg"}.get(subtype.lower(), subtype.split("+",1)[0])
        filename = part.get_filename()
        if filename:
            _, ext_from_name = os.path.splitext(filename)
            if ext_from_name:
                ext = ext_from_name.lstrip(".")
        counter += 1
        path = os.path.join(outdir, f"cidimg_{counter}.{ext}")
        with open(path, "wb") as f:
            f.write(part.get_content())
        cid_map[cid] = path
    return cid_map

def parse_eml(path, workdir):
    with open(path, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)
    mime, body = decode_eml_body(msg)
    assets_dir = os.path.join(workdir, "assets")
    os.makedirs(assets_dir, exist_ok=True)
    cid_map = collect_eml_cids(msg, assets_dir)
    return mime, body, cid_map

def parse_msg(path, workdir):
    if not HAS_EXTRACT_MSG:
        raise RuntimeError("MSG support not installed (add to requirements.txt).")
    m = extract_msg.Message(path)
    body_html = ensure_text(getattr(m, "htmlBody", None))
    body_text = ensure_text(getattr(m, "body", None))
    if body_html:
        mime, body = "text/html", body_html
    elif body_text:
        mime, body = "text/plain", body_text
    else:
        raise RuntimeError("No usable body in .msg (no HTML or text).")
    assets_dir = os.path.join(workdir, "assets")
    os.makedirs(assets_dir, exist_ok=True)
    cid_map, counter = {}, 0
    for att in m.attachments:
        cid = ensure_text(getattr(att, "contentId", "")).strip("<> \t\r\n")
        if not cid:
            continue
        data = att.data  # bytes
        name = ensure_text(att.longFilename or att.shortFilename or f"cidimg_{counter}")
        _, ext = os.path.splitext(name)
        if not ext:
            ext = ".bin"
        counter += 1
        path = os.path.join(assets_dir, f"cidimg_{counter}{ext}")
        with open(path, "wb") as f:
            f.write(data)
        cid_map[cid] = path
    return mime, body, cid_map

def rewrite_cids(html, cid_map, rel_from):
    def repl(m):
        cid = m.group("cid")
        path = cid_map.get(cid) or cid_map.get(f"<{cid}>")
        if not path:
            return m.group(0)
        rel = os.path.relpath(path, start=rel_from)
        return f"{rel}"
    html = ensure_text(html)
    return CID_RE.sub(lambda m: repl(m), html)

def run_pandoc(src_html, want_docx, want_pdf, pdf_engine):
    outputs = {}
    if want_docx:
        out_docx = src_html.replace(".html", ".docx")
        subprocess.run(["pandoc", src_html, "-o", out_docx], check=True)
        with open(out_docx, "rb") as f:
            outputs["docx"] = f.read()
    if want_pdf:
        out_pdf = src_html.replace(".html", ".pdf")
        cmd = ["pandoc", src_html, "-o", out_pdf]
        if pdf_engine:
            cmd += ["--pdf-engine", pdf_engine]
        subprocess.run(cmd, check=True)
        with open(out_pdf, "rb") as f:
            outputs["pdf"] = f.read()
    return outputs

def process_one_uploaded(uploaded_file, want_docx, want_pdf, pdf_engine):
    work = tempfile.mkdtemp(prefix="eml2pandoc_")
    try:
        name = uploaded_file.name
        base, ext = os.path.splitext(name)
        ext = ext.lower()
        input_path = os.path.join(work, name)
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        if ext == ".eml":
            mime, body, cid_map = parse_eml(input_path, work)
        elif ext == ".msg":
            mime, body, cid_map = parse_msg(input_path, work)
        else:
            raise RuntimeError("Unsupported file type (use .eml or .msg).")

        if not body:
            raise RuntimeError("No body found.")

        html_path = os.path.join(work, f"{base}.html")
        if mime == "text/plain":
            with open(html_path, "w", encoding="utf-8") as f:
                f.write("<!doctype html><meta charset='utf-8'><pre>")
                f.write(ensure_text(body))
                f.write("</pre>")
        else:
            html = ensure_text(body)
            if cid_map:
                html = rewrite_cids(html, cid_map, work)
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html)

        outputs = run_pandoc(html_path, want_docx, want_pdf, pdf_engine)
        return base, outputs, None
    except subprocess.CalledProcessError as e:
        return uploaded_file.name, {}, f"Pandoc failed: {e}"
    except Exception as e:
        return uploaded_file.name, {}, str(e)
    finally:
        shutil.rmtree(work, ignore_errors=True)

# ---------------- UI ----------------
st.set_page_config(page_title="Email → Document Converter", page_icon="📨", layout="wide")

st.markdown("""
<style>
  /* tighten the sidebar */
  section[data-testid="stSidebar"] { min-width: 260px; }
  /* status badge pill */
  .status-ok   { background:#d4edda; color:#155724; padding:2px 10px; border-radius:12px; font-size:.85rem; }
  .status-warn { background:#fff3cd; color:#856404; padding:2px 10px; border-radius:12px; font-size:.85rem; }
  .status-err  { background:#f8d7da; color:#721c24; padding:2px 10px; border-radius:12px; font-size:.85rem; }
  /* file list rows */
  .file-row { padding:6px 10px; border-radius:6px; margin-bottom:4px; background:#f8f9fa; font-size:.9rem; }
  /* download section header */
  .dl-header { font-size:1.1rem; font-weight:600; margin-bottom:8px; }
</style>
""", unsafe_allow_html=True)

# ---- Sidebar ----
with st.sidebar:
    st.header("⚙️ Output Options")
    want_docx = st.checkbox("Generate DOCX", value=True)
    want_pdf  = st.checkbox("Generate PDF",  value=False)
    if want_pdf:
        engine_choice = st.selectbox(
            "PDF engine",
            ["(default)", "wkhtmltopdf", "xelatex"],
            index=1,
            help="wkhtmltopdf is recommended on Streamlit Cloud (smaller footprint than TeX).",
        )
    else:
        engine_choice = "(default)"

    st.divider()
    st.subheader("System Status")
    pandoc_ok = have_prog("pandoc")
    if pandoc_ok:
        st.markdown('<span class="status-ok">✔ Pandoc ready</span>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="status-err">✘ Pandoc not found</span>', unsafe_allow_html=True)
        st.caption("Add `pandoc` to `packages.txt` (Streamlit Cloud) or install via your package manager.")

    st.write("")
    if HAS_EXTRACT_MSG:
        st.markdown('<span class="status-ok">✔ .msg support ready</span>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="status-warn">⚠ .msg support missing</span>', unsafe_allow_html=True)
        st.caption("Add the `extract-msg` stack to `requirements.txt` to enable `.msg` files.")

    st.divider()
    with st.expander("How it works"):
        st.markdown(
            "1. Upload one or more `.eml` / `.msg` files.\n"
            "2. The app extracts HTML (or plain text) and any inline `cid:` images.\n"
            "3. [Pandoc](https://pandoc.org) converts the result to DOCX and/or PDF.\n"
            "4. Download individual files or grab them all in a ZIP."
        )

pdf_engine = None if engine_choice == "(default)" else engine_choice

# ---- Main area ----
st.title("📨 Email → Document Converter")
st.caption("Convert `.eml` and `.msg` email files to DOCX or PDF — inline images included.")

st.divider()

uploads = st.file_uploader(
    "Drop `.eml` / `.msg` files here (multiple allowed)",
    type=["eml", "msg"],
    accept_multiple_files=True,
    label_visibility="visible",
)

if uploads:
    st.markdown(f"**{len(uploads)} file{'s' if len(uploads) != 1 else ''} queued**")
    for uf in uploads:
        size_kb = len(uf.getbuffer()) / 1024
        size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.2f} MB"
        ext_icon = "📧" if uf.name.lower().endswith(".eml") else "📁"
        st.markdown(
            f'<div class="file-row">{ext_icon} <b>{uf.name}</b> &nbsp;·&nbsp; {size_str}</div>',
            unsafe_allow_html=True,
        )
    st.write("")

can_convert = bool(uploads) and pandoc_ok and (want_docx or want_pdf)
if not want_docx and not want_pdf:
    st.warning("Select at least one output format (DOCX or PDF) in the sidebar.")

if st.button("Convert", type="primary", disabled=not can_convert, use_container_width=True):
    results, errors = [], []
    zipper = io.BytesIO()
    progress = st.progress(0, text="Starting…")

    with zipfile.ZipFile(zipper, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for i, uf in enumerate(uploads):
            progress.progress((i) / len(uploads), text=f"Converting **{uf.name}** ({i+1}/{len(uploads)})…")
            with st.status(f"Processing **{uf.name}**", expanded=False):
                base, outputs, err = process_one_uploaded(uf, want_docx, want_pdf, pdf_engine)
            if err:
                errors.append((uf.name, err))
                continue

            dl_cols = st.columns([1, 1, 2])
            if "docx" in outputs:
                docx_name = f"{base}.docx"
                dl_cols[0].download_button(
                    "⬇️ DOCX", outputs["docx"], docx_name,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"docx_{i}",
                )
                zf.writestr(docx_name, outputs["docx"])
            if "pdf" in outputs:
                pdf_name = f"{base}.pdf"
                dl_cols[1].download_button(
                    "⬇️ PDF", outputs["pdf"], pdf_name,
                    "application/pdf",
                    key=f"pdf_{i}",
                )
                zf.writestr(pdf_name, outputs["pdf"])
            if outputs:
                dl_cols[2].caption(f"✔ {uf.name}")
                results.append(base)

    progress.progress(1.0, text="Done!")

    for fname, err in errors:
        st.error(f"**{fname}**: {err}")

    if len(results) > 1:
        zipper.seek(0)
        st.divider()
        st.download_button(
            "⬇️ Download all as ZIP",
            zipper.getvalue(),
            "converted_emails.zip",
            "application/zip",
            use_container_width=True,
            type="primary",
        )

    if results:
        st.success(f"Converted {len(results)} file{'s' if len(results) != 1 else ''} successfully.")
