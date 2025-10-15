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
st.set_page_config(page_title="EML/MSG → DOCX/PDF", page_icon="📨", layout="centered")
st.title("📨 EML/MSG → DOCX/PDF")

pandoc_ok = have_prog("pandoc")
if not pandoc_ok:
    st.error("Pandoc not found. On Streamlit Cloud, add it to packages.txt. Locally, install via your package manager.")
if not HAS_EXTRACT_MSG:
    st.warning("`.msg` support not installed. Add the extract-msg stack to requirements.txt.")

c1, c2, c3 = st.columns(3)
with c1: want_docx = st.checkbox("Generate DOCX", value=True)
with c2: want_pdf  = st.checkbox("Generate PDF", value=False)
with c3:
    engine_choice = st.selectbox("PDF engine", ["(default)", "wkhtmltopdf", "xelatex"], index=1)
pdf_engine = None if engine_choice == "(default)" else engine_choice

uploads = st.file_uploader("Drop .eml/.msg files (multiple allowed)", type=["eml", "msg"], accept_multiple_files=True)

if st.button("Convert", disabled=not uploads or not pandoc_ok or (not want_docx and not want_pdf)):
    results = []
    zipper = io.BytesIO()
    with zipfile.ZipFile(zipper, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for uf in uploads:
            with st.status(f"Processing **{uf.name}**", expanded=False):
                base, outputs, err = process_one_uploaded(uf, want_docx, want_pdf, pdf_engine)
            if err:
                st.error(f"{uf.name}: {err}")
                continue
            if "docx" in outputs:
                docx_name = f"{base}.docx"
                st.download_button(f"⬇️ {docx_name}", outputs["docx"], docx_name,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                zf.writestr(docx_name, outputs["docx"])
            if "pdf" in outputs:
                pdf_name = f"{base}.pdf"
                st.download_button(f"⬇️ {pdf_name}", outputs["pdf"], pdf_name, "application/pdf")
                zf.writestr(pdf_name, outputs["pdf"])
            results.append(base)
    if results:
        zipper.seek(0)
        st.download_button("⬇️ Download all as ZIP", zipper.getvalue(),
                           "converted_emails.zip", "application/zip", use_container_width=True)

st.markdown("---")
st.caption("Tip: On Streamlit Cloud, PDFs work best with `wkhtmltopdf` (smaller footprint than TeX).")
