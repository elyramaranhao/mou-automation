import io
import os
import re
import zipfile
import tempfile
import subprocess
from datetime import datetime
from typing import Dict, List, Set

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt
from pydantic import BaseModel, Field, validator

st.set_page_config(page_title="Gerador de MOU – (sem Google)", page_icon="📝", layout="wide")

PLACEHOLDER_RE = re.compile(r"\{\{([A-Z0-9_]+)\}\}", re.IGNORECASE)

# ---------------------------
# Utilitários .docx
# ---------------------------
def _iter_all_paragraphs(doc: Document):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for section in doc.sections:
        if section.header:
            for p in section.header.paragraphs:
                yield p
        if section.footer:
            for p in section.footer.paragraphs:
                yield p

def _para_text(p) -> str:
    txt = "".join(run.text for run in p.runs) or p.text or ""
    # normaliza espaços para facilitar o match
    return re.sub(r"\s+", " ", txt).strip()

def extract_placeholders(doc: Document) -> Set[str]:
    found: Set[str] = set()
    for p in _iter_all_paragraphs(doc):
        text = _para_text(p)
        for m in PLACEHOLDER_RE.finditer(text):
            found.add(m.group(1).strip().upper())
    return found

def _is_exception_phrase(text_low: str) -> bool:
    """Parágrafos que DEVEM ficar sem negrito (PT/EN)."""
    # 1) Frase longa de requisitos
    if text_low.startswith("como parte integrante deste documento") and "continuidade do processo" in text_low:
        return True
    if text_low.startswith("as an integral part of this document") and "continuity of the nomination process" in text_low:
        return True
    # 2) Linha do Business Plan
    if text_low.startswith("business plan apresentado e validado em"):
        return True
    if text_low.startswith("business plan presented") and ("validated" in text_low or "approved" in text_low) and ("on" in text_low or ":" in text_low):
        return True
    # 3) Título da seção 2
    if text_low.startswith("2.") and "especifica" in text_low and ("acordadas" in text_low or "alterações" in text_low or "alteracoes" in text_low):
        return True
    if text_low.startswith("2.") and "specification" in text_low and ("agreed" in text_low or "approved" in text_low) and ("change" in text_low or "modification" in text_low):
        return True
    return False

def replace_placeholders_and_collect_exceptions(doc: Document, mapping: Dict[str, str]):
    """
    1) Marca como exceção (sem negrito) os parágrafos que CONTÊM {{BP_DATE}} e {{COMMENTS}}.
    2) Faz replace dos placeholders (case-insensitive) no documento inteiro.
    Retorna set com referências de parágrafos que devem ficar sem negrito.
    """
    normalized = {f"{{{{{k}}}}}": str(v) for k, v in mapping.items()}
    exceptions = set()  # parágrafos que NÃO devem ficar em negrito

    for p in _iter_all_paragraphs(doc):
        orig = _para_text(p)
        low = orig.lower()
        # marca exceções por placeholder
        if "{{bp_date}}" in low or "{{comments}}" in low:
            exceptions.add(p)

        # replace case-insensitive
        replaced = orig
        for k, v in normalized.items():
            pattern = re.compile(re.escape(k), re.IGNORECASE)
            replaced = pattern.sub(v, replaced)

        if replaced != orig:
            # limpar runs e escrever como único run
            for _ in range(len(p.runs)):
                p.runs[0].clear()
                p.runs[0].text = ""
                p._element.remove(p.runs[0]._element)
            p.add_run(replaced)

    return exceptions

def enforce_calibri11_and_bold_with_exceptions(doc: Document, exceptions: Set):
    """
    Aplica Calibri 11 e negrito em tudo, depois remove negrito
    1) nos parágrafos coletados em 'exceptions' e
    2) nos que casam com as frases de exceção (PT/EN).
    """
    # 1) Calibri 11 + bold=True em todo mundo
    for p in _iter_all_paragraphs(doc):
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(11)
            run.bold = True

    # 2) Desmarca bold nos parágrafos de exceção
    for p in _iter_all_paragraphs(doc):
        text_low = _para_text(p).lower()
        if (p in exceptions) or _is_exception_phrase(text_low):
            for run in p.runs:
                run.bold = False

# ---------------------------
# PDF
# ---------------------------
def try_export_pdf(doc_bytes: bytes) -> bytes:
    """
    Tenta converter DOCX->PDF com docx2pdf (Word), se falhar tenta LibreOffice.
    Retorna bytes do PDF ou levanta Exception.
    """
    # caminho temporário
    with tempfile.TemporaryDirectory() as td:
        docx_path = os.path.join(td, "out.docx")
        pdf_path = os.path.join(td, "out.pdf")
        with open(docx_path, "wb") as f:
            f.write(doc_bytes)

        # 1) Tenta docx2pdf (somente macOS/Windows com Word)
        try:
            from docx2pdf import convert
            convert(docx_path, pdf_path)
            with open(pdf_path, "rb") as f:
                return f.read()
        except Exception:
            pass

        # 2) Tenta LibreOffice (precisa 'soffice' no PATH)
        try:
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf", "--outdir", td, docx_path],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )
            with open(pdf_path, "rb") as f:
                return f.read()
        except Exception as e:
            raise RuntimeError("Não foi possível gerar PDF (docx2pdf/LibreOffice indisponíveis).") from e

# ---------------------------
# Modelo / validação
# ---------------------------
class JobConfig(BaseModel):
    title: str
    placeholders: Dict[str, str] = Field(default_factory=dict)

    @validator("placeholders")
    def upcase_keys(cls, v: Dict[str, str]):
        fixed: Dict[str, str] = {}
        for k, val in v.items():
            kk = k.strip().strip("{} ").upper()
            fixed[kk] = str(val)
        return fixed

# ---------------------------
# UI
# ---------------------------
st.title("Gerador de MOU – usando modelo .docx (sem Google)")
st.caption("Upload do template .docx, preenchimento e download do .docx/.pdf — Calibri 11 em tudo; negrito em tudo exceto as linhas especificadas (PT/EN).")

with st.sidebar:
    st.header("⚙️ Configuração")
    batch_mode = st.toggle("📦 Modo em lote (CSV)", value=False)
    st.markdown("**Modelo (.docx)**")
    template_file = st.file_uploader("Envie o template .docx", type=["docx"])

if template_file is None:
    st.info("Envie o arquivo **.docx** do modelo para começar.")
    st.stop()

# Carrega template em memória
template_bytes = template_file.read()
doc_template = Document(io.BytesIO(template_bytes))
placeholders_found = sorted(list(extract_placeholders(doc_template)))

if not placeholders_found:
    st.warning("Nenhum placeholder no formato {{CHAVE}} foi encontrado no modelo. Ex.: {{GROUP_NAME}}")
    st.stop()

# ---------------------------
# Modo individual
# ---------------------------
if not batch_mode:
    st.subheader("Gerar 1 documento")

    with st.form("single"):
        cols = st.columns(3)
        mapping: Dict[str, str] = {}
        for i, key in enumerate(placeholders_found):
            with cols[i % 3]:
                mapping[key] = st.text_input(key, "")

        default_title = f"MOU – {mapping.get('GROUP_NAME','Sem Nome')} – {datetime.now().strftime('%Y-%m-%d')}"
        title = st.text_input("Título do documento (nome do arquivo)", value=default_title)

        submitted = st.form_submit_button("Gerar arquivos", type="primary")

    if submitted:
        cfg = JobConfig(title=title.strip() or default_title, placeholders=mapping)

        # cria doc, substitui e aplica formatação/exceções
        doc = Document(io.BytesIO(template_bytes))
        exceptions = replace_placeholders_and_collect_exceptions(doc, cfg.placeholders)
        enforce_calibri11_and_bold_with_exceptions(doc, exceptions)

        # salva DOCX em memória
        out_buf = io.BytesIO()
        doc.save(out_buf)
        out_bytes = out_buf.getvalue()

        st.success("Documento gerado!")
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "⬇️ Baixar DOCX",
                data=out_bytes,
                file_name=f"{cfg.title}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        with col2:
            try:
                pdf_bytes = try_export_pdf(out_bytes)
                st.download_button(
                    "⬇️ Baixar PDF",
                    data=pdf_bytes,
                    file_name=f"{cfg.title}.pdf",
                    mime="application/pdf",
                )
            except Exception as e:
                st.info("PDF opcional: instale **Microsoft Word (docx2pdf)** ou **LibreOffice** para habilitar a conversão.")

# ---------------------------
# Modo CSV (lote)
# ---------------------------
else:
    st.subheader("Gerar vários documentos (CSV)")
    st.markdown("O CSV deve ter colunas com os **mesmos nomes** dos placeholders (sem `{{}}`). Ex.: `GROUP_NAME,CNPJ,...`. Opcional: `TITLE`.")

    csv_up = st.file_uploader("Envie o CSV", type=["csv"])
    if csv_up is not None and st.button("Gerar ZIP com .docx/.pdf", type="primary"):
        df = pd.read_csv(csv_up)
        missing_cols = [k for k in placeholders_found if k not in df.columns]
        if missing_cols:
            st.warning("Colunas ausentes no CSV: " + ", ".join(missing_cols))

        zip_mem = io.BytesIO()
        with zipfile.ZipFile(zip_mem, "w", zipfile.ZIP_DEFLATED) as zf:
            for _, row in df.iterrows():
                mapping = {k: str(row.get(k, "")) for k in placeholders_found}
                title = str(row.get("TITLE", f"MOU – {mapping.get('GROUP_NAME','Sem Nome')} – {datetime.now().strftime('%Y-%m-%d')}"))

                doc = Document(io.BytesIO(template_bytes))
                exceptions = replace_placeholders_and_collect_exceptions(doc, mapping)
                enforce_calibri11_and_bold_with_exceptions(doc, exceptions)

                # DOCX
                buf = io.BytesIO()
                doc.save(buf)
                docx_bytes = buf.getvalue()
                zf.writestr(f"{title}.docx", docx_bytes)

                # PDF (best effort)
                try:
                    pdf_bytes = try_export_pdf(docx_bytes)
                    zf.writestr(f"{title}.pdf", pdf_bytes)
                except Exception:
                    # se não conseguir PDF, segue só com DOCX
                    pass

        zip_mem.seek(0)
        st.success("Pacote gerado!")
        st.download_button("⬇️ Baixar todos (.zip)", data=zip_mem, file_name="mous_gerados.zip", mime="application/zip")

# ---------------------------
# Dicas
# ---------------------------
with st.expander("Dicas para template .docx"):
    st.markdown(
        "- Use placeholders **`{{CHAVE}}`** (MAIÚSCULAS). Ex.: `{{GROUP_NAME}}`, `{{CNPJ}}`.\n"
        "- Evite quebrar `{{CHAVE}}` entre linhas/colunas.\n"
        "- Tabelas, cabeçalhos e rodapés são suportados.\n"
        "- **Calibri 11** é aplicado em todo o documento.\n"
        "- **Negrito em tudo**, exceto: frase introdutória; linha do *Business Plan ... {{BP_DATE}}*; "
        "título *2. Especificações e alterações acordadas:*; e o parágrafo de *{{COMMENTS}}* (PT/EN)."
    )
