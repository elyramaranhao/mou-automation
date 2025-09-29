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

# -------------------------------------------------
# Config
# -------------------------------------------------
st.set_page_config(
    page_title="Gerador de MOU – (sem Google)",
    page_icon="📝",
    layout="wide"
)

PLACEHOLDER_RE = re.compile(r"\{\{([A-Z0-9_]+)\}\}", re.IGNORECASE)

# -------------------------------------------------
# Utilitários .docx
# -------------------------------------------------
def _iter_all_paragraphs(doc: Document):
    # Parágrafos do corpo
    for p in doc.paragraphs:
        yield p
    # Parágrafos em tabelas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    # Cabeçalho/Rodapé
    for section in doc.sections:
        if section.header:
            for p in section.header.paragraphs:
                yield p
        if section.footer:
            for p in section.footer.paragraphs:
                yield p

def _para_text(p) -> str:
    txt = "".join(run.text for run in p.runs) or p.text or ""
    return re.sub(r"\s+", " ", txt).strip()

def extract_placeholders(doc: Document) -> Set[str]:
    found: Set[str] = set()
    for p in _iter_all_paragraphs(doc):
        for m in PLACEHOLDER_RE.finditer(_para_text(p)):
            found.add(m.group(1).strip().upper())
    return found

def _is_exception_phrase(text_low: str) -> bool:
    """
    Parágrafos que DEVEM ficar sem negrito (PT/EN).
    Busca por palavras-chave em qualquer posição do parágrafo.
    """
    # 1) Frase longa de requisitos (PT/EN)
    if "como parte integrante deste documento" in text_low and "continuidade do processo" in text_low:
        return True
    if "as an integral part of this document" in text_low and "continuity of the nomination process" in text_low:
        return True

    # 2) Linha do Business Plan (PT/EN)
    if "business plan apresentado" in text_low and ("validado" in text_low or "apresentado" in text_low):
        return True
    if "business plan" in text_low and ("validated" in text_low or "approved" in text_low):
        return True

    # 3) Título da seção 2 (PT/EN) – tolerante a quebras
    pt_match = ("especifica" in text_low) and ("acordadas" in text_low or "alterações" in text_low or "alteracoes" in text_low)
    en_match = ("specification" in text_low or "specifications" in text_low) and ("agreed" in text_low or "approved" in text_low) and ("amendment" in text_low or "changes" in text_low or "modifications" in text_low)
    if pt_match or en_match:
        return True

    # 4) Linha que é apenas "N/A"
    if re.fullmatch(r"\s*n\s*/?\s*a\s*\.?\s*", text_low):
        return True

    return False

def replace_placeholders_and_collect_exceptions(doc: Document, mapping: Dict[str, str]):
    """
    1) Marca como exceção (sem negrito) os parágrafos que CONTÊM {{BP_DATE}} e {{COMMENTS}}.
    2) Faz replace dos placeholders no documento inteiro (case-insensitive).
    Retorna set com referências de parágrafos que devem ficar sem negrito.
    """
    normalized = {f"{{{{{k}}}}}": str(v) for k, v in mapping.items()}
    exceptions = set()

    for p in _iter_all_paragraphs(doc):
        orig = _para_text(p)
        low = orig.lower()

        # exceções por placeholder
        if "{{bp_date}}" in low or "{{comments}}" in low:
            exceptions.add(p)

        # replace case-insensitive
        replaced = orig
        for k, v in normalized.items():
            replaced = re.compile(re.escape(k), re.IGNORECASE).sub(v, replaced)

        if replaced != orig:
            # reescreve parágrafo como único run
            for _ in range(len(p.runs)):
                p.runs[0].clear()
                p.runs[0].text = ""
                p._element.remove(p.runs[0]._element)
            p.add_run(replaced)

    return exceptions

def enforce_calibri11_and_bold_with_exceptions(doc: Document, exceptions: Set):
    """
    Aplica Calibri 11 e negrito em tudo; remove negrito:
      - nos parágrafos coletados em 'exceptions' ({{BP_DATE}} e {{COMMENTS}}),
      - nos que casam com frases/padrões de exceção (inclui N/A),
      - e também no "2." quando o título vem quebrado em parágrafos.
    """
    all_paras = list(_iter_all_paragraphs(doc))

    # 1) Calibri 11 + bold=True em todos
    for p in all_paras:
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(11)
            run.bold = True

    def unbold(p):
        for run in p.runs:
            run.bold = False

    # 2) Desmarca bold conforme exceções e vizinhos de "2."
    for i, p in enumerate(all_paras):
        text_low = _para_text(p).lower()

        if (p in exceptions) or _is_exception_phrase(text_low):
            unbold(p)
            continue

        # Se o parágrafo for apenas "2." / "2", desnegrita e checa o seguinte
        if text_low in {"2.", "2"}:
            unbold(p)
            if i + 1 < len(all_paras):
                next_low = _para_text(all_paras[i + 1]).lower()
                if _is_exception_phrase(next_low):
                    unbold(all_paras[i + 1])

# -------------------------------------------------
# PDF (best effort)
# -------------------------------------------------
def try_export_pdf(doc_bytes: bytes) -> bytes:
    """
    Tenta DOCX->PDF com docx2pdf (Word/macOS/Windows).
    Se falhar, tenta LibreOffice (soffice).
    Retorna bytes do PDF ou levanta RuntimeError.
    """
    with tempfile.TemporaryDirectory() as td:
        docx_path = os.path.join(td, "out.docx")
        pdf_path = os.path.join(td, "out.pdf")
        with open(docx_path, "wb") as f:
            f.write(doc_bytes)

        # 1) docx2pdf
        try:
            from docx2pdf import convert
            convert(docx_path, pdf_path)
            with open(pdf_path, "rb") as f:
                return f.read()
        except Exception:
            pass

        # 2) LibreOffice
        try:
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf", "--outdir", td, docx_path],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )
            with open(pdf_path, "rb") as f:
                return f.read()
        except Exception as e:
            raise RuntimeError("Não foi possível gerar PDF (docx2pdf/LibreOffice indisponíveis).") from e

# -------------------------------------------------
# Modelo de dados
# -------------------------------------------------
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

# -------------------------------------------------
# UI
# -------------------------------------------------
st.title("Gerador de MOU – usando modelo .docx (sem Google)")
st.caption("Upload do template .docx, preenchimento e download do .docx/.pdf — Calibri 11 aplicado; negrito em tudo exceto linhas especificadas (PT/EN).")

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

# -------------------------------------------------
# Modo individual
# -------------------------------------------------
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

        # monta doc, substitui e aplica formatação/exceções
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
            except Exception:
                st.info("PDF opcional: instale **Microsoft Word (docx2pdf)** ou **LibreOffice** para habilitar a conversão.")

# -------------------------------------------------
# Modo CSV (lote)
# -------------------------------------------------
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
                    pass

        zip_mem.seek(0)
        st.success("Pacote gerado!")
        st.download_button("⬇️ Baixar todos (.zip)", data=zip_mem, file_name="mous_gerados.zip", mime="application/zip")

# -------------------------------------------------
# Dicas
# -------------------------------------------------
with st.expander("Dicas para template .docx"):
    st.markdown(
        "- Use placeholders **`{{CHAVE}}`** (MAIÚSCULAS). Ex.: `{{GROUP_NAME}}`, `{{CNPJ}}`.\n"
        "- Evite quebrar `{{CHAVE}}` entre linhas/colunas.\n"
        "- Tabelas, cabeçalhos e rodapés são suportados.\n"
        "- **Calibri 11** é aplicado em todo o documento.\n"
        "- **Negrito em tudo**, exceto: frase introdutória; linha do *Business Plan … {{BP_DATE}}*; "
        "título *2. Especificações e alterações acordadas:*; linhas *N/A*; e o parágrafo de *{{COMMENTS}}* (PT/EN)."
    )
