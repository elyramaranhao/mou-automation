import io
import re
import zipfile
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
    # Parágrafos “soltos”
    for p in doc.paragraphs:
        yield p
    # Parágrafos dentro de tabelas
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
    return txt

def extract_placeholders(doc: Document) -> Set[str]:
    found: Set[str] = set()
    for p in _iter_all_paragraphs(doc):
        text = _para_text(p)
        for m in PLACEHOLDER_RE.finditer(text):
            found.add(m.group(1).strip().upper())
    return found

def replace_placeholders_and_collect_exceptions(doc: Document, mapping: Dict[str, str]):
    """
    Substitui placeholders no documento inteiro.
    Coleta parágrafos-exceção (onde NÃO queremos negrito): os que contêm
    {{BP_DATE}} e {{COMMENTS}} ANTES da substituição.
    Retorna um set com objetos-parágrafo marcados como exceção.
    """
    normalized = {f"{{{{{k}}}}}": str(v) for k, v in mapping.items()}
    exceptions = set()  # parágrafos que NÃO devem ficar em negrito

    for p in _iter_all_paragraphs(doc):
        orig_text = _para_text(p)

        # marca exceções por placeholder antes do replace
        low = orig_text.lower()
        if "{{bp_date}}" in low or "{{comments}}" in low:
            exceptions.add(p)

        # substituição simples no parágrafo “inteiro”
        replaced = orig_text
        for k, v in normalized.items():
            # faz replace case-insensitive dos tokens {{CHAVE}}
            pattern = re.compile(re.escape(k), re.IGNORECASE)
            replaced = pattern.sub(v, replaced)

        if replaced != orig_text:
            # limpar runs e reescrever o parágrafo como 1 run
            for _ in range(len(p.runs)):
                p.runs[0].clear()
                p.runs[0].text = ""
                p._element.remove(p.runs[0]._element)
            p.add_run(replaced)

    return exceptions

def enforce_calibri11_and_bold_with_exceptions(doc: Document, exceptions: Set, extra_exceptions_phrases: List[re.Pattern]):
    """
    1) Define Calibri 11 para todo o documento.
    2) Coloca tudo em negrito por padrão.
    3) Remove negrito nos parágrafos marcados em 'exceptions' e
       nos que casarem com as frases de exceção (PT/EN).
    """
    # Primeiro: padroniza Calibri 11 e bold=True em tudo
    for p in _iter_all_paragraphs(doc):
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(11)
            run.bold = True

    # Depois: desmarca negrito apenas nos parágrafos de exceção
    for p in _iter_all_paragraphs(doc):
        text_norm = (_para_text(p) or "").strip()
        text_low = re.sub(r"\s+", " ", text_norm).lower()

        is_exception = (p in exceptions) or any(pat.search(text_low) for pat in extra_exceptions_phrases)
        if is_exception:
            for run in p.runs:
                run.bold = False

class JobConfig(BaseModel):
    title: str
    placeholders: Dict[str, str] = Field(default_factory=dict)

    @validator("placeholders")
    def upcase_keys(cls, v: Dict[str, str]):
        # Normaliza chaves: aceita com/sem {{}} e converte para UPPER_SNAKE
        fixed: Dict[str, str] = {}
        for k, val in v.items():
            kk = k.strip().strip("{} ").upper()
            fixed[kk] = str(val)
        return fixed

# ---------------------------
# Frases de exceção (PT/EN)
# ---------------------------
# Usamos padrões "flexíveis" para não depender de pontuação exata.
EXCEPTION_PHRASES = [
    # 1) Frase longa de requisitos (PT)
    re.compile(r"^como parte integrante deste documento.*requisitos.*continuidade do processo de nomea", re.IGNORECASE),
    # 1) Frase longa (EN)
    re.compile(r"^as an integral part of this document.*requirements.*continuity of the nomination process", re.IGNORECASE),

    # 2) Linha do Business Plan (PT) – aceita com/sem data substituída
    re.compile(r"^business plan apresentado e validado em", re.IGNORECASE),
    # 2) Linha do Business Plan (EN)
    re.compile(r"^business plan (presented|submitted).*(validated|approved).*on", re.IGNORECASE),

    # 3) Título da seção 2 (PT)
    re.compile(r"^2\.\s*especificaç(ões|oes) e altera(ções|coes) acordadas", re.IGNORECASE),
    # 3) Título da seção 2 (EN)
    re.compile(r"^2\.\s*specifications? and (agreed|approved) (changes|modifications)", re.IGNORECASE),
]

# ---------------------------
# UI
# ---------------------------
st.title("Gerador de MOU – usando modelo .docx (sem Google)")
st.caption("Faça upload do template .docx bi-coluna, preencha placeholders e baixe o .docx final (Calibri 11; negrito em tudo, com exceções).")

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

        submitted = st.form_submit_button("Gerar .docx", type="primary")

    if submitted:
        cfg = JobConfig(title=title.strip() or default_title, placeholders=mapping)

        # Duplica o template em memória e substitui
        doc = Document(io.BytesIO(template_bytes))
        exceptions = replace_placeholders_and_collect_exceptions(doc, cfg.placeholders)
        enforce_calibri11_and_bold_with_exceptions(doc, exceptions, EXCEPTION_PHRASES)

        out_buf = io.BytesIO()
        doc.save(out_buf)
        out_buf.seek(0)

        st.success("Documento gerado!")
        st.download_button(
            "⬇️ Baixar DOCX",
            data=out_buf,
            file_name=f"{cfg.title}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

# ---------------------------
# Modo CSV (lote)
# ---------------------------
else:
    st.subheader("Gerar vários documentos (CSV)")
    st.markdown("O CSV deve ter colunas com os **mesmos nomes** dos placeholders (sem `{{}}`). Ex.: `GROUP_NAME,CNPJ,...`. Opcional: `TITLE`.")

    csv_up = st.file_uploader("Envie o CSV", type=["csv"])
    if csv_up is not None and st.button("Gerar ZIP com .docx", type="primary"):
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
                enforce_calibri11_and_bold_with_exceptions(doc, exceptions, EXCEPTION_PHRASES)

                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                zf.writestr(f"{title}.docx", buf.read())

        zip_mem.seek(0)
        st.success("Pacote gerado!")
        st.download_button("⬇️ Baixar todos (.zip)", data=zip_mem, file_name="mous_gerados.zip", mime="application/zip")

# ---------------------------
# Dicas
# ---------------------------
with st.expander("Dicas para template .docx"):
    st.markdown(
        "- Use placeholders no formato **`{{CHAVE}}`** (MAIÚSCULAS, sem espaços). "
        "Ex.: `{{GROUP_NAME}}`, `{{CNPJ}}`, `{{FULL_ADDRESS}}`.\n"
        "- Evite quebrar `{{CHAVE}}` entre linhas/colunas ou aplicar formatações dentro das chaves.\n"
        "- Tabelas, cabeçalhos e rodapés são suportados.\n"
        "- Todo o texto final é padronizado para **Calibri 11**.\n"
        "- Negrito é aplicado em tudo, **exceto** nas frases e campos definidos como exceção."
    )
