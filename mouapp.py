import io
import os
import json
import pickle
import pandas as pd
import streamlit as st
from datetime import datetime
from typing import Dict, List
from pydantic import BaseModel, Field, validator

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# --------------------------------
# Config Streamlit
# --------------------------------
st.set_page_config(page_title="Gerador de MOU – Jetour (OAuth)", page_icon="🚗", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
]

# --------------------------------
# OAuth do usuário (lê de st.secrets OU de client_secret.json)
# --------------------------------
@st.cache_resource(show_spinner=False)
def get_user_creds():
    creds = None

    # tenta reaproveitar token salvo (apenas na execução atual)
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as f:
            creds = pickle.load(f)

    if not creds or not getattr(creds, "valid", False):
        if creds and getattr(creds, "expired", False) and getattr(creds, "refresh_token", None):
            creds.refresh(Request())
        else:
            # 1) Streamlit Cloud: ler de st.secrets["client_secret"]
            if "client_secret" in st.secrets:
                client_config = json.loads(st.secrets["client_secret"])
                flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
                creds = flow.run_local_server(port=0)
            # 2) Local: ler do arquivo client_secret.json
            elif os.path.exists("client_secret.json"):
                flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", SCOPES)
                creds = flow.run_local_server(port=0)
            else:
                st.error("Credenciais OAuth não encontradas. Adicione `client_secret` em `Secrets` (Cloud) ou coloque `client_secret.json` na mesma pasta do app.")
                st.stop()

        # salva token para reuso durante a sessão
        with open("token.pickle", "wb") as f:
            pickle.dump(creds, f)

    return creds

@st.cache_resource(show_spinner=False)
def get_google_clients():
    creds = get_user_creds()
    docs = build("docs", "v1", credentials=creds, cache_discovery=False)
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    return docs, drive

# --------------------------------
# Modelo de execução
# --------------------------------
class DocRunConfig(BaseModel):
    template_doc_id: str
    output_folder_id: str
    document_title: str
    placeholders: Dict[str, str] = Field(default_factory=dict)

    @validator("placeholders")
    def normalize_keys(cls, v: Dict[str, str]):
        fixed = {}
        for k, val in v.items():
            key = k.strip()
            if not key.startswith("{{"):
                key = "{{" + key
            if not key.endswith("}}"):
                key = key + "}}"
            fixed[key] = str(val)
        return fixed

# --------------------------------
# Funções Drive/Docs
# --------------------------------
def copy_template_to_folder(drive, template_id: str, new_title: str, folder_id: str) -> str:
    file_metadata = {"name": new_title, "parents": [folder_id]}
    copied = drive.files().copy(fileId=template_id, body=file_metadata).execute()
    return copied["id"]

def replace_all_text(docs, document_id: str, mapping: Dict[str, str]):
    requests = []
    for key, value in mapping.items():
        requests.append({
            "replaceAllText": {
                "containsText": {"text": key, "matchCase": True},
                "replaceText": value,
            }
        })
    if requests:
        docs.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()

def export_pdf(drive, document_id: str) -> bytes:
    request = drive.files().export(fileId=document_id, mimeType="application/pdf")
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return fh.read()

def export_docx(drive, document_id: str) -> bytes:
    request = drive.files().export(
        fileId=document_id,
        mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return fh.read()

# --------------------------------
# UI
# --------------------------------
st.title("Gerador de MOU – Jetour (PT/EN) – OAuth")
st.caption("Duplica um template do Google Docs, substitui placeholders e exporta como PDF/DOCX usando a sua conta (OAuth).")

# Inicializa clients (abre login na 1ª vez)
try:
    docs, drive = get_google_clients()
except Exception as e:
    st.error(f"Falha na autenticação OAuth: {e}")
    st.stop()

DEFAULT_KEYS = [
    "FANTASY_NAME",
    "GROUP_NAME",
    "CNPJ",
    "CONTRACT_DATE",
    "FULL_ADDRESS",
    "SHOWROOM_SIZE",
    "START_DATE",
    "END_DATE",
    "INSPECTION_DATE",
    "OPENING_DATE",
    "DEADLINE_DATE",
    "BP_DATE",
    "BP_FILE",
    "COMMENTS",
]

batch_mode = st.toggle("📦 Modo em lote (CSV)", value=False)

# --------------------------------
# Modo individual
# --------------------------------
if not batch_mode:
    st.subheader("Gerar 1 documento")
    with st.form("single_form"):
        template_doc_id = st.text_input("ID do Google Docs TEMPLATE")
        output_folder_id = st.text_input("ID da pasta de destino (Drive)")

        st.divider()
        st.markdown("**Preencha os placeholders**")
        mapping: Dict[str, str] = {}
        cols = st.columns(3)
        for i, key in enumerate(DEFAULT_KEYS):
            with cols[i % 3]:
                mapping[key] = st.text_input(key, "")

        document_title = st.text_input(
            "Título do documento",
            value=f"MOU – {mapping.get('GROUP_NAME','Sem Nome')} – {datetime.now().strftime('%Y-%m-%d')}"
        )

        submitted = st.form_submit_button("Gerar documento", type="primary")

    if submitted:
        try:
            cfg = DocRunConfig(
                template_doc_id=template_doc_id.strip(),
                output_folder_id=output_folder_id.strip(),
                document_title=document_title.strip(),
                placeholders=mapping,
            )

            new_doc_id = copy_template_to_folder(drive, cfg.template_doc_id, cfg.document_title, cfg.output_folder_id)
            replace_all_text(docs, new_doc_id, cfg.placeholders)

            pdf_bytes = export_pdf(drive, new_doc_id)
            docx_bytes = export_docx(drive, new_doc_id)

            st.success("Documento gerado!")
            doc_link = f"https://docs.google.com/document/d/{new_doc_id}/edit"
            st.markdown(f"📄 [Abrir no Google Docs]({doc_link})")

            st.download_button("⬇️ Baixar DOCX", docx_bytes, file_name=f"{cfg.document_title}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.download_button("⬇️ Baixar PDF", pdf_bytes, file_name=f"{cfg.document_title}.pdf", mime="application/pdf")

        except Exception as e:
            st.error(f"Erro ao gerar documento: {e}")

# --------------------------------
# Modo CSV (lote)
# --------------------------------
else:
    st.subheader("Gerar vários documentos via CSV")
    template_doc_id = st.text_input("ID do Google Docs TEMPLATE")
    output_folder_id = st.text_input("ID da pasta de destino (Drive)")
    csv_file = st.file_uploader("CSV de dados", type=["csv"])

    if csv_file is not None and st.button("Gerar documentos em lote", type="primary"):
        df = pd.read_csv(csv_file)
        results: List[Dict[str, str]] = []
        zip_buffer = io.BytesIO()
        import zipfile

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for _, row in df.iterrows():
                placeholders = {k: str(row[k]) for k in df.columns if pd.notna(row[k])}
                title = row.get("TITLE", f"MOU – {placeholders.get('GROUP_NAME','Sem Nome')} – {datetime.now().strftime('%Y-%m-%d')}")
                cfg = DocRunConfig(
                    template_doc_id=template_doc_id.strip(),
                    output_folder_id=output_folder_id.strip(),
                    document_title=str(title),
                    placeholders=placeholders,
                )
                try:
                    new_doc_id = copy_template_to_folder(drive, cfg.template_doc_id, cfg.document_title, cfg.output_folder_id)
                    replace_all_text(docs, new_doc_id, cfg.placeholders)
                    pdf_bytes = export_pdf(drive, new_doc_id)
                    docx_bytes = export_docx(drive, new_doc_id)
                    zf.writestr(f"{cfg.document_title}.pdf", pdf_bytes)
                    zf.writestr(f"{cfg.document_title}.docx", docx_bytes)
                    results.append({"title": cfg.document_title, "status": "OK"})
                except Exception as e:
                    results.append({"title": cfg.document_title, "status": f"ERRO: {e}"})

        st.success("Processo concluído!")
        st.dataframe(results)
        zip_buffer.seek(0)
        st.download_button("⬇️ Baixar todos (.zip)", data=zip_buffer, file_name="mous_gerados.zip", mime="application/zip")
