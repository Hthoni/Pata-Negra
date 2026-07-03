"""
Armazenamento de perfis de clientes no Google Cloud Storage.
"""
import os
import json
from google.cloud import storage as gcs

GCS_BUCKET_NAME = os.environ.get('GCS_BUCKET', 'pata-negra-perfis')
_gcs_client = None


def _bucket():
    global _gcs_client
    if _gcs_client is None:
        _gcs_client = gcs.Client()
    return _gcs_client.bucket(GCS_BUCKET_NAME)


def _perfil_blob(cliente):
    return _bucket().blob(f'perfis/{cliente}.xlsx')


def _meta_blob(cliente):
    return _bucket().blob(f'perfis/{cliente}_meta.json')


def perfil_existe(cliente):
    return _perfil_blob(cliente).exists()


def salvar_perfil(cliente, file_bytes, filename=None):
    _perfil_blob(cliente).upload_from_string(
        file_bytes,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    if filename:
        _meta_blob(cliente).upload_from_string(
            json.dumps({'filename': filename}),
            content_type='application/json')


def carregar_perfil_bytes(cliente):
    return _perfil_blob(cliente).download_as_bytes()


def perfil_filename(cliente):
    mb = _meta_blob(cliente)
    if mb.exists():
        return json.loads(mb.download_as_bytes()).get('filename', '')
    return ''


# ── Tabela mestra de produtos (MASTER.xlsx, no bucket de perfis) ──────────────
def _master_blob():
    return _bucket().blob('MASTER.xlsx')


def master_existe():
    return _master_blob().exists()


def salvar_master(file_bytes):
    _master_blob().upload_from_string(
        file_bytes,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


def carregar_master_bytes():
    return _master_blob().download_as_bytes()
