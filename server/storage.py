"""
Armazenamento de perfis de clientes no Google Cloud Storage.
Persistente entre deploys do Cloud Run (que não tem disco persistente).
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
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    if filename:
        _meta_blob(cliente).upload_from_string(
            json.dumps({'filename': filename}),
            content_type='application/json'
        )


def carregar_perfil_bytes(cliente):
    return _perfil_blob(cliente).download_as_bytes()


def perfil_filename(cliente):
    mb = _meta_blob(cliente)
    if mb.exists():
        return json.loads(mb.download_as_bytes()).get('filename', '')
    return ''


# ── Romaneios (bucket separado) ───────────────────────────────────────────────
GCS_ROMANEIOS_BUCKET = os.environ.get('GCS_ROMANEIOS_BUCKET', 'pata-negra-romaneios')
_gcs_romaneios_client = None


def _romaneios_bucket():
    global _gcs_romaneios_client
    if _gcs_romaneios_client is None:
        _gcs_romaneios_client = gcs.Client()
    return _gcs_romaneios_client.bucket(GCS_ROMANEIOS_BUCKET)


def salvar_romaneio(romaneio_id, dados):
    """Salva um JSON de romaneio (pin do mapa) no bucket de romaneios."""    blob = _romaneios_bucket().blob(f'{romaneio_id}.json')
    blob.upload_from_string(
        json.dumps(dados, ensure_ascii=False),
        content_type='application/json'
    )


def listar_romaneios():
    """Lista todos os romaneios pendentes. Retorna lista de dicts."""    blobs = _romaneios_bucket().list_blobs()
    result = []
    for blob in blobs:
        if not blob.name.endswith('.json'):
            continue
        try:
            data = json.loads(blob.download_as_bytes())
            result.append(data)
        except Exception:
            pass
    return result


def deletar_romaneio(romaneio_id):
    """Deleta um romaneio pelo ID. Retorna True se deletado, False se não encontrado."""    blob = _romaneios_bucket().blob(f'{romaneio_id}.json')
    if blob.exists():
        blob.delete()
        return True
    return False
