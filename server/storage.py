"""
Armazenamento de perfis de clientes no Google Cloud Storage.
"""
import os
import json
import datetime
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
# ── Romaneios (bucket separado) ───────────────────────────────────────────────
GCS_ROMANEIOS_BUCKET = os.environ.get('GCS_ROMANEIOS_BUCKET', 'pata-negra-romaneios')
_gcs_romaneios_client = None


def _romaneios_bucket():
    global _gcs_romaneios_client
    if _gcs_romaneios_client is None:
        _gcs_romaneios_client = gcs.Client()
    return _gcs_romaneios_client.bucket(GCS_ROMANEIOS_BUCKET)


def salvar_romaneio(romaneio_id, dados):
    """Salva um JSON de romaneio (pin do mapa) no bucket de romaneios."""
    blob = _romaneios_bucket().blob(f'{romaneio_id}.json')
    blob.upload_from_string(
        json.dumps(dados, ensure_ascii=False),
        content_type='application/json'
    )


def atualizar_status_romaneio(romaneio_id, status, data=None):
    """Atualiza o status do romaneio SEM apagar (pendente/em_rota/entregue/falhou).
    Preserva todos os dados originais (inclusive a data de inclusao). Registra
    a data da mudanca de status. Retorna True se o romaneio existia."""
    blob = _romaneios_bucket().blob(f'{romaneio_id}.json')
    if not blob.exists():
        return False
    dados = json.loads(blob.download_as_bytes())
    dados['status'] = status
    dados['statusData'] = data or datetime.datetime.utcnow().isoformat()
    blob.upload_from_string(json.dumps(dados, ensure_ascii=False),
                            content_type='application/json')
    return True


def listar_romaneios():
    """Lista todos os romaneios pendentes. Retorna lista de dicts."""
    blobs = _romaneios_bucket().list_blobs()
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


def salvar_pedido_pdf(romaneio_id, pdf_bytes):
    """Salva o PDF da filial associado a um romaneio (mesmo id, extensão .pdf)."""
    blob = _romaneios_bucket().blob(f'{romaneio_id}.pdf')
    blob.upload_from_string(pdf_bytes, content_type='application/pdf')


def carregar_pedido_pdf(romaneio_id):
    """Retorna os bytes do PDF do romaneio, ou None se não existir."""
    blob = _romaneios_bucket().blob(f'{romaneio_id}.pdf')
    if blob.exists():
        return blob.download_as_bytes()
    return None


def deletar_romaneio(romaneio_id):
    """Deleta um romaneio pelo ID (JSON + PDF). Retorna True se o JSON existia, False se não."""
    # apaga o PDF associado se existir (best-effort; soft-delete do bucket cobre recuperação)
    try:
        pdf_blob = _romaneios_bucket().blob(f'{romaneio_id}.pdf')
        if pdf_blob.exists():
            pdf_blob.delete()
    except Exception:
        pass
    blob = _romaneios_bucket().blob(f'{romaneio_id}.json')
    if blob.exists():
        blob.delete()
        return True
    return False


# ── Geofences / zonas (bucket próprio) ────────────────────────────────────────
# Dado DURÁVEL (desenhado uma vez, reutilizado por muito tempo) — fica num bucket
# separado dos romaneios efêmeros, pra poder esvaziar romaneios no console sem
# perder as zonas. Cada zona é {id, nome, cor, geojson}.
GCS_GEOFENCES_BUCKET = os.environ.get('GCS_GEOFENCES_BUCKET', 'pata-negra-geofences')
_gcs_geofences_client = None


def _geofences_bucket():
    global _gcs_geofences_client
    if _gcs_geofences_client is None:
        _gcs_geofences_client = gcs.Client()
    return _gcs_geofences_client.bucket(GCS_GEOFENCES_BUCKET)


def salvar_geofence(geofence_id, dados):
    """Salva (ou sobrescreve) uma zona como JSON no bucket de geofences."""
    blob = _geofences_bucket().blob(f'{geofence_id}.json')
    blob.upload_from_string(
        json.dumps(dados, ensure_ascii=False),
        content_type='application/json'
    )


def listar_geofences():
    """Lista todas as zonas salvas. Retorna lista de dicts."""
    blobs = _geofences_bucket().list_blobs()
    result = []
    for blob in blobs:
        if not blob.name.endswith('.json'):
            continue
        try:
            result.append(json.loads(blob.download_as_bytes()))
        except Exception:
            pass
    return result


def deletar_geofence(geofence_id):
    """Deleta uma zona pelo ID. Retorna True se existia, False se não."""
    blob = _geofences_bucket().blob(f'{geofence_id}.json')
    if blob.exists():
        blob.delete()
        return True
    return False


# ── Entregas salvas (rotas do dia — bucket próprio) ───────────────────────────
# Uma entrega agrupa pedidos numa rota nomeada ("Caminhao do Juca"). Dado
# compartilhado entre dispositivos (por isso no servidor, nao no localStorage).
# {id, nome, criadaEm, status:'em_andamento', pedidoIds:[...]}
GCS_ENTREGAS_BUCKET = os.environ.get('GCS_ENTREGAS_BUCKET', 'pata-negra-entregas')
_gcs_entregas_client = None


def _entregas_bucket():
    global _gcs_entregas_client
    if _gcs_entregas_client is None:
        _gcs_entregas_client = gcs.Client()
    return _gcs_entregas_client.bucket(GCS_ENTREGAS_BUCKET)


def salvar_entrega(entrega_id, dados):
    """Salva (ou sobrescreve) uma entrega como JSON."""
    blob = _entregas_bucket().blob(f'{entrega_id}.json')
    blob.upload_from_string(json.dumps(dados, ensure_ascii=False),
                            content_type='application/json')


def carregar_entrega(entrega_id):
    """Retorna o dict da entrega, ou None se nao existir."""
    blob = _entregas_bucket().blob(f'{entrega_id}.json')
    if not blob.exists():
        return None
    return json.loads(blob.download_as_bytes())


def listar_entregas():
    """Lista todas as entregas salvas. Retorna lista de dicts."""
    result = []
    for blob in _entregas_bucket().list_blobs():
        if not blob.name.endswith('.json'):
            continue
        try:
            result.append(json.loads(blob.download_as_bytes()))
        except Exception:
            pass
    return result


def deletar_entrega(entrega_id):
    """Remove uma entrega. Retorna True se existia."""
    blob = _entregas_bucket().blob(f'{entrega_id}.json')
    if blob.exists():
        blob.delete()
        return True
    return False
