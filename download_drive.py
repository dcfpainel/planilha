from pathlib import Path
import os
import json
import requests
from google.auth.transport.requests import Request
from upload_sheets import obter_cliente_gspread

# ID da pasta do Google Drive onde o Apps Script salva as planilhas
# Pode ser configurado via variável de ambiente GOOGLE_DRIVE_FOLDER_ID
ID_PASTA_DRIVE_PADRAO = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "")

def obter_credenciais_drive():
    """Obtém credenciais com escopo do Google Drive através da Service Account configurada."""
    gc = obter_cliente_gspread()
    creds = gc.http_client.auth
    if not creds.valid:
        creds.refresh(Request())
    return creds

def baixar_planilha_drive(id_pasta=None, nome_arquivo_busca="Execução da despesa", pasta_destino="./planilhas"):
    """
    Busca o arquivo mais recente no Google Drive (dentro da pasta compartilhada ou acessível
    à Service Account) e faz o download para a pasta local.
    """
    id_pasta = id_pasta or ID_PASTA_DRIVE_PADRAO
    destino = Path(pasta_destino)
    destino.mkdir(parents=True, exist_ok=True)
    caminho_final = destino / "Execução da despesa.xlsx"

    print("-> Autenticando com a Service Account no Google Drive...")
    creds = obter_credenciais_drive()
    token = creds.token
    headers = {"Authorization": f"Bearer {token}"}

    # Monta a query de busca no Google Drive
    queries = [
        "trashed = false",
        f"name contains '{nome_arquivo_busca}'"
    ]
    if id_pasta:
        queries.append(f"'{id_pasta}' in parents")

    q = " and ".join(queries)
    url_busca = f"https://www.googleapis.com/drive/v3/files?q={requests.utils.quote(q)}&orderBy=modifiedTime desc&fields=files(id,name,modifiedTime,size)"

    print(f"-> Buscando arquivo '{nome_arquivo_busca}' no Google Drive...")
    resp = requests.get(url_busca, headers=headers)
    if resp.status_code != 200:
        raise RuntimeError(f"Erro ao consultar arquivos no Google Drive (HTTP {resp.status_code}): {resp.text}")

    arquivos = resp.json().get("files", [])
    if not arquivos:
        # Se não achou na pasta específica, tenta buscar qualquer arquivo com o nome acessível à SA
        if id_pasta:
            print(f"[AVISO] Nenhum arquivo encontrado na pasta {id_pasta}. Buscando globalmente...")
            q_global = f"trashed = false and name contains '{nome_arquivo_busca}'"
            url_global = f"https://www.googleapis.com/drive/v3/files?q={requests.utils.quote(q_global)}&orderBy=modifiedTime desc&fields=files(id,name,modifiedTime,size)"
            resp_global = requests.get(url_global, headers=headers)
            arquivos = resp_global.json().get("files", [])

    if not arquivos:
        print("[ERRO] Nenhum arquivo correspondente encontrado no Google Drive.")
        print("Certifique-se de que a pasta ou arquivo foi compartilhado com o e-mail da Service Account:")
        print(f"  -> {creds.service_account_email}")
        return None

    arquivo_alvo = arquivos[0]
    id_arquivo = arquivo_alvo["id"]
    nome_arquivo = arquivo_alvo["name"]
    modificado = arquivo_alvo.get("modifiedTime", "Desconhecido")
    print(f"-> Arquivo encontrado: '{nome_arquivo}' (ID: {id_arquivo}, Modificado em: {modificado})")

    url_download = f"https://www.googleapis.com/drive/v3/files/{id_arquivo}?alt=media"
    print(f"-> Baixando arquivo para {caminho_final}...")
    resp_dl = requests.get(url_download, headers=headers, stream=True)
    if resp_dl.status_code != 200:
        raise RuntimeError(f"Falha ao baixar arquivo do Google Drive (HTTP {resp_dl.status_code}): {resp_dl.text}")

    with open(caminho_final, "wb") as f:
        for chunk in resp_dl.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)

    print(f"-> Download concluído com sucesso: {caminho_final.resolve()} ({caminho_final.stat().st_size} bytes)")
    return caminho_final

if __name__ == "__main__":
    baixar_planilha_drive()
