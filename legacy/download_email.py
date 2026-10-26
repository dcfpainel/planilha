import imaplib
import email
from email.header import decode_header
from pathlib import Path
import os
import sys

def descompactar_texto(header_text):
    """Decodifica textos de cabeçalho de e-mail (Assunto, Nome de anexo, etc.)."""
    if not header_text:
        return ""
    pedacos = decode_header(header_text)
    resultado = ""
    for texto, encoding in pedacos:
        if isinstance(texto, bytes):
            resultado += texto.decode(encoding or "utf-8", errors="ignore")
        else:
            resultado += str(texto)
    return resultado

def baixar_planilha_email(
    servidor_imap=None,
    usuario=None,
    senha=None,
    filtro_assunto=None,
    pasta_destino=None
):
    """
    Conecta ao servidor IMAP, busca o e-mail mais recente com o assunto
    especificado e baixa o anexo da planilha para a pasta local.
    """
    servidor_imap = servidor_imap or os.getenv("EMAIL_IMAP_SERVER", "imap.gmail.com")
    usuario = usuario or os.getenv("EMAIL_USER")
    senha = senha or os.getenv("EMAIL_PASS")
    filtro_assunto = filtro_assunto or os.getenv("EMAIL_SUBJECT", "Execução da despesa")
    pasta_destino = Path(pasta_destino or os.getenv("DOWNLOAD_DIR", "./planilhas"))

    if not usuario or not senha:
        print("[AVISO] EMAIL_USER e/ou EMAIL_PASS não estão configurados.")
        print("Defina as variáveis de ambiente ou passe como argumentos.")
        return None

    pasta_destino.mkdir(parents=True, exist_ok=True)

    print(f"-> Conectando ao servidor IMAP: {servidor_imap} como '{usuario}'...")
    try:
        mail = imaplib.IMAP4_SSL(servidor_imap)
        mail.login(usuario, senha)
        mail.select("INBOX")
    except Exception as e:
        print(f"[ERRO] Falha ao autenticar no servidor de e-mail: {e}")
        return None

    print(f"-> Buscando e-mails com assunto contendo: '{filtro_assunto}'...")
    try:
        # Busca por assunto
        status, mensagens = mail.search(None, f'(SUBJECT "{filtro_assunto}")')
        ids = mensagens[0].split()

        # Se não encontrar pelo assunto exato, tenta buscar os últimos 10 e-mails recebidos
        if not ids:
            print("-> Assunto específico não encontrou resultados. Buscando nos e-mails mais recentes...")
            status, mensagens = mail.search(None, "ALL")
            todos_ids = mensagens[0].split()
            ids = todos_ids[-10:] if len(todos_ids) >= 10 else todos_ids

        if not ids:
            print("[AVISO] Nenhum e-mail encontrado na caixa de entrada.")
            mail.logout()
            return None

        # Percorre dos mais recentes para os mais antigos
        arquivo_salvo = None
        for msg_id in reversed(ids):
            status, dados = mail.fetch(msg_id, "(RFC822)")
            msg_raw = dados[0][1]
            msg = email.message_from_bytes(msg_raw)
            assunto = descompactar_texto(msg.get("Subject", ""))

            # Verifica se o assunto bate com o filtro (caso tenhamos buscado em ALL)
            if filtro_assunto.lower() not in assunto.lower():
                continue

            print(f"-> E-mail correspondente: '{assunto}'")

            # Varre os anexos da mensagem
            for parte in msg.walk():
                if parte.get_content_disposition() == "attachment":
                    nome_original = parte.get_filename()
                    if nome_original:
                        nome_original = descompactar_texto(nome_original)
                        if nome_original.lower().endswith((".xlsx", ".xls", ".csv")):
                            caminho_final = pasta_destino / "Execução da despesa.xlsx"
                            payload = parte.get_payload(decode=True)
                            with open(caminho_final, "wb") as f:
                                f.write(payload)
                            print(f"-> Anexo '{nome_original}' salvo em: {caminho_final}")
                            arquivo_salvo = caminho_final
                            break

            if arquivo_salvo:
                break

        mail.logout()

        if not arquivo_salvo:
            print(f"[AVISO] Nenhum anexo de planilha encontrado nos e-mails com assunto '{filtro_assunto}'.")
        return arquivo_salvo

    except Exception as e:
        print(f"[ERRO] Ocorreu um erro durante a busca do anexo: {e}")
        try:
            mail.logout()
        except Exception:
            pass
        return None

if __name__ == "__main__":
    resultado = baixar_planilha_email()
    if not resultado:
        sys.exit(1)
