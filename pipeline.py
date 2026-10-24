import os
import sys
from pathlib import Path
from download_email import baixar_planilha_email

def executar_pipeline():
    print("=" * 60)
    print(" [1/3] DOWNLOAD DA PLANILHA DO E-MAIL")
    print("=" * 60)

    # Se as credenciais de e-mail estiverem configuradas, tenta baixar
    if os.getenv("EMAIL_USER") and os.getenv("EMAIL_PASS"):
        arquivo = baixar_planilha_email()
        if not arquivo:
            print("[ERRO] Falha ao baixar anexo do e-mail. Abortando pipeline.")
            sys.exit(1)
    else:
        print("[AVISO] Variáveis de e-mail (EMAIL_USER / EMAIL_PASS) não detectadas.")
        # Verifica se já existe uma planilha local para continuar (ex: testes locais)
        pasta_planilhas = Path("./planilhas")
        planilha_existente = pasta_planilhas / "Execução da despesa.xlsx"
        if planilha_existente.exists():
            print(f"-> Utilizando planilha local existente: {planilha_existente}")
        else:
            print("[ERRO] Nenhuma planilha encontrada em ./planilhas e sem credenciais de e-mail.")
            sys.exit(1)

    print("\n" + "=" * 60)
    print(" [2/3] TRATAMENTO DE DADOS E GERAÇÃO DA BASE")
    print("=" * 60)
    import runpy
    # Executa o script de tratamento existente (que ao final já realiza o upload para o Sheets)
    runpy.run_path("script 1 cestagio.py")

    print("\n" + "=" * 60)
    print(" [3/3] PIPELINE CONCLUÍDO COM SUCESSO!")
    print("=" * 60)

if __name__ == "__main__":
    executar_pipeline()
