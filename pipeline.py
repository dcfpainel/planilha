import os
import sys
from pathlib import Path
from download_email import baixar_planilha_email

def executar_pipeline():
    print("=" * 60)
    print(" [1/3] OBTENÇÃO DA PLANILHA (GOOGLE DRIVE / LOCAL)")
    print("=" * 60)

    # Tenta baixar do Google Drive via Service Account (Opção B)
    arquivo = None
    try:
        from download_drive import baixar_planilha_drive
        arquivo = baixar_planilha_drive()
    except Exception as e:
        print(f"[AVISO] Tentativa de download do Google Drive não concluída: {e}")

    if not arquivo:
        pasta_planilhas = Path("./planilhas")
        candidatos = [
            pasta_planilhas / "Execução da despesa.xlsx",
            pasta_planilhas / "Execução da despesa-1.xlsx"
        ]
        for p in candidatos:
            if p.exists():
                arquivo = p
                print(f"-> Utilizando planilha local existente: {arquivo}")
                break

    if not arquivo:
        print("[ERRO] Nenhuma planilha encontrada no Google Drive nem em ./planilhas.")
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
