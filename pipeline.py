import os
import sys
from pathlib import Path
from src.download_drive import baixar_planilha_drive
from src.tratamento import tratar_planilha_despesas
from src.upload_sheets import upload_para_google_sheets

def executar_pipeline():
    print("=" * 60)
    print(" [1/3] OBTENÇÃO DA PLANILHA (GOOGLE DRIVE / LOCAL)")
    print("=" * 60)

    arquivo = None
    try:
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
    print(" [2/3] TRATAMENTO DE DADOS E REGRAS ORÇAMENTÁRIAS")
    print("=" * 60)
    df_tratado = tratar_planilha_despesas(caminho_planilha=arquivo, salvar_arquivos_locais=True)

    print("\n" + "=" * 60)
    print(" [3/3] SINCRONIZAÇÃO COM O GOOGLE SHEETS")
    print("=" * 60)
    try:
        upload_para_google_sheets(df_tratado)
    except Exception as e:
        print(f"[AVISO] Não foi possível atualizar o Google Sheets: {e}")

    print("\n" + "=" * 60)
    print(" ✓ PIPELINE CONCLUÍDO COM SUCESSO!")
    print("=" * 60)

if __name__ == "__main__":
    executar_pipeline()
