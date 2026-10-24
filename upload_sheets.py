from pathlib import Path
import gspread
import pandas as pd

ID_PLANILHA_PADRAO = '1Evnau06wSxWNj8MKVeJCe4qStbksHBvCRQ5NCl6Q7sg'
NOME_ABA_PADRAO = 'Página1'

def obter_credenciais():
    """Localiza o arquivo JSON de credenciais da conta de serviço."""
    candidatos = [
        Path('automacao-prad-f068fa0cf06b.json'),
        Path('credenciais.json')
    ] + list(Path('.').glob('*automacao-prad*.json'))
    for p in candidatos:
        if p.exists():
            return str(p)
    raise FileNotFoundError("Arquivo de credenciais JSON do Google Cloud não foi encontrado na pasta.")

def upload_para_google_sheets(df=None, id_planilha=ID_PLANILHA_PADRAO, nome_aba=NOME_ABA_PADRAO, sincronizar_todas=True):
    """Envia os dados tratados (DataFrame ou arquivo local) para o Google Sheets."""
    if df is None:
        planilhas_path = Path('./planilhas')
        arquivo_base = planilhas_path / 'Base_Tratada_Painel_PRAD.xlsx'
        if not arquivo_base.exists():
            arquivo_base = planilhas_path / 'Base_Tratada_Painel_PRAD.csv'
            if not arquivo_base.exists():
                raise FileNotFoundError("Base tratada não encontrada. Execute o script de tratamento primeiro.")
            df = pd.read_csv(arquivo_base, sep=';', encoding='utf-8-sig')
        else:
            df = pd.read_excel(arquivo_base)

    print(f"-> Conectando ao Google Sheets (ID: {id_planilha[:6]}...{id_planilha[-4:]})...")
    arquivo_cred = obter_credenciais()
    gc = gspread.service_account(filename=arquivo_cred)
    sh = gc.open_by_key(id_planilha)

    # Identifica abas a atualizar
    abas_disponiveis = [ws.title for ws in sh.worksheets()]
    abas_alvo = []
    
    if sincronizar_todas and ('Página1' in abas_disponiveis and 'Sheet1' in abas_disponiveis):
        abas_alvo = ['Página1', 'Sheet1']
    elif nome_aba in abas_disponiveis:
        abas_alvo = [nome_aba]
    else:
        abas_alvo = [sh.sheet1.title]

    df_upload = df.copy()
    # Converte colunas de data/hora para texto legível que o Google Sheets interpreta nativamente
    for col in df_upload.select_dtypes(include=['datetime64', 'datetimetz']).columns:
        df_upload[col] = df_upload[col].dt.strftime('%Y-%m-%d %H:%M:%S')

    # Trata nulos para serialização JSON limpa
    df_upload = df_upload.astype(object).where(pd.notnull(df_upload), '')
    dados = [df_upload.columns.tolist()] + df_upload.values.tolist()
    num_linhas = len(dados)
    num_colunas = len(dados[0]) if dados else 0

    for titulo_aba in abas_alvo:
        ws = sh.worksheet(titulo_aba)
        print(f"-> Atualizando {len(df)} registros na aba '{titulo_aba}'...")
        if ws.row_count < num_linhas or ws.col_count < num_colunas:
            ws.resize(rows=max(num_linhas, ws.row_count), cols=max(num_colunas, ws.col_count))
        ws.clear()
        # raw=False (USER_ENTERED) simula exatamente o ato de colar/digitar no Google Sheets,
        # permitindo que datas e números sejam reconhecidos nativamente pelo Looker Studio.
        ws.update(dados, raw=False)
        print(f"-> Sucesso na aba '{titulo_aba}'!")

    return True

if __name__ == '__main__':
    upload_para_google_sheets()
