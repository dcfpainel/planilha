from pathlib import Path
import pandas as pd
import numpy as np

# Mapeamentos de Natureza da Despesa
mapa_naturezas_manual = {
    '335041': 'CONTRIBUIÇÕES',
    '335092': 'DESPESAS DE EXERCÍCIOS ANTERIORES',
    '339014': 'DIÁRIAS - PESSOAL CIVIL',
    '339018': 'AUXÍLIO FINANCEIRO A ESTUDANTES',
    '339020': 'AUXÍLIO FINANCEIRO A PESQUISADORES',
    '339030': 'MATERIAL DE CONSUMO',
    '339031': 'PREMIAÇÕES CULTURAIS, ARTÍSTICAS, CIENTÍFICAS, DESPORTIVAS E OUTROS',
    '339032': 'MATERIAL, BEM OU SERVIÇO PARA DISTRIBUIÇÃO GRATUITA',
    '339033': 'PASSAGENS E DESPESAS COM LOCOMOÇÃO',
    '339035': 'SERVIÇOS DE CONSULTORIA',
    '339036': 'OUTROS SERVIÇOS DE TERCEIROS - PESSOA FÍSICA - DIREITO PRIVADO',
    '339037': 'LOCAÇÃO DE MÃO DE OBRA',
    '339039': 'OUTROS SERVIÇOS DE TERCEIROS - PESSOA JURÍDICA',
    '339040': 'SERVIÇOS DE TECNOLOGIA DA INFORMAÇÃO E COMUNICAÇÃO - PESSOA JURÍDICA',
    '339047': 'OBRIGAÇÕES TRIBUTÁRIAS E CONTRIBUTIVAS',
    '339048': 'OUTROS AUXÍLIOS FINANCEIROS A PESSOAS FÍSICAS',
    '339092': 'DESPESAS DE EXERCÍCIOS ANTERIORES - APLICAÇÕES DIRETAS',
    '339093': 'INDENIZAÇÕES E RESTITUIÇÕES',
    '339139': 'OUTROS SERVIÇOS DE TERCEIROS - PESSOA FÍSICA - DIREITO PÚBLICO',
    '339147': 'OBRIGAÇÕES TRIBUTÁRIAS E CONTRIBUTIVAS EM OPERAÇÕES INTRAORÇAMENTÁRIAS',
    '339193': 'INDENIZAÇÕES E RESTITUIÇÕES EM OPERAÇÕES INTRAORÇAMENTÁRIAS',
    '449051': 'OBRAS E INSTALAÇÕES',
    '449052': 'EQUIPAMENTOS E MATERIAL PERMANENTE',
    '449039': 'SERVIÇOS DE TERCEIROS - PESSOA JURÍDICA (CAPITAL)'
}

mapa_naturezas_curto = {
    '335041': 'CONTRIBUIÇÕES',
    '335092': 'DESP. EXERC. ANTERIORES',
    '339014': 'DIÁRIAS PESSOAL CIVIL',
    '339018': 'AUXÍLIO A ESTUDANTES',
    '339020': 'AUXÍLIO A PESQUISADORES',
    '339030': 'MATERIAL DE CONSUMO',
    '339031': 'PREMIAÇÕES',
    '339032': 'MAT./SERV. DISTRIB. GRATUITA',
    '339033': 'PASSAGENS E LOCOMOÇÃO',
    '339035': 'SERV. CONSULTORIA',
    '339036': 'SERV. TERCEIROS PF PRIVADO',
    '339037': 'LOCAÇÃO DE MÃO DE OBRA',
    '339039': 'SERV. TERCEIROS PJ',
    '339040': 'SERV. TIC PJ',
    '339047': 'OBRIG. TRIBUTÁRIAS',
    '339048': 'AUX. FINANCEIROS PF',
    '339092': 'DESP. EXERC. ANTERIORES DIRETA',
    '339093': 'INDENIZAÇÕES E RESTITUIÇÕES',
    '339139': 'SERV. TERCEIROS PF PÚBLICO',
    '339147': 'OBRIGAÇÕES TRIBUTÁRIAS INTRA',
    '339193': 'INDENIZAÇÕES E RESTITUIÇÕES INTRA',
    '449051': 'OBRAS E INSTALAÇÕES',
    '449052': 'EQUIP. E MATERIAL PERMANENTE',
    '449039': 'SERV. TERCEIROS PJ CAPITAL'
}

def limpa_codigo(val):
    if pd.isna(val):
        return ''
    try:
        val_int = int(float(val))
        s = str(val_int)
    except (ValueError, TypeError):
        s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def limpa_cnpj_cpf(val):
    s = limpa_codigo(val)
    if not s:
        return ''
    if len(s) <= 11:
        return s.zfill(11)
    elif len(s) <= 14:
        return s.zfill(14)
    return s

def tratar_planilha_despesas(caminho_planilha=None, pasta_saida='./planilhas', salvar_arquivos_locais=True):
    """
    Processa a planilha de execução de despesas, aplica regras de negócio contábeis,
    segregação de saldos de RAP e gera métricas padronizadas para o Looker Studio.
    """
    planilhas_path = Path(pasta_saida)
    planilha_original = Path(caminho_planilha) if caminho_planilha else None

    if not planilha_original or not planilha_original.exists():
        candidatos = [
            planilhas_path / 'Execução da despesa.xlsx',
            planilhas_path / 'Execução da despesa-1.xlsx'
        ]
        planilha_original = next((p for p in candidatos if p.exists()), None)

    if not planilha_original or not planilha_original.exists():
        print("Arquivo não encontrado:", planilha_original)
        print("Arquivos na pasta:", [p.name for p in planilhas_path.glob('*')])
        raise FileNotFoundError(f"Planilha de despesas não encontrada em {planilhas_path}")

    print(f"-> Iniciando a leitura do arquivo: {planilha_original}...")

    # 1. Ler SEM cabeçalho e pular as 3 linhas de título (código, nome e rótulo)
    df = pd.read_excel(planilha_original, sheet_name=0, header=None, skiprows=3)
    print(f"-> Dados carregados! {df.shape[0]} linhas e {df.shape[1]} colunas.")

    # 2. Remover a linha "Total" do final da exportação
    df = df[df[0].astype(str).str.strip().str.lower() != 'total'].copy()

    # 3. Manter só as 20 colunas com significado e renomear
    df = df.iloc[:, :20]
    df.columns = [
        'UG_Responsavel_Cod', 'UG_Responsavel_Nome',
        'Favorecido_CNPJ', 'Favorecido_Nome',
        'PI_Cod', 'PI_Nome',
        'Natureza_Despesa_Cod', 'Natureza_Despesa_Nome',
        'NE_CCor_Empenho', 'NE_CCor_Info_Complementar', 'Ano_Emissao_Empenho',
        'Ano_Lancamento',
        'Despesas_Empenhadas',            # item 29
        'Despesas_Liquidadas',            # item 31
        'Despesas_Pagas',                 # item 34
        'RAP_NP_Inscritos',               # item 40
        'RAP_NP_Reinscritos',             # item 41
        'RAP_NP_Liquidados',              # item 44
        'RAP_NP_Pagos',                   # item 46
        'Total_Geral'
    ]

    # 4. Valores nulos viram 0 nas colunas numéricas
    colunas_valores = [
        'Despesas_Empenhadas', 'Despesas_Liquidadas', 'Despesas_Pagas',
        'RAP_NP_Inscritos', 'RAP_NP_Reinscritos', 'RAP_NP_Liquidados',
        'RAP_NP_Pagos', 'Total_Geral'
    ]
    for c in colunas_valores:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    # 5. Regras de negócio

    # 5.1 Com contrato / Sem contrato
    variacoes_contrato = r'contrat|ctn'
    df['Status_Contrato'] = np.where(
        df['PI_Nome'].astype(str).str.contains(variacoes_contrato, case=False, na=False) |
        df['PI_Cod'].astype(str).str.contains(variacoes_contrato, case=False, na=False),
        'Com contrato',
        'Sem contrato'
    )

    # 5.2 Natureza da despesa
    df['Natureza_Despesa_Cod'] = (
        pd.to_numeric(df['Natureza_Despesa_Cod'], errors='coerce')
        .astype('Int64').astype(str).str.strip()
    )

    df['Categoria_Natureza'] = np.where(
        df['Natureza_Despesa_Cod'].str.startswith('3'),
        'DESPESAS CORRENTES',
        'DESPESAS DE CAPITAL'
    )

    df['Natureza_Despesa_Nome'] = (
        df['Natureza_Despesa_Cod'].str[:6].map(mapa_naturezas_manual)
        .fillna(df['Natureza_Despesa_Nome'])
        .astype(str).str.strip().str.upper()
    )

    df['Natureza_Nome_Curto'] = (
        df['Natureza_Despesa_Cod'].str[:6].map(mapa_naturezas_curto)
        .fillna(df['Natureza_Despesa_Nome'])
        .astype(str).str.strip().str.upper()
    )

    # 5.3 Padronização de nomes em maiúsculas (UG Responsável e Favorecido)
    df['UG_Responsavel_Nome'] = df['UG_Responsavel_Nome'].fillna('').astype(str).str.strip().str.upper()
    df['Favorecido_Nome'] = df['Favorecido_Nome'].fillna('').astype(str).str.strip().str.upper()

    # 5.4 MÉTRICAS DA EXECUÇÃO ANUAL
    # RAP não processado "a pagar" = inscritos + reinscritos
    df['RAP_NP_Total'] = df['RAP_NP_Inscritos'] + df['RAP_NP_Reinscritos']

    # (1) Montante = Empenho + Restos a pagar não processados
    df['Montante'] = df['Despesas_Empenhadas'] + df['RAP_NP_Total']

    # (2) Execução (liquidação) = Empenhos liquidados + RAP NP liquidados
    df['Execucao_Liquidacao'] = df['Despesas_Liquidadas'] + df['RAP_NP_Liquidados']

    # Saldo = (1) - (2)  -> "a liquidar"
    df['Saldo_a_Liquidar'] = df['Montante'] - df['Execucao_Liquidacao']

    # RAP Pago (Tudo de RAP Pago)
    df['Total_RAP_Pagos'] = df['RAP_NP_Pagos']

    # (3) Total Geral Pago = Soma de Despesas Pagas + Total RAP Pago
    df['Total_Pago_Geral'] = df['Despesas_Pagas'] + df['Total_RAP_Pagos']

    # 5.5 Anos_RAP: apenas os anos de emissão do empenho anterior ao ano vigente
    df['Ano_Emissao_Empenho'] = pd.to_numeric(df['Ano_Emissao_Empenho'], errors='coerce').astype('Int64')
    df['Ano_Lancamento'] = pd.to_numeric(df['Ano_Lancamento'], errors='coerce').astype('Int64')
    ano_vigente = df['Ano_Lancamento'].max()
    df['Ano_Vigente'] = ano_vigente

    mask = df['Ano_Emissao_Empenho'].notna() & (df['Ano_Emissao_Empenho'] != ano_vigente)
    df['Anos_RAP'] = df['Ano_Emissao_Empenho'].where(mask).astype('Int64')

    # 5.6 Novas Variáveis e Parâmetros
    tolerancia = 0.01

    df['Ano_Empenho_Filtro'] = df['Ano_Emissao_Empenho'].astype('string').fillna('Sem ano')
    df['Origem_Empenho'] = np.where(
        df['Ano_Emissao_Empenho'] == ano_vigente,
        'Exercício atual',
        'Restos a pagar'
    )

    # Saldos a liquidar segregados por origem (evita contaminação do RAP pelo exercício atual)
    df['Saldo_a_Liquidar_RAP'] = np.where(
        df['Origem_Empenho'] == 'Restos a pagar',
        df['RAP_NP_Total'] - df['RAP_NP_Liquidados'],
        0.0
    )
    df['Saldo_a_Liquidar_Exercicio'] = np.where(
        df['Origem_Empenho'] == 'Exercício atual',
        df['Despesas_Empenhadas'] - df['Despesas_Liquidadas'],
        0.0
    )
    df['Saldo_a_Pagar'] = df['Execucao_Liquidacao'] - df['Total_Pago_Geral']
    df['Qtd_Empenho'] = 1
    df['Flag_Inconsistencia'] = np.where(
        (df['Saldo_a_Liquidar'] < -tolerancia) |
        (df['Total_Pago_Geral'] > df['Execucao_Liquidacao'] + tolerancia),
        'Verificar',
        'OK'
    )
    # Data e hora da atualização fixada no Horário de Brasília (UTC-3)
    df['Data_Atualizacao'] = pd.Timestamp.now(tz='America/Sao_Paulo').tz_localize(None).floor('s')

    # 5.7 Colunas de Filtro com Código e Nome (para o Looker Studio)
    df['UG_Responsavel_Filtro'] = (
        df['UG_Responsavel_Cod'].map(limpa_codigo).replace('', 'Sem código') +
        ' - ' + df['UG_Responsavel_Nome'].fillna('').astype(str).str.strip()
    )
    df['Favorecido_Filtro'] = (
        df['Favorecido_CNPJ'].map(limpa_cnpj_cpf).replace('', 'Sem código') +
        ' - ' + df['Favorecido_Nome'].fillna('').astype(str).str.strip()
    )
    df['Natureza_Filtro'] = (
        df['Natureza_Despesa_Cod'].map(limpa_codigo).replace('', 'Sem código') +
        ' - ' + df['Natureza_Despesa_Nome'].fillna('').astype(str).str.strip()
    )

    # 6. Conferência rápida dos totais
    print("\n--- CONFERÊNCIA (R$) ---")
    print(f"Empenhado ............. {df['Despesas_Empenhadas'].sum():>18,.2f}")
    print(f"RAP NP (insc.+reinsc.)  {df['RAP_NP_Total'].sum():>18,.2f}")
    print(f"(1) Montante .......... {df['Montante'].sum():>18,.2f}")
    print(f"(2) Execução/Liquidação {df['Execucao_Liquidacao'].sum():>18,.2f}")
    print(f"    Saldo a liquidar Total ... {df['Saldo_a_Liquidar'].sum():>18,.2f}")
    print(f"      -> Saldo RAP (só diminui): {df['Saldo_a_Liquidar_RAP'].sum():>18,.2f}")
    print(f"      -> Saldo Exercício 2026:   {df['Saldo_a_Liquidar_Exercicio'].sum():>18,.2f}")
    print(f"(3) Pagamento ......... {df['Total_Pago_Geral'].sum():>18,.2f}")
    print(f"    Saldo a pagar ..... {df['Saldo_a_Pagar'].sum():>18,.2f}")
    print(f"Linhas com saldo negativo: {(df['Saldo_a_Liquidar'] < -tolerancia).sum()}")
    print(f"Linhas com pagamento > liquidação: {(df['Total_Pago_Geral'] > df['Execucao_Liquidacao'] + tolerancia).sum()}")
    print(f"Inconsistências registradas (Flag_Inconsistencia = Verificar): {(df['Flag_Inconsistencia'] == 'Verificar').sum()}")

    # 7. Exportar arquivos locais se solicitado
    if salvar_arquivos_locais:
        planilhas_path.mkdir(parents=True, exist_ok=True)
        print(f"\n-> Exportando {df.shape[0]} linhas...")
        df.to_excel(planilhas_path / 'Base_Tratada_Painel_PRAD.xlsx', index=False)
        df.to_csv(planilhas_path / 'Base_Tratada_Painel_PRAD.csv', index=False, sep=';', encoding='utf-8-sig')
        print("-> Arquivos gerados com sucesso em:", planilhas_path.resolve())

    return df

if __name__ == '__main__':
    df = tratar_planilha_despesas()
    print("\n-> Enviando dados tratados para o Google Sheets...")
    try:
        try:
            from src.upload_sheets import upload_para_google_sheets
        except ImportError:
            from upload_sheets import upload_para_google_sheets
        upload_para_google_sheets(df)
    except Exception as e:
        print(f"[AVISO] Não foi possível atualizar o Google Sheets automaticamente: {e}")