/**
 * ==============================================================================
 * GOOGLE APPS SCRIPT: SALVAR PLANILHA DO GMAIL NO GOOGLE DRIVE
 * ==============================================================================
 * 
 * Como usar:
 * 1. Acesse https://script.google.com com o e-mail que recebe as planilhas.
 * 2. Clique em "Novo projeto".
 * 3. Cole este código no editor (substitua o código padrão).
 * 4. Ajuste a constante NOME_OU_ID_PASTA com o ID ou nome da sua pasta no Google Drive.
 * 5. Clique em "Salvar" (ícone de disquete) e teste executando a função salvarPlanilhaEmailNoDrive.
 * 6. Na primeira vez, o Google pedirá autorização para ler o Gmail e salvar no Drive.
 * 7. Para automatizar: clique no ícone de Relógio (Acionadores) à esquerda:
 *    - Adicionar acionador -> Escolha a função: salvarPlanilhaEmailNoDrive
 *    - Origem do evento: Baseado no tempo
 *    - Tipo: Temporizador por dia (ex: entre 07:00 e 08:00).
 * ==============================================================================
 */

// ID da pasta Planilhas_PRAD no Google Drive
const PASTA_DRIVE_ID = ""; 
const NOME_PASTA_PADRAO = "Planilhas_PRAD"; 
const FILTRO_ASSUNTO = "Execução da despesa";
const ETIQUETA_PROCESSADO = "Processado_Drive";

/**
 * Obtém a pasta de destino no Google Drive pelo ID ou pelo Nome.
 */
function obterPastaDrive() {
  let pasta = null;
  if (PASTA_DRIVE_ID && PASTA_DRIVE_ID.trim() !== "") {
    try {
      pasta = DriveApp.getFolderById(PASTA_DRIVE_ID.trim());
    } catch (e) {
      console.warn("-> ID de pasta não acessível via getFolderById. Tentando por nome...");
    }
  }
  if (!pasta) {
    const pastas = DriveApp.getFoldersByName(NOME_PASTA_PADRAO);
    pasta = pastas.hasNext() ? pastas.next() : DriveApp.createFolder(NOME_PASTA_PADRAO);
  }
  return pasta;
}

/**
 * Função principal (executada pelo acionador diário).
 * Processa novos e-mails pendentes, salva APENAS a versão mais recente no Drive
 * e etiqueta todos os pendentes para evitar repetições.
 */
function salvarPlanilhaEmailNoDrive() {
  console.log("-> Iniciando busca por e-mails com a planilha...");
  
  const pasta = obterPastaDrive();
  console.log("-> Pasta no Google Drive: " + pasta.getName() + " (ID: " + pasta.getId() + ")");
  
  let etiqueta = GmailApp.getUserLabelByName(ETIQUETA_PROCESSADO);
  if (!etiqueta) {
    etiqueta = GmailApp.createLabel(ETIQUETA_PROCESSADO);
  }
  
  // Busca threads com anexo .xlsx, assunto correspondente e SEM a etiqueta
  const query = 'has:attachment filename:xlsx subject:("' + FILTRO_ASSUNTO + '") -label:' + ETIQUETA_PROCESSADO;
  console.log("-> Query de busca: " + query);
  const threads = GmailApp.search(query, 0, 10);
  
  if (threads.length === 0) {
    console.log("-> Nenhum novo e-mail pendente encontrado.");
    return;
  }
  
  console.log("-> Encontrada(s) " + threads.length + " thread(s) pendente(s).");
  
  // O GmailApp.search retorna do MAIS RECENTE (índice 0) para o mais antigo.
  // Localizamos o anexo mais recente na thread mais recente:
  let anexoMaisRecente = null;
  let dataAnexoMaisRecente = null;
  
  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const mensagens = thread.getMessages();
    
    // Percorre da mensagem mais recente para a mais antiga na thread
    for (let m = mensagens.length - 1; m >= 0; m--) {
      const msg = mensagens[m];
      const anexos = msg.getAttachments();
      
      for (let a = 0; a < anexos.length; a++) {
        const anexo = anexos[a];
        if (anexo.getName().toLowerCase().endsWith(".xlsx")) {
          anexoMaisRecente = anexo;
          dataAnexoMaisRecente = msg.getDate();
          break;
        }
      }
      if (anexoMaisRecente) break;
    }
    if (anexoMaisRecente) break; // Garante que pegamos apenas o mais recente absoluto
  }
  
  if (anexoMaisRecente) {
    console.log("-> Salvando o anexo mais recente da mensagem de: " + dataAnexoMaisRecente);
    const nomeFinal = "Execução da despesa.xlsx";
    
    // Remove versões anteriores da pasta para manter sempre a mais atual
    const arquivosAntigos = pasta.getFilesByName(nomeFinal);
    while (arquivosAntigos.hasNext()) {
      const arqAntigo = arquivosAntigos.next();
      console.log("-> Removendo versão anterior da pasta: " + arqAntigo.getName());
      arqAntigo.setTrashed(true);
    }
    
    pasta.createFile(anexoMaisRecente.copyBlob().setName(nomeFinal));
    console.log("-> Sucesso: Arquivo atualizado no Drive como '" + nomeFinal + "'!");
  } else {
    console.log("-> Nenhum anexo .xlsx válido encontrado.");
  }
  
  // Marca TODAS as threads encontradas com a etiqueta para não acumularem
  for (let i = 0; i < threads.length; i++) {
    threads[i].addLabel(etiqueta);
  }
  console.log("-> " + threads.length + " thread(s) marcada(s) com a etiqueta '" + ETIQUETA_PROCESSADO + "'.");
}

/**
 * Função utilitária para recuperação / reprocessamento imediato.
 * Ignora a etiqueta "Processado_Drive" e busca diretamente o e-mail mais recente recebido,
 * salvando sua planilha no Drive.
 */
function forcarReprocessamentoMaisRecente() {
  console.log("-> [FORÇAR] Buscando e-mail mais recente com o anexo (mesmo se já etiquetado)...");
  
  const pasta = obterPastaDrive();
  console.log("-> Pasta no Google Drive: " + pasta.getName() + " (ID: " + pasta.getId() + ")");
  
  const query = 'has:attachment filename:xlsx subject:("' + FILTRO_ASSUNTO + '")';
  const threads = GmailApp.search(query, 0, 1);
  
  if (threads.length === 0) {
    console.log("-> Nenhum e-mail encontrado com a query: " + query);
    return;
  }
  
  const thread = threads[0];
  const mensagens = thread.getMessages();
  let anexoMaisRecente = null;
  let dataMsg = null;
  
  for (let m = mensagens.length - 1; m >= 0; m--) {
    const msg = mensagens[m];
    const anexos = msg.getAttachments();
    for (let a = 0; a < anexos.length; a++) {
      if (anexos[a].getName().toLowerCase().endsWith(".xlsx")) {
        anexoMaisRecente = anexos[a];
        dataMsg = msg.getDate();
        break;
      }
    }
    if (anexoMaisRecente) break;
  }
  
  if (anexoMaisRecente) {
    console.log("-> [FORÇAR] Processando anexo da mensagem de: " + dataMsg);
    const nomeFinal = "Execução da despesa.xlsx";
    
    const arquivosAntigos = pasta.getFilesByName(nomeFinal);
    while (arquivosAntigos.hasNext()) {
      const arqAntigo = arquivosAntigos.next();
      console.log("-> Removendo versão anterior da pasta: " + arqAntigo.getName());
      arqAntigo.setTrashed(true);
    }
    
    pasta.createFile(anexoMaisRecente.copyBlob().setName(nomeFinal));
    console.log("-> [FORÇAR] Sucesso: Planilha mais recente (" + dataMsg + ") salva como '" + nomeFinal + "'!");
    
    let etiqueta = GmailApp.getUserLabelByName(ETIQUETA_PROCESSADO);
    if (!etiqueta) {
      etiqueta = GmailApp.createLabel(ETIQUETA_PROCESSADO);
    }
    thread.addLabel(etiqueta);
  } else {
    console.log("-> [FORÇAR] Nenhum anexo .xlsx encontrado na mensagem mais recente.");
  }
}