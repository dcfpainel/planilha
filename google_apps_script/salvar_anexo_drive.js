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

// Se colocar o ID da pasta do Google Drive (da URL: drive.google.com/drive/folders/ID_AQUI), é mais direto.
// Se deixar apenas o nome da pasta, o script criará a pasta se ela não existir.
const PASTA_DRIVE_ID = ""; // Ex: "1AbCdEfGhIjKlMnOpQrStUvWxYz" (opcional se definir NOME_PASTA)
const NOME_PASTA_PADRAO = "Planilhas_PRAD"; 
const FILTRO_ASSUNTO = "Execução da despesa";
const ETIQUETA_PROCESSADO = "Processado_Drive";

function salvarPlanilhaEmailNoDrive() {
  console.log("-> Iniciando busca por e-mails com a planilha...");
  
  // 1. Localiza ou cria a pasta no Google Drive
  let pasta;
  if (PASTA_DRIVE_ID && PASTA_DRIVE_ID.trim() !== "") {
    pasta = DriveApp.getFolderById(PASTA_DRIVE_ID.trim());
  } else {
    const pastas = DriveApp.getFoldersByName(NOME_PASTA_PADRAO);
    pasta = pastas.hasNext() ? pastas.next() : DriveApp.createFolder(NOME_PASTA_PADRAO);
  }
  console.log("-> Pasta no Google Drive: " + pasta.getName() + " (ID: " + pasta.getId() + ")");
  
  // 2. Garante que a etiqueta de controle exista no Gmail
  let etiqueta = GmailApp.getUserLabelByName(ETIQUETA_PROCESSADO);
  if (!etiqueta) {
    etiqueta = GmailApp.createLabel(ETIQUETA_PROCESSADO);
  }
  
  // 3. Busca threads que tenham anexo .xlsx, assunto correspondente e ainda NÃO tenham a etiqueta
  const query = 'has:attachment filename:xlsx subject:("' + FILTRO_ASSUNTO + '") -label:' + ETIQUETA_PROCESSADO;
  console.log("-> Query de busca: " + query);
  const threads = GmailApp.search(query, 0, 5);
  
  if (threads.length === 0) {
    console.log("-> Nenhum novo e-mail pendente encontrado.");
    return;
  }
  
  console.log("-> Encontrada(s) " + threads.length + " mensagem(ns) pendente(s).");
  
  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const mensagens = thread.getMessages();
    
    // Pega a mensagem mais recente da thread
    for (let m = mensagens.length - 1; m >= 0; m--) {
      const msg = mensagens[m];
      const anexos = msg.getAttachments();
      
      for (let a = 0; a < anexos.length; a++) {
        const anexo = anexos[a];
        const nomeAnexo = anexo.getName();
        
        if (nomeAnexo.toLowerCase().endsWith(".xlsx")) {
          console.log("-> Processando anexo: " + nomeAnexo + " da mensagem de " + msg.getDate());
          
          // Salva ou sobrescreve o arquivo com o nome padronizado
          const nomeFinal = "Execução da despesa.xlsx";
          
          // Se já existir arquivo com esse nome na pasta, remove para manter sempre o mais atual
          const arquivosAntigos = pasta.getFilesByName(nomeFinal);
          while (arquivosAntigos.hasNext()) {
            const arqAntigo = arquivosAntigos.next();
            console.log("-> Removendo versão anterior da pasta: " + arqAntigo.getName());
            arqAntigo.setTrashed(true);
          }
          
          // Cria o novo arquivo na pasta
          pasta.createFile(anexo.copyBlob().setName(nomeFinal));
          console.log("-> Sucesso: Arquivo salvo na pasta do Drive como '" + nomeFinal + "'!");
          break;
        }
      }
    }
    
    // Marca a thread com a etiqueta para não processar novamente amanhã
    thread.addLabel(etiqueta);
  }
}
