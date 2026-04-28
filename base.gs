const NOME_ABA_DADOS = "TOTAL"; 



// ==========================================
// ROTEADOR PÚBLICO (ERP vs PORTAL DE ACEITE)
// ==========================================
function doGet(e) {
  const pagina = e?.parameter?.p;
  const hash = e?.parameter?.h;

  // ROTA: PORTAL DE ACEITE DO PACIENTE
  if (pagina === 'aceite' && hash) {
    const template = HtmlService.createTemplateFromFile('PortalAceite');
    template.hash = hash;
    return template.evaluate()
      .setTitle('Assinatura Eletrónica - Clínica Ramos')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ROTA: ERP PADRÃO (O seu código original continua aqui)
  const menu = e?.parameter?.page || 'ModuloERP';
  const template = HtmlService.createTemplateFromFile(menu);
  return template.evaluate()
    .setTitle('Clínica Ramos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// REGISTRO DE ASSINATURA DIGITAL (PORTAL -> PLANILHA)
// ==========================================
// ============================================================
// REGISTRO DE ASSINATURA DIGITAL - VERSÃO FINAL PARA COLUNA G
// ============================================================
function apiRegistarAceiteDigital(hash) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaDocs = ss.getSheetByName("DOCUMENTOS");
    if (!abaDocs) throw new Error("Aba DOCUMENTOS não encontrada!");

    const dados = abaDocs.getDataRange().getValues();
    const hashBusca = String(hash).trim();
    let linhaAlvo = -1;
    let nomePac = "";

    // Percorre a planilha procurando o hash na coluna F (Índice 5)
    for (let i = 1; i < dados.length; i++) {
      // Testamos a coluna F (5) e a coluna D (3) por segurança
      if (String(dados[i][5]).trim() === hashBusca || String(dados[i][3]).includes(hashBusca)) {
        linhaAlvo = i + 1;
        nomePac = dados[i][1];
        break;
      }
    }

    if (linhaAlvo === -1) throw new Error("Chave " + hash + " não encontrada na planilha.");

    // FORÇA A GRAVAÇÃO NA COLUNA G (7)
    // Usamos o número 7 diretamente para bater com o seu print
    abaDocs.getRange(linhaAlvo, 7).setValue("✅ ASSINADO");

    // Registra na aba de auditoria para você conferir
    let abaLog = ss.getSheetByName("ACEITES_DIGITAIS") || ss.insertSheet("ACEITES_DIGITAIS");
    abaLog.appendRow([new Date(), nomePac, hash, "OK"]);

    SpreadsheetApp.flush(); 
    return { success: true, paciente: nomePac };

  } catch (e) {
    throw new Error("Erro no servidor: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// SALVAR FOTO E LOCALIZAÇÃO DA ASSINATURA
// ==========================================
function apiProcessarAssinaturaAvancada(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDocs = ss.getSheetByName("DOCUMENTOS");
  const dadosDocs = abaDocs.getDataRange().getValues();
  
  let linhaAlvo = -1;
  let nomePac = "";

  // 1. Localiza a linha do documento pelo Hash
  for (let i = 1; i < dadosDocs.length; i++) {
    if (String(dadosDocs[i][5]).trim() === String(dados.hash).trim()) {
      linhaAlvo = i + 1;
      nomePac = dadosDocs[i][1];
      break;
    }
  }

  if (linhaAlvo === -1) throw new Error("Documento não localizado.");

  // 2. Localização das Pastas no Drive
  const pastaRaiz = DriveApp.getFoldersByName("Documentos_Clinica_Ramos").next();
  const pastaPaciente = pastaRaiz.getFoldersByName(nomePac).next();
  
  let pastasAssinaturas = pastaPaciente.getFoldersByName("Assinaturas");
  let pastaAssinaturas = pastasAssinaturas.hasNext() ? pastasAssinaturas.next() : pastaPaciente.createFolder("Assinaturas");

  // 3. Converte a Base64 da foto em arquivo real
  const imagemBlob = Utilities.newBlob(Utilities.base64Decode(dados.foto.split(",")[1]), "image/png", "FOTO_ASSINATURA_" + dados.hash + ".png");
  const arquivoFoto = pastaAssinaturas.createFile(imagemBlob);
  const urlFoto = arquivoFoto.getUrl();

  // 4. Atualiza a Planilha com Status, Localização e Link da Foto
  abaDocs.getRange(linhaAlvo, 7).setValue("✅ ASSINADO"); // Coluna G: Status
  
  // Registra no Log de Auditoria com as Coordenadas (GPS)
  let abaLog = ss.getSheetByName("ACEITES_DIGITAIS") || ss.insertSheet("ACEITES_DIGITAIS");
  abaLog.appendRow([
    Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss"),
    nomePac,
    dados.hash,
    "ASSINADO COM FOTO",
    urlFoto,
    dados.lat + ", " + dados.lng // Localização exata
  ]);

  return { success: true, paciente: nomePac };
}

function getScriptURL() { return ScriptApp.getService().getUrl(); }

function abrirERP() {
  var html = HtmlService.createTemplateFromFile('ModuloERP').evaluate().setWidth(1350).setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'ERP Clínica Ramos'); 
}

function showDashboard() {
  var html = HtmlService.createTemplateFromFile('Index').evaluate().setWidth(1200).setHeight(850);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Painel Gerencial'); 
}

function abrirJanelaResultado() {
  var html = HtmlService.createTemplateFromFile('ResultadoClinica').evaluate().setWidth(1100).setHeight(850);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Resultado');
}

function incluirNoERP(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return "";
  }
}

function getListasParaForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = ss.getSheetByName("config"); 
  if (!sheetConfig) return { doutores: [], procedimentos: [], planos: [], pagamentos: [] };
  var data = sheetConfig.getDataRange().getValues();
  var dentistas = [], procedimentos = [], planos = [], pagamentos = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][10]) dentistas.push(data[i][10].toString().trim().toUpperCase());
    if (data[i][0]) procedimentos.push(data[i][0].toString().trim());
    if (data[i][2]) planos.push(data[i][2].toString().trim());
    if (data[i][8]) pagamentos.push(data[i][8].toString().trim());
  }
  return {
    doutores: [...new Set(dentistas)].sort(),
    procedimentos: [...new Set(procedimentos)].sort(),
    planos: [...new Set(planos)].sort(),
    pagamentos: [...new Set(pagamentos)].sort()
  };
}

function getDashboardData() {
  const dados = getDadosTabelaERP();
  return dados.map(item => ({
    ...item,
    mes: item.data ? item.data.split('/')[1] + '/' + item.data.split('/')[2] : '',
    total: (item.particular || 0) + (item.vplano || 0)
  }));
}

// ===== LEITURA DOS DADOS =====
function getDadosTabelaERP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOME_ABA_DADOS);
  if (!sheet) return [];
  
  const values = sheet.getDataRange().getValues();
  let resultados = [];
  
  for (let i = 2; i < values.length; i++) {
    let r = values[i];
    if (!r[1] || String(r[1]).trim() === "") continue; 
    
    let dtExibicao = "";
    if (r[0] instanceof Date) {
      const data = r[0];
      dtExibicao = `${String(data.getDate()).padStart(2,'0')}/${String(data.getMonth()+1).padStart(2,'0')}/${data.getFullYear()}`;
    } else {
      dtExibicao = r[0];
    }
    
    resultados.push({
      id: r[13] || (i + 1),
      data: dtExibicao,
      doutor: r[1],
      paciente: r[2],
      procedimento: r[3],
      plano: r[4], 
      pagamento: r[5],
      obs: r[6],
      particular: Number(r[7]) || 0,
      vplano: Number(r[8]) || 0,
      vlab: Number(r[9]) || 0,
      lucroClinica: Number(r[12]) || 0,
      parteDentista: (Number(r[10]) || 0) + (Number(r[11]) || 0)
    });
  }
  
  return resultados;
}

function encontrarPrimeiraLinhaVazia(sheet) {
  const ultimaLinha = sheet.getLastRow();
  for (let i = 3; i <= ultimaLinha + 5; i++) {
    const valorColunaB = sheet.getRange(i, 2).getValue();
    if (!valorColunaB || valorColunaB.toString().trim() === "") {
      return i;
    }
  }
  return ultimaLinha + 1;
}

function copiarFormatacaoEFormulas(sheet, linhaOrigem, linhaDestino) {
  try {
    for (let col = 11; col <= 13; col++) {
      const formula = sheet.getRange(linhaOrigem, col).getFormula();
      if (formula) {
        sheet.getRange(linhaDestino, col).setFormula(formula);
      }
    }
    for (let col = 8; col <= 10; col++) {
      const numeroFormatado = sheet.getRange(linhaOrigem, col).getNumberFormat();
      if (numeroFormatado) {
        sheet.getRange(linhaDestino, col).setNumberFormat(numeroFormatado);
      }
    }
    const dataFormatada = sheet.getRange(linhaOrigem, 1).getNumberFormat();
    if (dataFormatada) {
      sheet.getRange(linhaDestino, 1).setNumberFormat(dataFormatada);
    }
  } catch (e) {
    console.log("Erro ao copiar formatação: " + e);
  }
}

// ===== SALVAR LANÇAMENTO UNIFICADO =====
function salvarLancamentoUnificado(dados) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(10000);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Usa a variável global NOME_ABA_DADOS (ou "DADOS" como segurança)
    const sheet = ss.getSheetByName(typeof NOME_ABA_DADOS !== 'undefined' ? NOME_ABA_DADOS : "DADOS");
    if (!sheet) return "ERRO: Aba não encontrada";
    
    let dataObj = new Date();
    if (dados.data && dados.data.includes('/')) {
      const partes = dados.data.split("/");
      dataObj = new Date(partes[2], partes[1] - 1, partes[0], 12, 0, 0);
    }

    // LIMPADOR INTELIGENTE: Transforma "1.500,00" em 1500.00 perfeito pro Google Sheets
    const limparValor = (v) => Number((v || "0").toString().replace(/\./g, '').replace(',', '.')) || 0;

    const valorParticular = limparValor(dados.particular);
    const valorPlano = limparValor(dados.vplano);
    const valorLab = limparValor(dados.vlab);
    
    // CAPTURA A TAXA DE CARTÃO DO PACOTE DE DADOS
    const valorTaxaCartao = limparValor(dados.taxaCartao);
    
    // ==========================================
    // EDIÇÃO DE REGISTRO EXISTENTE
    // ==========================================
    if (dados.idLinha && dados.idLinha !== "") {
      const busca = sheet.getRange("N:N").createTextFinder(dados.idLinha).matchEntireCell(true).findNext();
      
      if (busca) {
        const linha = busca.getRow();
        const valoresAtualizados = [[
          dataObj, dados.doutor, dados.paciente, dados.procedimento,
          dados.plano, dados.pagamento, dados.obs, valorParticular,
          valorPlano, valorLab
        ]];
        
        sheet.getRange(linha, 1, 1, 10).setValues(valoresAtualizados);
        
        // GRAVA A TAXA NA COLUNA "O" (15) DURANTE A EDIÇÃO
        sheet.getRange(linha, 15).setValue(valorTaxaCartao);
        
        return "Registro atualizado com sucesso!";
      }
      return "Registro não encontrado para edição";
    }
    
    // ==========================================
    // CRIANDO UM NOVO REGISTRO
    // ==========================================
    const ultimaLinha = sheet.getLastRow();
    let linhaDestino = ultimaLinha + 1;
    const inicioBusca = Math.max(3, ultimaLinha - 100);
    
    // Acha a primeira linha vazia
    const valoresColunaB = sheet.getRange(inicioBusca, 2, ultimaLinha - inicioBusca + 1, 1).getValues();
    for (let i = 0; i < valoresColunaB.length; i++) {
      if (!valoresColunaB[i][0] || valoresColunaB[i][0].toString().trim() === "") {
        linhaDestino = inicioBusca + i;
        break;
      }
    }
    
    const novoId = new Date().getTime(); 
    
    const novaLinhaDados = [[
      dataObj, dados.doutor, dados.paciente, dados.procedimento,
      dados.plano, dados.pagamento, dados.obs, valorParticular,
      valorPlano, valorLab
    ]];
    
    // Copia as Fórmulas das Colunas K(11), L(12) e M(13)
    if (sheet.getLastRow() >= 3) {
      const linhaTemplate = 3;
      
      let formulaR1C1_K = sheet.getRange(linhaTemplate, 11).getFormulaR1C1();
      if (formulaR1C1_K) sheet.getRange(linhaDestino, 11).setFormulaR1C1(formulaR1C1_K);
      else sheet.getRange(linhaDestino, 11).setValue(sheet.getRange(linhaTemplate, 11).getValue());
      
      let formulaR1C1_L = sheet.getRange(linhaTemplate, 12).getFormulaR1C1();
      if (formulaR1C1_L) sheet.getRange(linhaDestino, 12).setFormulaR1C1(formulaR1C1_L);
      else sheet.getRange(linhaDestino, 12).setValue(sheet.getRange(linhaTemplate, 12).getValue());
      
      let formulaR1C1_M = sheet.getRange(linhaTemplate, 13).getFormulaR1C1();
      if (formulaR1C1_M) sheet.getRange(linhaDestino, 13).setFormulaR1C1(formulaR1C1_M);
      else sheet.getRange(linhaDestino, 13).setValue(sheet.getRange(linhaTemplate, 13).getValue());
    }
    
    // Grava as 10 primeiras colunas (A até J)
    sheet.getRange(linhaDestino, 1, 1, 10).setValues(novaLinhaDados);
    
    // Grava o ID único na Coluna N (14)
    sheet.getRange(linhaDestino, 14).setValue(novoId);
    
    // GRAVA A TAXA DO CARTÃO NA COLUNA "O" (15) DO NOVO REGISTRO
    sheet.getRange(linhaDestino, 15).setValue(valorTaxaCartao);
    
    return "Registro salvo com sucesso!";
    
  } catch (erro) {
    return "Erro ao salvar: " + erro.toString();
  } finally {
    lock.releaseLock();
  }
}

function salvarLancamento(dados) {
  return salvarLancamentoUnificado(dados);
}

// ===== EXCLUIR REGISTRO (VERSÃO OTIMIZADA) =====
function excluirLancamento(idLinha) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Procura a aba principal
    const sheet = ss.getSheetByName(typeof NOME_ABA_DADOS !== 'undefined' ? NOME_ABA_DADOS : "DADOS");
    
    // Prepara a aba de LOG
    let sheetLog = ss.getSheetByName("LOG_EXCLUSOES");
    if (!sheetLog) {
      sheetLog = ss.insertSheet("LOG_EXCLUSOES");
    }

    // BUSCA RÁPIDA: Encontra o ID diretamente na Coluna N (14) sem ler a tabela toda
    const busca = sheet.getRange("N:N").createTextFinder(idLinha).matchEntireCell(true).findNext();
    
    if (busca) {
      const linha = busca.getRow();
      
      // 1. FAZ O BACKUP DE SEGURANÇA (Copia da Coluna A até à Coluna O)
      const dadosLinha = sheet.getRange(linha, 1, 1, 15).getValues()[0];
      sheetLog.appendRow(dadosLinha);
      
      // 2. LIMPEZA CIRÚRGICA (Preserva as Fórmulas K, L, M)
      // Limpa os dados principais (Colunas A até J)
      sheet.getRange(linha, 1, 1, 10).clearContent();
      
      // Limpa o ID da linha (Coluna N / 14)
      sheet.getRange(linha, 14).clearContent();
      
      // Limpa a Taxa do Cartão (Coluna O / 15)
      sheet.getRange(linha, 15).clearContent();
      
      return "Registo excluído com segurança e guardado no LOG!";
    }
    
    return "Registo não encontrado.";
    
  } catch (erro) {
    return "Erro ao excluir: " + erro.toString();
  } finally {
    lock.releaseLock();
  }
}


function getListaPacientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOME_ABA_DADOS);
  if (!sheet) return [];
  const nomes = sheet.getRange(3, 3, sheet.getLastRow() - 2, 1).getValues().flat();
  return [...new Set(nomes)].filter(String).sort();
}

function getAnalisePerformance(mesAno) {
  const dados = getDadosTabelaERP();
  const filtrados = dados.filter(d => {
    if (!d.data) return false;
    const m = d.data.split('/')[1] + '/' + d.data.split('/')[2];
    return m === mesAno;
  });
  
  let lucroD = {};              
  let atendimentosD = {};        
  let pacientesUnicosD = {};     
  let contagemP = {};            
  let rentabilidadeP = {};       
  let recorrenciaC = {};         

  filtrados.forEach(d => {
    lucroD[d.doutor] = (lucroD[d.doutor] || 0) + (d.lucroClinica || 0);
    atendimentosD[d.doutor] = (atendimentosD[d.doutor] || 0) + 1;
    if (!pacientesUnicosD[d.doutor]) pacientesUnicosD[d.doutor] = new Set();
    pacientesUnicosD[d.doutor].add(d.paciente);
    
    const chave = d.doutor + " | " + d.procedimento;
    contagemP[chave] = (contagemP[chave] || 0) + 1;
    rentabilidadeP[d.procedimento] = (rentabilidadeP[d.procedimento] || 0) + (d.lucroClinica || 0);
    recorrenciaC[d.paciente] = (recorrenciaC[d.paciente] || 0) + 1;
  });

  let pacientesUnicosCount = {};
  for (let dr in pacientesUnicosD) {
    pacientesUnicosCount[dr] = pacientesUnicosD[dr].size;
  }

  const totalFaturamento = filtrados.reduce((acc, d) => acc + (d.particular || 0) + (d.vplano || 0), 0);

  return {
    lucroDoutor: lucroD,
    atendimentosDoutor: atendimentosD,
    pacientesUnicosDoutor: pacientesUnicosCount,
    topProcedimentos: Object.entries(contagemP)
      .sort((a,b) => b[1] - a[1])
      .slice(0, 10),
    rentabilidade: Object.entries(rentabilidadeP)
      .sort((a,b) => b[1] - a[1])
      .slice(0, 5),
    clientes: Object.entries(recorrenciaC)
      .sort((a,b) => b[1] - a[1])
      .slice(0, 5),
    totalAtendimentos: filtrados.length,
    totalFaturamento: totalFaturamento
  };
}

function getDadosHolerite(nomeD, mesR) {
  const dados = getDadosTabelaERP();
  
  const filtrados = dados.filter(d => {
    if (!d.data) return false;
    const mesAno = d.data.split('/')[1] + '/' + d.data.split('/')[2];
    return mesAno === mesR && d.doutor === nomeD;
  });
  
  let totalBruto = 0;
  let totalLiquidoDentista = 0;
  let totalClinica = 0;
  let totalLab = 0;
  let pagamentos = {};
  
  filtrados.forEach(f => {
    const valorAtendimento = (f.particular + f.vplano);
    totalBruto += valorAtendimento;
    totalLiquidoDentista += (f.parteDentista || 0);
    totalClinica += (f.lucroClinica || 0);
    totalLab += (f.vlab || 0);
    
    const forma = f.pagamento ? f.pagamento : '';
    if (forma) {
      pagamentos[forma] = (pagamentos[forma] || 0) + valorAtendimento;
    }
  });
  
  return {
    nome: nomeD, 
    mes: mesR, 
    totalBruto: totalBruto,
    totalLiquidoDentista: totalLiquidoDentista,
    totalClinica: totalClinica,
    totalLab: totalLab,
    pagamentos: pagamentos,
    contagem: filtrados.length,
    dataGeracao: Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy"),
    particular: filtrados.filter(f => f.particular > 0).map(f => ({ 
      paciente: f.paciente, 
      procedimento: f.procedimento, 
      valor: f.particular,
      parteDentista: f.parteDentista || 0,
      vlab: f.vlab || 0,
      forma: f.pagamento ? f.pagamento : ''
    })),
    plano: filtrados.filter(f => f.vplano > 0).map(f => ({ 
      paciente: f.paciente, 
      procedimento: f.procedimento, 
      valor: f.vplano,
      parteDentista: f.parteDentista || 0,
      vlab: f.vlab || 0,
      forma: f.pagamento ? f.pagamento : ''
    }))
  };
}

function getListasDropdown() {
  var listas = getListasParaForm();
  listas.pacientes = getListaPacientes();
  return listas;
}

function cadastrarItemAuxiliar(tipo, valor, percPart, percPlano) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetConfig = ss.getSheetByName("config");
  if (!sheetConfig) return "Erro: aba 'config' não encontrada.";
  let coluna;
  if (tipo === 'procedimento') coluna = 1;
  else if (tipo === 'plano') coluna = 3;
  else if (tipo === 'pagamento') coluna = 9;
  else if (tipo === 'doutor') {
    const ultimaLinha = sheetConfig.getLastRow() + 1;
    sheetConfig.getRange(ultimaLinha, 11).setValue(valor.toUpperCase());
    sheetConfig.getRange(ultimaLinha, 12).setValue(Number(percPart) || 0.4);
    sheetConfig.getRange(ultimaLinha, 13).setValue(Number(percPlano) || 0.45);
    return "Doutor cadastrado com sucesso!";
  } else {
    return "Tipo de item inválido.";
  }
  const ultimaLinha = sheetConfig.getLastRow() + 1;
  sheetConfig.getRange(ultimaLinha, coluna).setValue(valor);
  return "Item cadastrado com sucesso!";
}

function reorganizarPlanilha() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(NOME_ABA_DADOS);
    const ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < 3) return "Sem dados para organizar";
    
    const dadosValidos = [];
    const formulas = [];
    
    for (let i = 3; i <= ultimaLinha; i++) {
      const valorB = sheet.getRange(i, 2).getValue();
      if (valorB && valorB.toString().trim() !== "") {
        const linhaCompleta = [];
        for (let j = 1; j <= 14; j++) {
          linhaCompleta.push(sheet.getRange(i, j).getValue());
        }
        dadosValidos.push(linhaCompleta);
        formulas.push({
          M: sheet.getRange(i, 13).getFormula()
        });
      }
    }
    
    dadosValidos.sort((a, b) => {
      const dataA = a[0] instanceof Date ? a[0].getTime() : new Date(a[0]).getTime();
      const dataB = b[0] instanceof Date ? b[0].getTime() : new Date(b[0]).getTime();
      return dataA - dataB;
    });
    
    if (ultimaLinha >= 3) {
      sheet.getRange("A3:N" + (ultimaLinha + 10)).clearContent();
    }
    
    for (let i = 0; i < dadosValidos.length; i++) {
      const linha = 3 + i;
      const dados = dadosValidos[i];
      for (let j = 1; j <= 14; j++) {
        sheet.getRange(linha, j).setValue(dados[j-1]);
      }
      if (formulas[i] && formulas[i].M) {
        sheet.getRange(linha, 13).setFormula(formulas[i].M);
      }
    }
    return "Planilha reorganizada: " + dadosValidos.length + " registros (antigos primeiro)";
  } catch (erro) {
    return "Erro ao reorganizar: " + erro.toString();
  }
}

// ===== LOGIN E USUÁRIOS =====
function validarLogin(login, senha) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return { success: false, msg: "Aba USUARIOS não existe. Contate o administrador." };
    
    const dados = sheet.getDataRange().getValues();
    const loginDigitado = (login || "").toString().trim().toLowerCase();
    const senhaDigitada = (senha || "").toString().trim();

    for (let i = 1; i < dados.length; i++) {
      let loginPlanilha = (dados[i][1] || "").toString().trim().toLowerCase();
      let senhaPlanilha = (dados[i][2] || "").toString().trim();

      if (loginPlanilha === loginDigitado && senhaPlanilha === senhaDigitada) {
        return {
          success: true,
          nome: (dados[i][0] || "").toString(),
          nivel: (dados[i][3] || "COMUM").toString().trim().toUpperCase(),
          permissoes: (dados[i][4] || "").toString()
        };
      }
    }
    return { success: false, msg: "Credenciais Incorretas" };
  } catch (erro) {
    return { success: false, msg: "Erro no servidor: " + erro.toString() }; 
  }
}

function listarUsuarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("USUARIOS");
  if (!sheet) return [];
  const dados = sheet.getDataRange().getValues();
  const usuarios = [];
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][1]) {
      usuarios.push({
        id: i,
        nome: dados[i][0],
        login: dados[i][1],
        senha: dados[i][2],
        nivel: dados[i][3] || "COMUM",
        permissoes: dados[i][4] || ""
      });
    }
  }
  return usuarios;
}

function getUsuarioPorId(idLinha) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("USUARIOS");
  if (!sheet) return null;
  const linha = Number(idLinha) + 1;
  const valores = sheet.getRange(linha, 1, 1, 5).getValues()[0];
  return {
    id: idLinha,
    nome: valores[0],
    login: valores[1],
    senha: valores[2],
    nivel: valores[3],
    permissoes: valores[4]
  };
}

function salvarUsuario(dados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) {
      sheet = ss.insertSheet("USUARIOS");
      sheet.getRange(1,1,1,5).setValues([["NOME","LOGIN","SENHA","NIVEL","PERMISSOES"]]);
    }
    
    if (dados.id && dados.id !== "") {
      const linha = Number(dados.id) + 1;
      sheet.getRange(linha, 1, 1, 5).setValues([[
        dados.nome, dados.login, dados.senha, dados.nivel, dados.permissoes
      ]]);
      return "Usuário atualizado com sucesso!";
    } else {
      const existing = sheet.createTextFinder(dados.login).matchEntireCell(true).findNext();
      if (existing) return "Erro: Login já existe!";
      const ultimaLinha = sheet.getLastRow();
      sheet.getRange(ultimaLinha + 1, 1, 1, 5).setValues([[
        dados.nome, dados.login, dados.senha, dados.nivel, dados.permissoes
      ]]);
      return "Usuário criado com sucesso!";
    }
  } catch (e) {
    return "Erro ao salvar: " + e.toString();
  }
}

function excluirUsuario(idLinha) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return "Aba não encontrada";
    const linha = Number(idLinha) + 1;
    sheet.deleteRow(linha);
    return "Usuário excluído permanentemente.";
  } catch (e) {
    return "Erro ao excluir: " + e.toString();
  }
}

// ===== MÓDULO ESTOQUE =====
function getDadosEstoqueCompleto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetProd = ss.getSheetByName("ESTOQUE_PRODUTOS");
  const sheetMov = ss.getSheetByName("ESTOQUE_MOV");
  if (!sheetProd || !sheetMov) return { produtos: [], unidadesTotais: 0 };

  const produtos = sheetProd.getDataRange().getValues();
  const movimentos = sheetMov.getDataRange().getValues();
  let resumo = {};
  let unidadesTotaisGeral = 0;

  for (let i = 1; i < produtos.length; i++) {
    let id = produtos[i][0];
    let nome = produtos[i][1];
    if (!nome) continue;
    resumo[nome] = {
      id: id || i, 
      nome: nome,
      categoria: produtos[i][2] || 'Geral',
      unidade: produtos[i][3] || 'Un',
      minimo: Number(produtos[i][4]) || 0,
      atual: 0,
      valorPatrimonio: 0
    };
  }
  
  for (let i = 1; i < movimentos.length; i++) {
    let tipo = movimentos[i][2]; let prod = movimentos[i][3]; let qtd = Number(movimentos[i][4]) || 0;
    let custo = Number(movimentos[i][6]) || 0;
    if (resumo[prod]) {
      if (tipo === "ENTRADA") { resumo[prod].atual += qtd; resumo[prod].valorPatrimonio += custo; }
      else if (tipo === "SAIDA") { resumo[prod].atual -= qtd; }
    }
  }

  const listaFinal = Object.values(resumo).map(item => {
    unidadesTotaisGeral += item.atual;
    let calculoSaude = item.minimo > 0 ? (item.atual / (item.minimo * 2)) * 100 : 100;
    return { 
      ...item, 
      saude: Math.min(Math.round(calculoSaude), 100), 
      custoMedio: item.atual > 0 ? (item.valorPatrimonio / item.atual) : 0 
    };
  });

  listaFinal.sort((a, b) => a.categoria.localeCompare(b.categoria) || a.nome.localeCompare(b.nome));
  return { produtos: listaFinal, unidadesTotais: unidadesTotaisGeral };
}

function salvarMovimentoEstoque(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetMov = ss.getSheetByName("ESTOQUE_MOV");
    const sheetProd = ss.getSheetByName("ESTOQUE_PRODUTOS");
    
    if (!sheetMov || !sheetProd) return "Erro: Abas de estoque não encontradas.";

    if (dados.tipo === "ENTRADA") {
      const buscaProd = sheetProd.createTextFinder(dados.produto).matchEntireCell(true).findNext();
      if (!buscaProd) {
        const novoId = sheetProd.getLastRow() + 1;
        const categoria = (dados.categoria && dados.categoria.trim() !== "") ? dados.categoria : "Geral";
        sheetProd.appendRow([novoId, dados.produto, categoria, "Un", 0]);
      }
    }

    let qtd = Number(dados.quantidade) || 1;
    let custoTotal = Number(dados.custoTotal) || 0;
    let custoUnit = custoTotal > 0 ? (custoTotal / qtd) : 0;

    const dataObj = new Date();
    const dataFormatada = `${String(dataObj.getDate()).padStart(2,'0')}/${String(dataObj.getMonth()+1).padStart(2,'0')}/${dataObj.getFullYear()}`;
    
    const linhaMovimento = [
      dataFormatada,
      dados.usuarioLogado,
      dados.tipo, 
      dados.produto,
      qtd,
      custoUnit,
      custoTotal,
      dados.tipo === "SAIDA" ? dados.doutor : "-", 
      dados.obs || ""
    ];

    sheetMov.appendRow(linhaMovimento);
    return "✅ Movimentação registrada com sucesso!";

  } catch (erro) {
    return "Erro no servidor: " + erro.toString();
  } finally {
    lock.releaseLock();
  }
}

function editarProdutoCatalogo(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("ESTOQUE_PRODUTOS");
    const data = sheet.getDataRange().getValues();

    let linhaAEditar = -1;
    for(let i = 1; i < data.length; i++){
      if(data[i][1] === dados.nomeOriginal){
         linhaAEditar = i + 1;
         break;
      }
    }

    if(linhaAEditar === -1) return "Erro: Produto não encontrado.";

    sheet.getRange(linhaAEditar, 2).setValue(dados.nomeNovo);
    sheet.getRange(linhaAEditar, 3).setValue(dados.categoria);
    sheet.getRange(linhaAEditar, 5).setValue(Number(dados.minimo) || 0);

    if (dados.nomeNovo !== dados.nomeOriginal) {
      const sheetMov = ss.getSheetByName("ESTOQUE_MOV");
      if (sheetMov) {
        const finder = sheetMov.getRange("D:D").createTextFinder(dados.nomeOriginal).matchEntireCell(true);
        finder.replaceAllWith(dados.nomeNovo);
      }
    }

    return "✅ Produto atualizado com sucesso!";
  } catch(e) {
    return "Erro no servidor: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

function excluirProdutoCatalogo(nomeProduto) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("ESTOQUE_PRODUTOS");
    const data = sheet.getDataRange().getValues();

    for(let i = 1; i < data.length; i++){
      if(data[i][1] === nomeProduto){
         sheet.deleteRow(i + 1);
         return "✅ Produto excluído do catálogo!";
      }
    }
    return "Erro: Produto não encontrado.";
  } catch(e) {
    return "Erro: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

function salvarMovimentoEstoqueLote(loteDados, tipoGeral) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetMov = ss.getSheetByName("ESTOQUE_MOV");
    const sheetProd = ss.getSheetByName("ESTOQUE_PRODUTOS");
    
    const dataObj = new Date();
    const dataFormatada = Utilities.formatDate(dataObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    let linhasParaAdicionar = [];

    for (let i = 0; i < loteDados.length; i++) {
      let dados = loteDados[i];
      
      if (tipoGeral === "ENTRADA") {
        const buscaProd = sheetProd.createTextFinder(dados.produto).matchEntireCell(true).findNext();
        if (!buscaProd) {
          const novoId = sheetProd.getLastRow() + 1;
          const categoria = (dados.categoria && dados.categoria.trim() !== "") ? dados.categoria : "Geral";
          sheetProd.appendRow([novoId, dados.produto, categoria, "Un", 0]);
        }
      }

      let qtd = Number(dados.quantidade) || 0;
      let custoTotal = Number(dados.custoTotal) || 0;
      let custoUnit = custoTotal > 0 ? (custoTotal / qtd) : 0;

      linhasParaAdicionar.push([
        dataFormatada,
        dados.usuarioLogado,
        tipoGeral, 
        dados.produto,
        qtd,
        custoUnit,
        custoTotal,
        tipoGeral === "SAIDA" ? (dados.doutor || "-") : "-", 
        dados.obs || ""
      ]);
    }

    if (linhasParaAdicionar.length > 0) {
      sheetMov.getRange(sheetMov.getLastRow() + 1, 1, linhasParaAdicionar.length, linhasParaAdicionar[0].length).setValues(linhasParaAdicionar);
    }

    return `✅ ${linhasParaAdicionar.length} itens registrados como ${tipoGeral}!`;
  } catch (erro) {
    return "Erro: " + erro.toString();
  } finally { lock.releaseLock(); }
}

// ===== COTAÇÕES =====
function gerarLinkCotacao() {
  const url = ScriptApp.getService().getUrl();
  const idCotacao = "COT" + Math.floor(Date.now() / 1000);
  return {
    id: idCotacao,
    link: url + "?page=cotacao&id=" + idCotacao
  };
}

function getItensParaCotar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetProd = ss.getSheetByName("ESTOQUE_PRODUTOS");
  const sheetMov = ss.getSheetByName("ESTOQUE_MOV");
  
  if (!sheetProd || !sheetMov) return [];
  
  const produtos = sheetProd.getDataRange().getValues();
  const movimentos = sheetMov.getDataRange().getValues();
  
  let estoque = {};
  for (let i = 1; i < produtos.length; i++) {
    if (produtos[i][1]) {
      estoque[produtos[i][1]] = { 
        id: produtos[i][0], 
        nome: produtos[i][1], 
        min: produtos[i][4] || 0, 
        atual: 0 
      };
    }
  }
  
  for (let i = 1; i < movimentos.length; i++) {
    let tipo = movimentos[i][2]; 
    let prod = movimentos[i][3]; 
    let qtd = Number(movimentos[i][4]) || 0;
    if (estoque[prod]) {
      if (tipo === "ENTRADA") estoque[prod].atual += qtd;
      if (tipo === "SAIDA") estoque[prod].atual -= qtd;
    }
  }
  
  return Object.values(estoque).filter(p => p.atual <= p.min);
}

function salvarRespostaFornecedor(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetCot = ss.getSheetByName("COTACAO_DADOS");
  if (!sheetCot) {
    sheetCot = ss.insertSheet("COTACAO_DADOS");
    sheetCot.getRange(1,1,1,6).setValues([["ID_COTACAO", "FORNECEDOR", "PRODUTO", "PRECO", "MARCA", "DATA"]]);
  }
  
  const dataHoje = new Date().toLocaleDateString('pt-BR');
  
  dados.itens.forEach(item => {
    if(item.preco > 0) {
      sheetCot.appendRow([
        dados.idCotacao,
        dados.fornecedor,
        item.nome,
        item.preco,
        item.marca || "-",
        dataHoje
      ]);
    }
  });
  return "✅ Preços enviados com sucesso! Obrigado.";
}

// ===== ANÁLISE DE COTAÇÃO =====
function getAnaliseCotacao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCot = ss.getSheetByName("COTACAO_DADOS");
  if (!sheetCot) return [];
  const dados = sheetCot.getDataRange().getValues();
  let analise = {};

  for (let i = 1; i < dados.length; i++) {
    let [idCot, fornecedor, produto, preco, marca] = dados[i];
    if (!analise[produto] || preco < analise[produto].preco) {
      analise[produto] = { produto, fornecedor, preco, marca, todos: [] };
    }
    analise[produto].todos.push({ fornecedor, preco, marca });
  }
  return Object.values(analise);
}

// ===== ANÁLISE DE MELHOR PREÇO =====
function getAnaliseMelhorPreco() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCot = ss.getSheetByName("COTACAO_DADOS");
  if (!sheetCot) return [];
  const dados = sheetCot.getDataRange().getValues();
  let analise = {};

  for (let i = 1; i < dados.length; i++) {
    let [idCot, fornecedor, produto, preco, marca] = dados[i];
    if (!analise[produto]) {
      analise[produto] = { produto, vencedor: {fornecedor, preco, marca}, opcoes: [] };
    }
    analise[produto].opcoes.push({ fornecedor, preco, marca });
    if (preco < analise[produto].vencedor.preco) {
      analise[produto].vencedor = { fornecedor, preco, marca };
    }
  }
  return Object.values(analise);
}

// ==========================================
// INTEGRAÇÃO COM O CRM DE PACIENTES (BLINDADO)
// ==========================================
function getPacientesCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("PACIENTES"); 
  if (!aba) return [];

  const dados = aba.getDataRange().getValues();
  if (dados.length < 2) return [];

  // Limpeza profunda do cabeçalho
  const cabecalho = dados[0].map(c => c.toString().toLowerCase().trim());

  // Função interna para achar a coluna certa sem erro
  const achar = (nome) => cabecalho.indexOf(nome.toLowerCase().trim());

  const lista = [];
  for (let ln = 1; ln < dados.length; ln++) {
    const r = dados[ln];
    const nomePaciente = r[achar("nome")];
    if (!nomePaciente) continue;

    // Tratamento de Data
    let dtNasc = r[achar("nascimento")];
    if (dtNasc instanceof Date) {
      dtNasc = Utilities.formatDate(dtNasc, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    }

    lista.push({
      nome: String(nomePaciente).trim(),
      // Se não achar a coluna, ele coloca vazio em vez de dar erro
      cpf: achar("cpf") >= 0 ? String(r[achar("cpf")] || "") : "",
      rg: achar("rg") >= 0 ? String(r[achar("rg")] || "") : "", // Verifique se o nome na planilha é "RG"
      nascimento: dtNasc ? String(dtNasc) : "",
      celular: achar("celular") >= 0 ? String(r[achar("celular")] || "") : "",
      telefone: achar("telefone") >= 0 ? String(r[achar("telefone")] || "") : "",
      email: achar("email") >= 0 ? String(r[achar("email")] || "") : "",
      convenio: achar("convenio") >= 0 ? String(r[achar("convenio")] || "") : "Particular",
      endereco: achar("endereco") >= 0 ? String(r[achar("endereco")] || "") : "",
      responsavel: achar("responsavel") >= 0 ? String(r[achar("responsavel")] || "") : ""
    });
  }
  return lista;
}

// ==========================================
// SALVAR/EDITAR PACIENTE COM TRAVA E OTIMIZAÇÃO DE ARRAY
// ==========================================
function cadastrarPacienteNoCRM(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // Espera na fila por até 15 segundos
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("PACIENTES");
    if (!aba) throw new Error("Aba PACIENTES não encontrada.");
    
    // Pega o cabeçalho inteiro e converte pra minúsculo para comparar
    const cabecalho = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
    const cabecalhoLower = cabecalho.map(c => c.toString().toLowerCase().trim());
    
    let linhaEdicao = -1;
    
    // SE FOR EDIÇÃO: Procura a linha usando o NOME ORIGINAL
    if (dados.nome_original && dados.nome_original.trim() !== "") {
       const iNomeBusca = cabecalhoLower.indexOf("nome");
       if (iNomeBusca !== -1) {
           const nomesNaPlanilha = aba.getRange(1, iNomeBusca + 1, aba.getLastRow(), 1).getValues();
           for (let i = 1; i < nomesNaPlanilha.length; i++) {
               if (nomesNaPlanilha[i][0].toString().toLowerCase() === dados.nome_original.toLowerCase()) {
                   linhaEdicao = i + 1;
                   break;
               }
           }
       }
    }
    
    // Arruma a data de YYYY-MM-DD para DD/MM/YYYY antes de salvar
    if (dados.Nascimento && dados.Nascimento.includes('-')) {
        const [a, m, d] = dados.Nascimento.split('-');
        dados.Nascimento = `${d}/${m}/${a}`;
    }
    
    if (linhaEdicao !== -1) {
        // === ATUALIZAR PACIENTE EXISTENTE (10x MAIS RÁPIDO AGORA) ===
        // Puxa a linha atual inteira para a memória
        let linhaAtual = aba.getRange(linhaEdicao, 1, 1, cabecalho.length).getValues()[0];
        
        for (let key in dados) {
            if (key === "nome_original") continue; // Pula o campo oculto
            let indexColuna = cabecalhoLower.indexOf(key.toLowerCase().trim());
            if (indexColuna !== -1) {
                linhaAtual[indexColuna] = dados[key]; // Atualiza na memória
            }
        }
        // Injeta a linha inteira de volta na planilha de uma vez só!
        aba.getRange(linhaEdicao, 1, 1, cabecalho.length).setValues([linhaAtual]);
        
    } else {
        // === CADASTRAR NOVO PACIENTE ===
        let novaLinha = new Array(cabecalho.length).fill("");
        for (let key in dados) {
            if (key === "nome_original") continue;
            let indexColuna = cabecalhoLower.indexOf(key.toLowerCase().trim());
            if (indexColuna !== -1) {
                novaLinha[indexColuna] = dados[key];
            }
        }
        aba.appendRow(novaLinha);
    }
    
    SpreadsheetApp.flush(); // FORÇA A GRAVAÇÃO ANTES DE LIBERAR A TRAVA!
    return linhaEdicao !== -1 ? "Dados do paciente atualizados com sucesso!" : "Novo paciente cadastrado com sucesso!";
    
  } catch (e) {
    throw new Error("O sistema está ocupado processando outro cadastro. Tente salvar novamente.");
  } finally {
    lock.releaseLock();
  }
}
// ==========================================
// SALVAR AGENDAMENTO COM TRAVA E DEDUPLICAÇÃO
// ==========================================
function salvarAgendamentoNaPlanilha(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // 15 segundos para dar folga na fila
    
    // Tratamento de data brasileira
    if (d.data && d.data.includes('-')) {
        const [ano, mes, dia] = d.data.split('-');
        d.data = `${dia}/${mes}/${ano}`;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("AGENDA");
    if (!aba) return "ERRO: Aba AGENDA não encontrada";

    // --- LÓGICA DE EDIÇÃO (BUSCA RÁPIDA) ---
    if (d.id && d.id !== "") {
      const busca = aba.getRange("A:A").createTextFinder(d.id).matchEntireCell(true).findNext();
      if (busca) {
        const linha = busca.getRow();
        aba.getRange(linha, 1, 1, 9).setValues([[d.id, d.data, d.hora, d.paciente, "", d.doutor, d.procedimento, d.status, d.obs]]);
        SpreadsheetApp.flush();
        return "OK";
      }
    }

    // --- LÓGICA DE NOVO REGISTRO (COM CHECAGEM BLINDADA) ---
    const registros = aba.getDataRange().getValues();
    
    // Verifica se já existe um agendamento IGUAL (Paciente + Data + Hora)
    const duplicado = registros.some(r => {
      // Traduz os Objetos do Sheets para texto antes de comparar
      let rData = r[1] instanceof Date ? Utilities.formatDate(r[1], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : String(r[1]).trim();
      let rHora = r[2] instanceof Date ? Utilities.formatDate(r[2], ss.getSpreadsheetTimeZone(), "HH:mm") : String(r[2]).trim().substring(0, 5);
      
      return rData === d.data && 
             rHora === d.hora && 
             String(r[3]).toLowerCase().trim() === d.paciente.toLowerCase().trim();
    });

    // Se o clone tentar passar, a porta bate na cara dele
    if (duplicado) {
      return "ERRO: Este agendamento já foi registrado por outro usuário simultaneamente.";
    }

    // Se for limpo, grava a linha
    aba.appendRow([Date.now(), d.data, d.hora, d.paciente, "", d.doutor, d.procedimento, d.status, d.obs]);
    
    SpreadsheetApp.flush(); 
    return "OK";
    
  } catch (e) {
    return "ERRO: O servidor está ocupado. Tente salvar novamente em 5 segundos.";
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// EXCLUIR AGENDAMENTO (BACK-END)
// ==========================================
function excluirAgendamentoNaPlanilha(idAgendamento) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("AGENDA");
    
    if (!aba) return "ERRO: Aba AGENDA não encontrada.";

    // Busca ultrarrápida pelo ID na Coluna A (1)
    const busca = aba.getRange("A:A").createTextFinder(idAgendamento.toString()).matchEntireCell(true).findNext();
    
    if (busca) {
      const linha = busca.getRow();
      // Apaga a linha inteira para manter o banco de dados limpo
      aba.deleteRow(linha); 
      return "OK";
    }
    
    return "ERRO: Agendamento não encontrado no banco de dados.";
    
  } catch (e) {
    return "ERRO: Falha ao excluir - " + e.message;
  } finally {
    lock.releaseLock();
  }
}

function getAgendamentosCalendario(filtroDoutor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("AGENDA");
  if (!aba) return [];
  
  const dados = aba.getDataRange().getValues();
  const eventos = [];
  
  const cores = {
    "AGENDADO": "#4318FF",        // Roxo Premium
    "CONFIRMADO": "#05CD99",      // Verde Sucesso
    "EM ATENDIMENTO": "#FFB800",   // Amarelo
    "FINALIZADO": "#2B3674",      // Azul Escuro (Financeiro)
    "FALTOU": "#EE5D50",          // Vermelho
    "CANCELADO": "#A3AED0"        // Cinza
  };

  for (let i = 1; i < dados.length; i++) {
    const [id, data, hora, paciente, cel, doutor, proc, status, obs] = dados[i];
    
    if (!data || !hora) continue;

    // === TRAVA DO FILTRO POR DOUTOR ===
    // Se existir um filtro selecionado lá na tela, e o doutor desta linha for diferente, ele ignora e pula pro próximo!
    if (filtroDoutor && filtroDoutor.trim() !== "" && doutor !== filtroDoutor) {
        continue; 
    }
    // ==================================

    // Converter data DD/MM/AAAA para AAAA-MM-DD
    let dataISO = "";
    if (data instanceof Date) {
      dataISO = Utilities.formatDate(data, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    } else {
      const p = data.split('/');
      dataISO = `${p[2]}-${p[1]}-${p[0]}`;
    }
    
    // Formatar hora para HH:mm
    let horaFormatada = "";
    if (hora instanceof Date) {
      horaFormatada = Utilities.formatDate(hora, ss.getSpreadsheetTimeZone(), "HH:mm");
    } else {
      horaFormatada = hora.toString().substring(0, 5);
    }

    eventos.push({
      id: id.toString(),
      title: paciente + " - " + proc,
      start: dataISO + "T" + horaFormatada + ":00",
      backgroundColor: cores[status] || "#4318FF",
      borderColor: "transparent",
      extendedProps: { 
        data: data instanceof Date ? Utilities.formatDate(data, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : data,
        hora: horaFormatada,
        paciente: paciente,
        doutor: doutor,
        proc: proc,
        status: status,
        obs: obs
      }
    });
  }
  return eventos;
}

// ==========================================
// PRONTUÁRIO: EVOLUÇÃO CLÍNICA
// ==========================================
function salvarNovaEvolucao(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let aba = ss.getSheetByName("EVOLUCAO_CLINICA");
    if (!aba) aba = ss.insertSheet("EVOLUCAO_CLINICA"); // Cria a aba se não existir
    
    const dataRegistro = new Date();
    
    aba.appendRow([
      dados.paciente,
      dados.data,
      dados.proc,
      dados.dente,
      dados.doutor,
      dados.obs,
      dataRegistro
    ]);
    
    SpreadsheetApp.flush();
    return "OK";
    
  } catch (e) {
    throw new Error("Sistema ocupado gravando outra evolução. Tente novamente.");
  } finally {
    lock.releaseLock();
  }
}

function getEvolucaoPaciente(nomePaciente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("EVOLUCAO");
  if (!aba) return [];
  
  const dados = aba.getDataRange().getValues();
  const historico = [];
  
  // Pula o cabeçalho (i=1)
  for (let i = 1; i < dados.length; i++) {
    const [id, data, paciente, doutor, proc, regiao, obs] = dados[i];
    
    // Se o nome da linha for igual ao do paciente solicitado, adiciona na lista
    if (paciente.toString().toLowerCase() === nomePaciente.toLowerCase()) {
       // Tratamento de data caso o Google Sheets entenda como objeto Date
       let dataStr = data;
       if (data instanceof Date) {
         dataStr = Utilities.formatDate(data, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
       }
       
       historico.push({
         id: id,
         data: dataStr,
         doutor: doutor,
         procedimento: proc,
         regiao: regiao,
         obs: obs
       });
    }
  }
  
  // Ordena para os mais recentes ficarem no topo
  historico.sort((a, b) => {
     let [dA, mA, aA] = a.data.split('/');
     let [dB, mB, aB] = b.data.split('/');
     return new Date(aB, mB-1, dB) - new Date(aA, mA-1, dA);
  });
  
  return historico;
}

function salvarDadosAnamnese(nome, jsonDados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let aba = ss.getSheetByName("ANAMNESE");
    if (!aba) aba = ss.insertSheet("ANAMNESE");
    
    const dados = aba.getDataRange().getValues();
    let linhaDestino = -1;
    
    for(let i=0; i<dados.length; i++) {
      if(dados[i][0] === nome) { linhaDestino = i + 1; break; }
    }
    
    const dataHoje = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
    
    if(linhaDestino !== -1) {
      aba.getRange(linhaDestino, 2, 1, 2).setValues([[dataHoje, jsonDados]]);
    } else {
      aba.appendRow([nome, dataHoje, jsonDados]);
    }
    
    SpreadsheetApp.flush();
    return "OK";
    
  } catch (e) {
    throw new Error("Sistema de prontuário ocupado. Tente novamente.");
  } finally {
    lock.releaseLock();
  }
}

function getDadosAnamnese(nome) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("ANAMNESE");
  if (!aba) return null;
  const dados = aba.getDataRange().getValues();
  for(let i=0; i<dados.length; i++) {
    if(dados[i][0] === nome) return dados[i][2];
  }
  return null;
}

function salvarEstadoDente(nome, dente, status) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let aba = ss.getSheetByName("ODONTOGRAMA_DADOS");
    if (!aba) aba = ss.insertSheet("ODONTOGRAMA_DADOS");
    
    const dados = aba.getDataRange().getValues();
    let linha = -1;
    
    for(let i=0; i<dados.length; i++) {
      if(dados[i][0] === nome && dados[i][1] == dente) {
        linha = i + 1; break;
      }
    }
    
    const data = new Date();
    if (linha !== -1) {
      aba.getRange(linha, 3, 1, 3).setValues([[status, "", data]]);
    } else {
      aba.appendRow([nome, dente, status, "", data]);
    }
    
    SpreadsheetApp.flush();
    
  } catch (e) {
    // Odontograma salva silenciosamente, então podemos apenas dar um log no erro.
    console.log("Concorrência no Odontograma. Ocultando erro para não assustar o usuário.");
  } finally {
    lock.releaseLock();
  }
}


function getOdontogramaPaciente(nome) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("ODONTOGRAMA_DADOS");
  if (!aba) return [];
  const dados = aba.getDataRange().getValues();
  const resultado = [];
  for(let i=1; i<dados.length; i++) {
    if(dados[i][0] === nome) {
      resultado.push({ dente: dados[i][1], status: dados[i][2] });
    }
  }
  return resultado;
}

// ==========================================
// 1. O MOTOR DE REGRAS E CÁLCULO
// ==========================================
function calcularRetornosPendentes(regrasRecorrencia) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaEvo = ss.getSheetByName("EVOLUCAO"); 
  const abaTotal = ss.getSheetByName("TOTAL");  
  const abaAgenda = ss.getSheetByName("AGENDA");
  const abaPacientes = ss.getSheetByName("PACIENTES"); 
  
  const agendamentos = abaAgenda ? abaAgenda.getDataRange().getValues() : [];
  const pacientes = abaPacientes ? abaPacientes.getDataRange().getValues() : [];
  
  // Mapear telefones
  const mapaTelefones = {};
  if (pacientes.length > 0) {
    const cabPac = pacientes[0].map(c => c.toString().toLowerCase().trim());
    const iNomeP = cabPac.indexOf("nome");
    const iCelularP = cabPac.indexOf("celular");
    if (iNomeP !== -1 && iCelularP !== -1) {
      for (let i = 1; i < pacientes.length; i++) {
        if(pacientes[i][iNomeP]) mapaTelefones[pacientes[i][iNomeP].toString().toLowerCase()] = pacientes[i][iCelularP];
      }
    }
  }

  // Mapear Agendados
  const pacientesAgendados = new Set();
  const hoje = new Date();
  hoje.setHours(0,0,0,0);
  
  for (let i = 1; i < agendamentos.length; i++) {
    let dataAg = agendamentos[i][1]; 
    let statusAg = agendamentos[i][7]; 
    let pacAg = agendamentos[i][3];
    
    if (statusAg !== "FINALIZADO" && statusAg !== "FALTOU" && dataAg && pacAg) {
      let dataConsulta = null;
      if (dataAg instanceof Date) dataConsulta = dataAg;
      else {
          let str = dataAg.toString();
          if (str.includes('/')) { let [d, m, a] = str.split('/'); dataConsulta = new Date(a, m - 1, d); } 
          else if (str.includes('-')) { let [a, m, d] = str.split('-'); dataConsulta = new Date(a, m - 1, d); }
      }
      if (dataConsulta && dataConsulta >= hoje) pacientesAgendados.add(pacAg.toString().toLowerCase().trim()); 
    }
  }

  const ultimaVisita = {}; 

  function registrarProcedimento(nomePac, dataVal, nomeProc) {
      if (!nomePac || !dataVal || !nomeProc) return;
      let pacStr = nomePac.toString().toLowerCase().trim();
      let procStr = nomeProc.toString().toLowerCase();
      let regra = regrasRecorrencia.find(r => procStr.includes(r.proc.toLowerCase()));
      
      if (regra) {
          let dataProcObj; let dataFormatada = "";
          if (dataVal instanceof Date) { dataProcObj = dataVal; dataFormatada = Utilities.formatDate(dataVal, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy"); } 
          else {
              let str = dataVal.toString();
              if (str.includes('/')) { let [d, m, a] = str.split('/'); dataProcObj = new Date(a, m - 1, d); dataFormatada = str; } 
              else return;
          }
          if (!ultimaVisita[pacStr]) ultimaVisita[pacStr] = {};
          if (!ultimaVisita[pacStr][regra.proc] || dataProcObj > ultimaVisita[pacStr][regra.proc].dataReal) {
              ultimaVisita[pacStr][regra.proc] = { dataReal: dataProcObj, dataStr: dataFormatada, diasRecorrencia: regra.dias, nomeOriginal: nomePac };
          }
      }
  }

  if (abaEvo) { const evoData = abaEvo.getDataRange().getValues(); for (let i = 1; i < evoData.length; i++) registrarProcedimento(evoData[i][2], evoData[i][1], evoData[i][4]); }
  if (abaTotal) { const totalData = abaTotal.getDataRange().getValues(); for (let i = 2; i < totalData.length; i++) registrarProcedimento(totalData[i][2], totalData[i][0], totalData[i][3]); }

  const listaRetornos = [];
  const daquiA15Dias = new Date(); daquiA15Dias.setDate(hoje.getDate() + 15); 

  // Puxar Lixeira
  const abaOcultos = ss.getSheetByName("RECALL_OCULTOS");
  const listaNegra = new Set();
  if (abaOcultos) {
      const dadosOcultos = abaOcultos.getDataRange().getValues();
      for (let i = 1; i < dadosOcultos.length; i++) {
          let chave = dadosOcultos[i][0].toString().toLowerCase().trim() + "|" + dadosOcultos[i][1].toString().toLowerCase().trim() + "|" + dadosOcultos[i][2].toString().trim();
          listaNegra.add(chave);
      }
  }

  for (let pac in ultimaVisita) {
    for (let proc in ultimaVisita[pac]) {
      let info = ultimaVisita[pac][proc];
      if (listaNegra.has(pac + "|" + proc.toLowerCase().trim() + "|" + info.dataStr.trim())) continue; // Pula os apagados

      let dataVencimento = new Date(info.dataReal);
      dataVencimento.setDate(dataVencimento.getDate() + info.diasRecorrencia);
      
      if (dataVencimento <= daquiA15Dias) {
        let jaAgendado = pacientesAgendados.has(pac);
        let diasAtraso = Math.floor((hoje - dataVencimento) / (1000 * 60 * 60 * 24));
        let status = jaAgendado ? "AGENDADO" : (diasAtraso > 0 ? "VENCIDO" : "VENCE EM BREVE");
        
        listaRetornos.push({ paciente: info.nomeOriginal, procedimento: proc, ultimaData: info.dataStr, vencimento: Utilities.formatDate(dataVencimento, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy"), celular: mapaTelefones[pac] || "", status: status, atraso: diasAtraso > 0 ? diasAtraso + " dias" : "No prazo" });
      }
    }
  }
  return listaRetornos.sort((a, b) => { if (a.status === "AGENDADO" && b.status !== "AGENDADO") return 1; if (a.status !== "AGENDADO" && b.status === "AGENDADO") return -1; return 0; });
}

// ==========================================
// 2. API DA LIXEIRA INTELIGENTE
// ==========================================
function apiOcultarRecall(paciente, procedimento, dataAtendimento) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let aba = ss.getSheetByName("RECALL_OCULTOS");
    if(!aba) { aba = ss.insertSheet("RECALL_OCULTOS"); aba.appendRow(["PACIENTE_IGNORADO", "PROCEDIMENTO", "DATA_ATENDIMENTO", "DATA_REMOCAO"]); }
    aba.appendRow([paciente, procedimento, dataAtendimento, new Date()]);
    SpreadsheetApp.flush();
    return "Registro específico ignorado!";
  } finally { lock.releaseLock(); }
}

// ==========================================
// 3. API DAS CONFIGURAÇÕES
// ==========================================
function apiRegrasRecorrencia(acao, procNome, dias) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let aba = ss.getSheetByName("CONFIG_REGRAS");
    if(!aba) { aba = ss.insertSheet("CONFIG_REGRAS"); aba.appendRow(["PROCEDIMENTO", "DIAS_RECORRENCIA"]); aba.appendRow(["Limpeza", 180]); aba.appendRow(["Manutenção", 30]); }
    const dados = aba.getDataRange().getValues();

    if(acao === "LER") {
      let regras = [];
      for(let i=1; i<dados.length; i++) if(dados[i][0]) regras.push({ proc: dados[i][0].toString(), dias: Number(dados[i][1]) });
      return regras;
    }
    if(acao === "SALVAR") {
      let linha = -1;
      for(let i=1; i<dados.length; i++) if(dados[i][0].toString().toLowerCase() === procNome.toLowerCase()) { linha = i + 1; break; }
      if(linha !== -1) aba.getRange(linha, 2).setValue(dias); else aba.appendRow([procNome, dias]);
      SpreadsheetApp.flush(); return "Salvo";
    }
    if(acao === "EXCLUIR") {
      for(let i=1; i<dados.length; i++) if(dados[i][0].toString().toLowerCase() === procNome.toLowerCase()) { aba.deleteRow(i + 1); SpreadsheetApp.flush(); return "Excluído"; }
    }
  } finally { lock.releaseLock(); }
}

// ==========================================
// MÓDULO DE LABORATÓRIO E NOTIFICAÇÕES
// ==========================================

function apiSalvarLaboratorio(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("LABORATORIO") || ss.insertSheet("LABORATORIO");
  
  // Se for nova aba, coloca cabeçalho
  if (aba.getLastRow() === 0) {
    aba.appendRow(["ID_LAB", "PACIENTE", "NOME_LABORATORIO", "DATA_ENVIO", "PRAZO_DIAS", "PREVISAO_ENTREGA", "STATUS"]);
  }

  const idLab = "LAB_" + new Date().getTime();
  const hoje = new Date();
  const previsao = new Date();
  previsao.setDate(hoje.getDate() + parseInt(dados.prazo));

  aba.appendRow([
    idLab,
    dados.paciente,
    dados.nomeLab,
    Utilities.formatDate(hoje, "GMT-3", "dd/MM/yyyy"),
    dados.prazo,
    Utilities.formatDate(previsao, "GMT-3", "dd/MM/yyyy"),
    "PENDENTE"
  ]);

  return { success: true, id: idLab };
}

function apiGetNotificacoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaLab = ss.getSheetByName("LABORATORIO");
  const notificacoes = [];
  const hoje = new Date();
  hoje.setHours(0,0,0,0);

  if (abaLab) {
    const dados = abaLab.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      const [id, paciente, lab, envio, prazo, previsaoStr, status] = dados[i];
      
      if (status === "FINALIZADO") continue;

      // Converte string dd/MM/yyyy para objeto Date
      const partes = previsaoStr.split('/');
      const dataPrevisao = new Date(partes[2], partes[1] - 1, partes[0]);
      
      const diffTime = dataPrevisao - hoje;
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

      if (diffDays <= 1) { // Notifica se falta 1 dia ou se já venceu
        notificacoes.push({
          tipo: "LABORATORIO",
          titulo: diffDays === 1 ? "Entrega de Lab Amanhã" : "Atraso de Laboratório!",
          mensagem: `${paciente} - ${lab} (Previsto: ${previsaoStr})`,
          paciente: paciente,
          idRef: id,
          urgencia: diffDays < 0 ? "alta" : "media"
        });
      }
    }
  }

  // Aqui você pode adicionar lógica para outras notificações (estoque baixo, aniversários, etc)
  
  return notificacoes;
}

// ==========================================
// MÓDULO DE LABORATÓRIO E NOTIFICAÇÕES
// ==========================================
function apiSalvarLaboratorio(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let aba = ss.getSheetByName("LABORATORIO");

    if (!aba) {
      aba = ss.insertSheet("LABORATORIO");
      aba.appendRow(["ID_LAB", "PACIENTE", "PROCEDIMENTO", "NOME_LABORATORIO", "DATA_ENVIO", "PRAZO_DIAS", "PREVISAO_ENTREGA", "STATUS"]);
    }
    
    const hoje = new Date();
    const previsao = new Date();
    previsao.setDate(hoje.getDate() + parseInt(dados.prazo));
    const previsaoFormatada = Utilities.formatDate(previsao, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");

    // SE FOR EDIÇÃO DE UM TRABALHO EXISTENTE
    if (dados.idEdicao) {
        const dataRange = aba.getDataRange().getValues();
        for(let i=1; i<dataRange.length; i++) {
            if(dataRange[i][0] === dados.idEdicao) {
                aba.getRange(i+1, 2).setValue(dados.paciente);
                aba.getRange(i+1, 3).setValue(dados.proc);
                aba.getRange(i+1, 4).setValue(dados.nomeLab);
                aba.getRange(i+1, 6).setValue(dados.prazo);
                aba.getRange(i+1, 7).setValue(previsaoFormatada); // Recalcula a entrega
                SpreadsheetApp.flush();
                return { success: true, msg: "Trabalho atualizado com sucesso!" };
            }
        }
    }

    // SE FOR UM TRABALHO NOVO
    const idLab = "LAB_" + new Date().getTime();
    aba.appendRow([idLab, dados.paciente, dados.proc, dados.nomeLab, Utilities.formatDate(hoje, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy"), dados.prazo, previsaoFormatada, "PENDENTE"]);

    SpreadsheetApp.flush();
    return { success: true, msg: "Trabalho enviado para o laboratório!" };
  } catch(e) { throw new Error("Erro ao salvar lab: " + e.message); } finally { lock.releaseLock(); }
}

function apiGetTodosLaboratorios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("LABORATORIO");
  if (!aba) return [];
  
  const dados = aba.getDataRange().getValues();
  const lista = [];
  
  for (let i = 1; i < dados.length; i++) {
    let envioFormatado = dados[i][4];
    let previsaoFormatada = dados[i][6];

    // MÁGICA: Se a planilha transformou em Data, nós forçamos a voltar pra Texto!
    if (envioFormatado instanceof Date) {
        envioFormatado = Utilities.formatDate(envioFormatado, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    }
    if (previsaoFormatada instanceof Date) {
        previsaoFormatada = Utilities.formatDate(previsaoFormatada, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    }

    lista.push({
      id: dados[i][0], 
      paciente: dados[i][1], 
      proc: dados[i][2],
      lab: dados[i][3], 
      envio: envioFormatado, 
      prazo: dados[i][5],
      previsao: previsaoFormatada, 
      status: dados[i][7]
    });
  }
  
  // Ordena para que os Pendentes fiquem em cima
  return lista.sort((a, b) => (a.status === "PENDENTE" ? -1 : 1));
}

function apiExcluirLaboratorio(idLab) {
   const lock = LockService.getScriptLock();
   try {
     lock.waitLock(10000);
     const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LABORATORIO");
     const dados = aba.getDataRange().getValues();
     for(let i=1; i<dados.length; i++) {
       if(dados[i][0] === idLab) { aba.deleteRow(i + 1); SpreadsheetApp.flush(); return "Pedido excluído!"; }
     }
   } finally { lock.releaseLock(); }
}

function apiFinalizarLaboratorio(idRef) {
   const lock = LockService.getScriptLock();
   try {
     lock.waitLock(10000);
     const ss = SpreadsheetApp.getActiveSpreadsheet();
     const aba = ss.getSheetByName("LABORATORIO");
     if(!aba) return;

     const dados = aba.getDataRange().getValues();
     for(let i=1; i<dados.length; i++) {
       if(dados[i][0] === idRef) {
         aba.getRange(i + 1, 8).setValue("FINALIZADO"); // Muda o status na coluna 8
         break;
       }
     }
     SpreadsheetApp.flush();
   } finally {
     lock.releaseLock();
   }
}

function apiSalvarPaciente(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("CADASTRO_CRM");
  if (!aba) throw new Error("Aba CADASTRO_CRM não encontrada!");
  
  const valores = aba.getDataRange().getValues();
  const nomeOriginal = dados.nome_original ? dados.nome_original.toString().trim() : "";
  let linhaAlvo = -1;

  // 1. TENTA LOCALIZAR O PACIENTE PARA EDIÇÃO
  if (nomeOriginal !== "") {
    for (let i = 1; i < valores.length; i++) {
      if (valores[i][0].toString().trim() === nomeOriginal) {
        linhaAlvo = i + 1;
        break;
      }
    }
  }

  // 2. MONTA OS DADOS (Mantenha a ordem exata das suas colunas!)
  const novaLinha = [
    dados.Nome, 
    dados.CPF, 
    dados.RG, 
    dados.Nascimento, 
    dados.Celular, 
    dados.Telefone, 
    dados.Email, 
    dados.sexo, 
    dados.profissao, 
    dados.convenio, 
    dados.Endereco, 
    dados.Responsavel
  ];

  // 3. DECIDE: EDITAR OU CRIAR NOVO
  if (linhaAlvo > -1) {
    // Grava por cima da linha existente
    aba.getRange(linhaAlvo, 1, 1, novaLinha.length).setValues([novaLinha]);
    return "Paciente '" + dados.Nome + "' atualizado com sucesso!";
  } else {
    // Verifica se já existe um paciente com esse nome exato para evitar duplicidade acidental
    for (let i = 1; i < valores.length; i++) {
       if (valores[i][0].toString().trim() === dados.Nome.toString().trim()) {
          return "Erro: Já existe um paciente cadastrado com este nome!";
       }
    }
    
    aba.appendRow(novaLinha);
    return "Paciente '" + dados.Nome + "' cadastrado com sucesso!";
  }
}

function apiGetNotificacoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaLab = ss.getSheetByName("LABORATORIO");
  const notificacoes = [];

  if (!abaLab) return notificacoes;

  const hoje = new Date();
  hoje.setHours(0,0,0,0);
  const dados = abaLab.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const [id, paciente, proc, lab, envio, prazo, previsaoStr, status] = dados[i];

    if (status === "FINALIZADO") continue;

    let dataPrevisao = null;
    if (previsaoStr instanceof Date) {
       dataPrevisao = previsaoStr;
    } else if (typeof previsaoStr === 'string' && previsaoStr.includes('/')) {
       const partes = previsaoStr.split('/');
       dataPrevisao = new Date(partes[2], partes[1] - 1, partes[0]);
    }

    if (dataPrevisao) {
        const diffTime = dataPrevisao - hoje;
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

        // Se faltar 2 dias ou menos, o sininho acorda!
        if (diffDays <= 2) {
          notificacoes.push({
            tipo: "LABORATORIO",
            titulo: diffDays < 0 ? "⚠️ Atraso no Lab!" : (diffDays === 0 ? "🚨 Entrega HOJE!" : "⏳ Chegando em breve"),
            mensagem: `${paciente} - ${proc}<br><small class="text-muted">Lab: ${lab}</small>`,
            paciente: paciente,
            idRef: id,
            urgencia: diffDays <= 0 ? "alta" : "media"
          });
        }
    }
  }
  return notificacoes;
}

// ==========================================
// MÓDULO: PROCEDIMENTOS E PREÇOS (ABA CONFIG COL A e B)
// ==========================================
function apiGetProcedimentosComPreco() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIG");
  if (!aba) return [];
  
  // Pega os dados apenas das colunas A e B, a partir da linha 2
  const dados = aba.getRange("A2:B").getValues(); 
  const lista = [];
  
  for (let i = 0; i < dados.length; i++) {
    if (dados[i][0]) {
      lista.push({ proc: dados[i][0].toString(), preco: dados[i][1] || 0 });
    }
  }
  return lista;
}

function apiSalvarProcedimentoComPreco(nome, preco) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIG");
    const dados = aba.getRange("A2:A").getValues();
    let linhaAlvo = -1;
    
    // 1. Tenta achar se o procedimento já existe para ATUALIZAR o preço
    for (let i = 0; i < dados.length; i++) {
      if (dados[i][0] && dados[i][0].toString().toUpperCase() === nome.toUpperCase()) {
        linhaAlvo = i + 2;
        break;
      }
    }
    
    // 2. Se não existe, acha a primeira linha vazia na Coluna A
    if (linhaAlvo === -1) {
      for (let i = 0; i < dados.length; i++) {
        if (!dados[i][0]) { linhaAlvo = i + 2; break; }
      }
      if (linhaAlvo === -1) linhaAlvo = dados.length + 2; // Vai pro fim
    }
    
    // Converte para formato numérico puro
    let valorNum = parseFloat(preco.toString().replace(/\./g, '').replace(',', '.'));
    if (isNaN(valorNum)) valorNum = 0;
    
    // Salva exatamente na Col A e Col B
    aba.getRange(linhaAlvo, 1).setValue(nome.toUpperCase());
    aba.getRange(linhaAlvo, 2).setValue(valorNum);
    
    SpreadsheetApp.flush();
    return "Salvo com sucesso!";
  } finally {
    lock.releaseLock();
  }
}

function apiExcluirProcedimentoComPreco(nome) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIG");
    const dados = aba.getRange("A2:B").getValues();
    const novosDados = [];
    
    // Pega todos, MENOS o que queremos excluir
    for (let i = 0; i < dados.length; i++) {
      if (dados[i][0] && dados[i][0].toString().toUpperCase() !== nome.toUpperCase()) {
        novosDados.push([dados[i][0], dados[i][1]]);
      }
    }
    
    // Preenche com linhas vazias até o tamanho original para limpar o final
    while(novosDados.length < dados.length) { novosDados.push(["", ""]); }
    
    // Sobrescreve apenas as colunas A e B, puxando tudo pra cima!
    aba.getRange(2, 1, novosDados.length, 2).setValues(novosDados);
    
    SpreadsheetApp.flush();
    return "Procedimento excluído!";
  } finally {
    lock.releaseLock();
  }
}

function apiEditarProcedimentoComPreco(nomeOriginal, nomeNovo, precoNovo) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIG");
    const dados = aba.getRange("A2:A").getValues();
    let linhaAlvo = -1;

    // Acha a linha exata onde está o procedimento antigo
    for (let i = 0; i < dados.length; i++) {
      if (dados[i][0] && dados[i][0].toString().toUpperCase() === nomeOriginal.toUpperCase()) {
        linhaAlvo = i + 2;
        break;
      }
    }

    if (linhaAlvo !== -1) {
      let valorNum = parseFloat(precoNovo.toString().replace(/\./g, '').replace(',', '.'));
      if (isNaN(valorNum)) valorNum = 0;

      // Sobrescreve as Colunas A e B com os novos dados
      aba.getRange(linhaAlvo, 1).setValue(nomeNovo.toUpperCase());
      aba.getRange(linhaAlvo, 2).setValue(valorNum);
      
      SpreadsheetApp.flush();
      return "Procedimento atualizado!";
    } else {
      throw new Error("Procedimento original não encontrado.");
    }
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MÓDULO GERADOR DE DOCUMENTOS EM PDF
// ==========================================

// ⚠️ ATENÇÃO: Cole aqui o ID da pasta que você criou no Google Drive!
const PASTA_DOCUMENTOS_ID = "1jA5yXoaiTJ4ZgfibZFYLVjLpbpZQtq2k"; 



// ==========================================
// GERADOR DE PDF COM ORGANIZAÇÃO EM PASTAS
// ==========================================

function apiGerarDocumentoPDF(tipo, paciente, dadosDocumento) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fusoHorario = ss.getSpreadsheetTimeZone();
    const dataAtual = Utilities.formatDate(new Date(), fusoHorario, "dd/MM/yyyy");
    
    // 1. Cria um HTML temporário lindíssimo para o documento
    let html = `
      <html>
        <head>
          <style>
            body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #333; margin: 40px; }
            .header { text-align: center; border-bottom: 2px solid #4318FF; padding-bottom: 20px; margin-bottom: 30px; }
            .logo { font-size: 28px; font-weight: bold; color: #2B3674; margin-bottom: 5px; }
            .sub-logo { font-size: 14px; color: #A3AED0; }
            .title { font-size: 22px; font-weight: bold; color: #4318FF; margin-bottom: 20px; text-transform: uppercase; }
            .info-box { background: #F4F7FE; padding: 15px; border-radius: 8px; margin-bottom: 30px; }
            .info-text { font-size: 14px; margin: 5px 0; }
            .content { font-size: 15px; line-height: 1.6; margin-bottom: 50px; min-height: 300px; }
            .signature { text-align: center; margin-top: 60px; }
            .line { border-top: 1px solid #333; width: 250px; margin: 0 auto 10px auto; }
            .footer { text-align: center; font-size: 12px; color: #A3AED0; margin-top: 40px; border-top: 1px solid #eee; padding-top: 20px; }
            
            /* Tabela de Orçamento */
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th { background-color: #4318FF; color: white; padding: 12px; text-align: left; }
            td { padding: 12px; border-bottom: 1px solid #eee; }
            .total-row { font-weight: bold; background-color: #F4F7FE; }
          </style>
        </head>
        <body>
          <div class="header">
            <div class="logo">Clínica Odontológica Ramos</div>
            <div class="sub-logo">Rua Exemplo, 123 - Centro | (11) 99999-9999</div>
          </div>
          
          <div class="info-box">
            <div class="info-text"><strong>Paciente:</strong> ${paciente}</div>
            <div class="info-text"><strong>Data:</strong> ${dataAtual}</div>
          </div>
    `;

    // 2. Preenche o miolo do documento dependendo do Tipo (Receita, Atestado ou Orçamento)
    if (tipo === "RECEITA") {
      html += `
        <div class="title">Receituário</div>
        <div class="content">${dadosDocumento.texto.replace(/\n/g, '<br>')}</div>
      `;
    } 
    else if (tipo === "ATESTADO") {
      html += `
        <div class="title">Atestado Odontológico</div>
        <div class="content">
          Atesto para os devidos fins que o(a) paciente <strong>${paciente}</strong>, 
          esteve em atendimento odontológico neste consultório no dia ${dataAtual}, 
          das ${dadosDocumento.horaInicio} às ${dadosDocumento.horaFim}.<br><br>
          Necessita de <strong>${dadosDocumento.diasRepouso} dia(s)</strong> de repouso por motivo de saúde.
          ${dadosDocumento.cid ? `<br><br>CID: ${dadosDocumento.cid}` : ''}
        </div>
      `;
    }
    else if (tipo === "ORCAMENTO") {
      let linhasTabela = '';
      let total = 0;
      dadosDocumento.itens.forEach(item => {
        linhasTabela += `<tr><td>${item.proc}</td><td style="text-align: right;">R$ ${item.valor}</td></tr>`;
        total += parseFloat(item.valor.replace(/\./g, '').replace(',', '.'));
      });
      
      html += `
        <div class="title">Orçamento de Tratamento</div>
        <div class="content">
          <table>
            <thead><tr><th>Procedimento</th><th style="text-align: right;">Valor</th></tr></thead>
            <tbody>
              ${linhasTabela}
              <tr class="total-row"><td>TOTAL ESTIMADO</td><td style="text-align: right; color: #4318FF;">R$ ${total.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td></tr>
            </tbody>
          </table>
          <br>
          <div style="font-size: 13px; color: #666;">* Este orçamento tem validade de 15 dias a partir da data de emissão.</div>
        </div>
      `;
    }

    // 3. Fecha o HTML com as assinaturas
    html += `
          <div class="signature">
            <div class="line"></div>
            <div><strong>Dr(a). Responsável</strong></div>
            <div style="font-size: 13px; color: #666;">CRO: XXXXXX</div>
          </div>
          <div class="footer">Documento gerado eletronicamente pelo Sistema de Gestão Clínica Ramos</div>
        </body>
      </html>
    `;

    // =========================================================
    // LÓGICA DE PASTAS INTELIGENTES (DRIVE) E GERAÇÃO DO PDF
    // =========================================================
    const PASTA_RAIZ_ID = "1jA5yXoaiTJ4ZgfibZFYLVjLpbpZQtq2k"; // <--- ATENÇÃO: COLOQUE SEU ID AQUI!
    const pastaPrincipal = DriveApp.getFolderById(PASTA_RAIZ_ID);
    
    // Acha ou Cria a pasta do Paciente
    let pastaPaciente;
    const buscaPastaPac = pastaPrincipal.getFoldersByName(paciente);
    if (buscaPastaPac.hasNext()) {
      pastaPaciente = buscaPastaPac.next();
    } else {
      pastaPaciente = pastaPrincipal.createFolder(paciente);
    }

    // Acha ou Cria a subpasta por Tipo (Receitas, Atestados, etc)
    const nomesPastas = { "RECEITA": "Receitas", "ATESTADO": "Atestados", "ORCAMENTO": "Orcamentos" };
    const nomeSubpasta = nomesPastas[tipo] || "Outros";
    
    let pastaTipo;
    const buscaPastaTipo = pastaPaciente.getFoldersByName(nomeSubpasta);
    if (buscaPastaTipo.hasNext()) {
      pastaTipo = buscaPastaTipo.next();
    } else {
      pastaTipo = pastaPaciente.createFolder(nomeSubpasta);
    }

    // Converte a obra de arte em PDF e salva na pasta correta
    const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
    const nomeArquivo = `${tipo}_${paciente}_${dataAtual.replace(/\//g, '-')}`;
    blob.setName(nomeArquivo + ".pdf");
    
    const arquivo = pastaTipo.createFile(blob);
    const urlPdf = arquivo.getUrl();

    // =========================================================
    // REGISTRA NO BANCO DE DADOS (Aba DOCUMENTOS)
    // =========================================================
    let abaDocs = ss.getSheetByName("DOCUMENTOS");
    if (!abaDocs) {
      abaDocs = ss.insertSheet("DOCUMENTOS");
      abaDocs.appendRow(["DATA", "PACIENTE", "TIPO", "NOME_ARQUIVO", "URL"]);
    }
    abaDocs.appendRow([dataAtual, paciente, tipo, nomeArquivo, urlPdf]);

    return { success: true, url: urlPdf, msg: "Documento salvo na pasta do paciente!" };

  } catch(e) {
    throw new Error("Erro ao gerar PDF: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// BUSCA DOCUMENTOS (VERSÃO DEFINITIVA)
// ==========================================
function apiGetDocumentosPaciente(nomePaciente) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("DOCUMENTOS");
    if (!aba) return [];
    
    const dados = aba.getDataRange().getValues();
    const docs = [];
    const busca = String(nomePaciente || "").toLowerCase().trim();

    for (let i = 1; i < dados.length; i++) {
      if (!dados[i][1]) continue; 
      const pacienteNaPlanilha = String(dados[i][1]).toLowerCase().trim();
      
      if (pacienteNaPlanilha === busca || pacienteNaPlanilha.includes(busca)) {
        let dataStr = dados[i][0];
        if (dataStr instanceof Date) dataStr = Utilities.formatDate(dataStr, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");

        // MONTANDO O OBJETO QUE VAI PARA A TELA
        docs.push({
          data: dataStr,
          tipo: dados[i][2],
          nome: dados[i][3],
          url: dados[i][4],
          hash: dados[i][5] || "",        // Coluna F (Índice 5)
          status: dados[i][6] || "⏳ PENDENTE" // Coluna G (Índice 6)
        });
      }
    }
    return docs.reverse(); // Mais novos no topo
  } catch (e) {
    console.error("Erro ao buscar docs: " + e.message);
    return []; 
  }
}

// ==========================================
// EXCLUIR DOCUMENTO (PLANILHA + DRIVE)
// ==========================================
function apiExcluirDocumentoPDF(urlDocumento) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("DOCUMENTOS");
    if (!aba) throw new Error("Aba de documentos não encontrada.");

    const dados = aba.getDataRange().getValues();
    let linhaParaApagar = -1;

    // 1. Procura qual linha tem essa exata URL
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][4] === urlDocumento) {
        linhaParaApagar = i + 1;
        break;
      }
    }

    if (linhaParaApagar > -1) {
      aba.deleteRow(linhaParaApagar); // Apaga da Planilha
      
      // 2. Tenta caçar o arquivo no Drive para jogar na lixeira
      try {
        // Extrai o ID do arquivo de dentro do link gigante do Google
        const matchId = urlDocumento.match(/[-\w]{25,}/); 
        if (matchId && matchId[0]) {
          DriveApp.getFileById(matchId[0]).setTrashed(true);
        }
      } catch (e) {
        // Se der erro no Drive (ex: arquivo já foi apagado à mão), ignora e segue a vida
      }
      
      SpreadsheetApp.flush();
      return "Documento excluído com sucesso!";
    } else {
      throw new Error("Documento não encontrado no banco de dados.");
    }
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MOTOR DE GERAÇÃO E ORGANIZAÇÃO - CLINICA RAMOS (VERSÃO ULTRA-BLINDADA)
// ==========================================
function registrarEGerarPDF(dados) {
  // 1. TRAVA DE SEGURANÇA (Evita que o Sheets ignore a gravação)
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // Espera até 15 segundos se a planilha estiver ocupada

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 2. NAVEGAÇÃO NA ESTRUTURA DO DRIVE
    const nomePastaRaiz = "Documentos_Clinica_Ramos";
    let pastasRaiz = DriveApp.getFoldersByName(nomePastaRaiz);
    let pastaRaiz = pastasRaiz.hasNext() ? pastasRaiz.next() : DriveApp.createFolder(nomePastaRaiz);

    const nomePac = dados.paciente ? dados.paciente.trim() : "Paciente_Sem_Nome";
    let pastasPac = pastaRaiz.getFoldersByName(nomePac);
    let pastaPaciente = pastasPac.hasNext() ? pastasPac.next() : pastaRaiz.createFolder(nomePac);

    const tipoDoc = dados.tipo ? dados.tipo.trim() : "Outros";
    let pastasTipo = pastaPaciente.getFoldersByName(tipoDoc);
    let pastaDestino = pastasTipo.hasNext() ? pastasTipo.next() : pastaPaciente.createFolder(tipoDoc);

    // 3. GERAÇÃO DO PDF
    const htmlParaPDF = `
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;700&display=swap');
        body { font-family: 'Plus Jakarta Sans', sans-serif; color: #2B3674; margin: 0; padding: 0; }
        .page { width: 100%; padding: 40px; box-sizing: border-box; }
        .header { text-align: center; border-bottom: 2px solid #4318FF; padding-bottom: 20px; margin-bottom: 30px; }
        .logo { font-size: 26px; font-weight: bold; color: #4318FF; margin: 0; }
        .title { text-align: center; font-size: 20px; font-weight: bold; margin: 30px 0; text-transform: uppercase; }
        .caixa-paciente { background: #F4F7FE; padding: 20px; border-radius: 12px; margin-bottom: 30px; font-size: 14px; }
        .content { font-size: 16px; line-height: 2; text-align: justify; white-space: pre-wrap; min-height: 400px; }
        .signature { margin-top: 50px; text-align: center; }
        .line { width: 250px; border-top: 1px solid #2B3674; margin: 0 auto 10px; }
        .page-break { page-break-after: always; }
      </style>
      ${dados.htmlCompleto}
    `;

    const blob = HtmlService.createHtmlOutput(htmlParaPDF).getAs('application/pdf');
    const nomeArquivo = `${dados.tipo}_${nomePac}_${Utilities.formatDate(new Date(), "GMT-3", "dd-MM-yyyy_HH-mm")}.pdf`;
    
    const arquivo = pastaDestino.createFile(blob).setName(nomeArquivo);
    
    // Tenta compartilhar, mas não trava se a conta for corporativa restrita
    try {
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) { console.log("Acesso público restrito pela organização."); }
    
    const urlArquivo = arquivo.getUrl();

    // 4. REGISTRO NA PLANILHA (COM FORÇAMENTO DE GRAVAÇÃO E STATUS)
    let aba = ss.getSheetByName("DOCUMENTOS") || ss.insertSheet("DOCUMENTOS");
    
    if (aba.getLastRow() === 0) {
      aba.appendRow(["DATA", "PACIENTE", "TIPO", "DETALHES", "LINK DRIVE", "CHAVE_HASH", "STATUS"]);
    }

    const dataHoje = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
    
    // Adiciona a linha salvando o Hash e o Status Pendente!
    aba.appendRow([
      dataHoje, 
      nomePac, 
      dados.tipo, 
      dados.resumo || "Documento gerado", 
      urlArquivo,
      dados.hash || "",   // <--- A Chave Exata!
      "⏳ PENDENTE"       // <--- O Status inicial
    ]);

    SpreadsheetApp.flush(); 
    return { url: urlArquivo, success: true }

  } catch (e) {
    throw new Error("Falha Crítica no Registro: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MÓDULO CONSTRUTOR DE RECEITAS INTELIGENTE
// ==========================================

// 1. Banco de Protocolos (Você pode adicionar quantos quiser aqui)
const protocolosReceita = {
    "Articulação (Dor Neuropática)": { 
        med: "Pregabalina 75 mg", 
        qtd: "30 cp", 
        uso: "Uso Interno", 
        poso: "Tomar 1 cápsula, 1x por dia, durante 30 dias." 
    },
    "Pós-Operatório Padrão": { 
        med: "Ibuprofeno 600mg", 
        qtd: "10 cp", 
        uso: "Uso Interno", 
        poso: "Tomar 1 comprimido a cada 8 horas, em caso de dor." 
    },
    "Infecção Moderada": { 
        med: "Amoxicilina 500mg", 
        qtd: "21 cp", 
        uso: "Uso Interno", 
        poso: "Tomar 1 comprimido a cada 8 horas, durante 7 dias." 
    },
    "Bochecho / Assepsia": { 
        med: "Digluconato de Clorexidina 0,12%", 
        qtd: "1 Frasco (250ml)", 
        uso: "Uso Externo", 
        poso: "Bochechar 15ml por 1 minuto, 2x ao dia, após a escovação." 
    }
};

let contadorMedicamentos = 0;

// 2. Adiciona uma nova linha na tela
function adicionarLinhaMedicamento() {
    contadorMedicamentos++;
    const id = contadorMedicamentos;
    const container = document.getElementById("containerMedicamentosReceita");

    // Monta as opções do Select baseado no banco de protocolos
    let opcoesFinalidade = `<option value="">Selecione a Finalidade...</option>`;
    for (let finalidade in protocolosReceita) {
        opcoesFinalidade += `<option value="${finalidade}">${finalidade}</option>`;
    }
    opcoesFinalidade += `<option value="Livre">Digitação Livre (Outro)</option>`;

    const htmlLinha = `
    <div class="p-3 border rounded shadow-sm bg-white" id="blocoMed_${id}" style="position: relative;">
        <button type="button" class="btn btn-sm btn-danger position-absolute top-0 end-0 m-2 rounded-circle" style="width:28px; height:28px; padding:0;" onclick="removerLinhaMedicamento(${id})">
           <span class="material-icons" style="font-size:16px; line-height:28px;">close</span>
        </button>
        
        <div class="row g-2">
            <div class="col-md-6">
                <label class="form-label small fw-bold text-muted mb-0">Finalidade / Protocolo</label>
                <select id="recFin_${id}" class="form-select form-select-sm" onchange="autoPreencherReceita(${id})">
                    ${opcoesFinalidade}
                </select>
            </div>
            <div class="col-md-6">
                <label class="form-label small fw-bold text-muted mb-0">Medicamento</label>
                <input type="text" id="recMed_${id}" class="form-control form-control-sm">
            </div>
            <div class="col-md-4">
                <label class="form-label small fw-bold text-muted mb-0">Qtd</label>
                <input type="text" id="recQtd_${id}" class="form-control form-control-sm" placeholder="Ex: 30 cp">
            </div>
            <div class="col-md-8">
                <label class="form-label small fw-bold text-muted mb-0">Uso</label>
                <input type="text" id="recUso_${id}" class="form-control form-control-sm" placeholder="Ex: Interno, Externo, Local">
            </div>
            <div class="col-12">
                <label class="form-label small fw-bold text-muted mb-0">Posologia / Instruções</label>
                <input type="text" id="recPoso_${id}" class="form-control form-control-sm">
            </div>
        </div>
    </div>`;

    container.insertAdjacentHTML('beforeend', htmlLinha);
}

// 3. Remove a linha se o usuário desistir
function removerLinhaMedicamento(id) {
    const bloco = document.getElementById(`blocoMed_${id}`);
    if (bloco) bloco.remove();
}

// 4. Auto-Preenche os campos quando seleciona a finalidade
function autoPreencherReceita(id) {
    const finalidade = document.getElementById(`recFin_${id}`).value;
    if (protocolosReceita[finalidade]) {
        const dados = protocolosReceita[finalidade];
        document.getElementById(`recMed_${id}`).value = dados.med;
        document.getElementById(`recQtd_${id}`).value = dados.qtd;
        document.getElementById(`recUso_${id}`).value = dados.uso;
        document.getElementById(`recPoso_${id}`).value = dados.poso;
    } else {
        // Se for digitação livre, limpa para o doutor digitar
        document.getElementById(`recMed_${id}`).value = "";
        document.getElementById(`recQtd_${id}`).value = "";
        document.getElementById(`recUso_${id}`).value = "";
        document.getElementById(`recPoso_${id}`).value = "";
    }
}

// ==========================================
// BANCO DE DADOS DE RECEITAS E PROFISSIONAIS
// ==========================================
function getDadosReceitaERP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Cria ou Lê a aba de Profissionais
  let abaProf = ss.getSheetByName("PROFISSIONAIS");
  if (!abaProf) {
    abaProf = ss.insertSheet("PROFISSIONAIS");
    abaProf.appendRow(["NOME", "REGISTRO_CRO", "ENDERECO_CLINICA"]);
    // Cadastra o primeiro doutor automaticamente
    abaProf.appendRow(["Dra. Giovanna Ramos", "CRO/SC 19773", "Clínica Odontologia Ramos • Palhoça / SC"]);
  }
  
  // 2. Cria ou Lê a aba de Protocolos
  let abaProt = ss.getSheetByName("PROTOCOLOS");
  if (!abaProt) {
    abaProt = ss.insertSheet("PROTOCOLOS");
    abaProt.appendRow(["FINALIDADE", "MEDICAMENTO", "QUANTIDADE", "USO", "POSOLOGIA"]);
    // Cadastra os exemplos base
    abaProt.appendRow(["Articulação (Dor Neuropática)", "Pregabalina 75 mg", "30 cp", "Uso Interno", "Tomar 1 cápsula, 1x por dia, durante 30 dias."]);
    abaProt.appendRow(["Pós-Operatório Padrão", "Ibuprofeno 600mg", "10 cp", "Uso Interno", "Tomar 1 comprimido a cada 8 horas, em caso de dor."]);
    abaProt.appendRow(["Infecção Moderada", "Amoxicilina 500mg", "21 cp", "Uso Interno", "Tomar 1 comprimido a cada 8 horas, durante 7 dias."]);
    abaProt.appendRow(["Bochecho / Assepsia", "Digluconato de Clorexidina 0,12%", "1 Frasco (250ml)", "Uso Externo", "Bochechar 15ml por 1 minuto, 2x ao dia, após a escovação."]);
  }

  // 3. Extrai os dados das abas para enviar ao sistema
  const profs = abaProf.getDataRange().getValues();
  const listaProf = [];
  for (let i = 1; i < profs.length; i++) {
    if (profs[i][0]) listaProf.push({ nome: profs[i][0], cro: profs[i][1], endereco: profs[i][2] });
  }

  const prots = abaProt.getDataRange().getValues();
  const listaProt = {};
  for (let i = 1; i < prots.length; i++) {
    if (prots[i][0]) {
      listaProt[prots[i][0]] = { med: prots[i][1], qtd: prots[i][2], uso: prots[i][3], poso: prots[i][4] };
    }
  }
  
  return { profissionais: listaProf, protocolos: listaProt };
}
// ==========================================
// GESTOR DE DADOS (CRUD) - PROFISSIONAIS E PROTOCOLOS
// ==========================================

function getDadosConfiguracao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    profissionais: buscarDadosAba(ss, "PROFISSIONAIS"),
    protocolos: buscarDadosAba(ss, "PROTOCOLOS")
  };
}

function buscarDadosAba(ss, nomeAba) {
  let aba = ss.getSheetByName(nomeAba);
  if (!aba) return [];
  const dados = aba.getDataRange().getValues();
  const cabecalho = dados.shift();
  return dados.map(linha => {
    let obj = {};
    cabecalho.forEach((col, i) => obj[col.toLowerCase()] = linha[i]);
    return obj;
  });
}

function salvarItemConfig(abaNome, dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(abaNome) || ss.insertSheet(abaNome);
  const valores = aba.getDataRange().getValues();
  const cabecalho = valores[0];
  
  // Se for a primeira vez, cria o cabeçalho
  if (valores.length === 1 && valores[0][0] === "") {
    const novoCabecalho = Object.keys(dados).map(k => k.toUpperCase());
    aba.appendRow(novoCabecalho);
  }

  const novaLinha = cabecalho.map(col => dados[col.toLowerCase()] || "");
  
  // Verifica se é edição (pelo primeiro campo, ex: Nome ou Finalidade)
  let linhaExistente = -1;
  for (let i = 1; i < valores.length; i++) {
    if (valores[i][0] === novaLinha[0]) { linhaExistente = i + 1; break; }
  }

  if (linhaExistente !== -1) {
    aba.getRange(linhaExistente, 1, 1, novaLinha.length).setValues([novaLinha]);
  } else {
    aba.appendRow(novaLinha);
  }
  return "Salvo com sucesso!";
}

function excluirItemConfig(abaNome, identificador) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(abaNome);
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] === identificador) {
      aba.deleteRow(i + 1);
      return "Excluído!";
    }
  }
}

// ==========================================
// MÓDULO DE ENVIO DE DOCUMENTOS (E-MAIL)
// ==========================================
function apiEnviarEmailDocumento(urlDocumento, emailPaciente, nomePaciente, tipoDoc) {
  try {
    if (!emailPaciente || emailPaciente.trim() === "") {
      throw new Error("O paciente não possui e-mail cadastrado na ficha.");
    }

    const assunto = `Seu Documento Odontológico (${tipoDoc}) - Clínica Ramos`;
    const mensagemHtml = `
      <div style="font-family: Arial, sans-serif; color: #2B3674; max-width: 600px; margin: auto; padding: 20px; border: 1px solid #E9EDF7; border-radius: 12px;">
        <h2 style="color: #4318FF;">Olá, ${nomePaciente}!</h2>
        <p>A Clínica Odontológica Ramos está enviando o seu documento (<strong>${tipoDoc}</strong>) gerado em nosso sistema.</p>
        <p>Para visualizar, baixar ou imprimir, basta clicar no botão seguro abaixo:</p>
        <br>
        <div style="text-align: center;">
          <a href="${urlDocumento}" style="background-color: #4318FF; color: white; padding: 12px 25px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px;">📄 ACESSAR MEU DOCUMENTO</a>
        </div>
        <br><br>
        <hr style="border: 0; border-top: 1px solid #E9EDF7;">
        <p style="font-size: 12px; color: #A3AED0;">Este é um e-mail automático. Em caso de dúvidas, entre em contato com a nossa recepção.</p>
      </div>
    `;

    MailApp.sendEmail({
      to: emailPaciente,
      subject: assunto,
      htmlBody: mensagemHtml
    });

    return "E-mail enviado com sucesso para " + emailPaciente;
  } catch (e) {
    throw new Error(e.message);
  }
}

function getDocumentUrl(hash) {
  // Procura o link do PDF na folha DOCUMENTOS usando o hash
  const dados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DOCUMENTOS").getDataRange().getValues();
  for(let i=1; i<dados.length; i++) { if(dados[i][3].includes(hash)) return dados[i][4]; }
  return "#";
}

// ==========================================
// SALVAR ARQUIVO ANEXADO EXTERNAMENTE
// ==========================================
function apiUploadArquivoPaciente(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    const pastaRaiz = DriveApp.getFoldersByName("Documentos_Clinica_Ramos").next();
    let pastaPaciente;
    
    // Procura ou cria a pasta do paciente
    const pastas = pastaRaiz.getFoldersByName(payload.paciente);
    if (pastas.hasNext()) {
      pastaPaciente = pastas.next();
    } else {
      pastaPaciente = pastaRaiz.createFolder(payload.paciente);
    }
    
    // Converte o arquivo e salva no Drive
    const blob = Utilities.newBlob(Utilities.base64Decode(payload.base64), payload.mimeType, payload.nomeArquivo);
    const arquivoCriado = pastaPaciente.createFile(blob);
    
    // Registra na aba DOCUMENTOS para aparecer na tabela do prontuário
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("DOCUMENTOS");
    const dataStr = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");
    
    // Colunas: Data, Paciente, Tipo, Nome do Arquivo, Link, Hash, Status
    aba.appendRow([
      dataStr, 
      payload.paciente, 
      "ANEXO EXTERNO", 
      payload.nomeArquivo, 
      arquivoCriado.getUrl(), 
      "", 
      "✅ ARQUIVADO"
    ]);
    
    return "Sucesso";
  } catch (e) {
    throw new Error("Falha no upload: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MOTOR DO GESTOR FINANCEIRO (BACK-END)
// ==========================================

const ABA_FINANCEIRO = "FINANCEIRO"; // Nome da aba na sua planilha

/**
 * Salva um movimento (Entrada ou Saída) no Fluxo de Caixa
 */
// ==========================================
// SALVAR MOVIMENTO FINANCEIRO (AGORA COM PAGAMENTO E DOUTOR)
// ==========================================
function apiSalvarMovimentoFinanceiro(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("FINANCEIRO");
    
    if (!sheet) {
      sheet = ss.insertSheet("FINANCEIRO");
      sheet.appendRow(["ID", "DATA", "TIPO", "CATEGORIA", "DESCRIÇÃO", "VALOR BRUTO", "DESCONTOS", "VALOR LÍQUIDO", "STATUS", "ORIGEM_ID", "PAGAMENTO", "DOUTOR"]);
    }

    const limparN = (v) => {
       if (typeof v === 'number') return v;
       return Number((v || "0").toString().replace(/\./g, '').replace(',', '.')) || 0;
    };
    
    const bruto = limparN(dados.valorBruto);
    const descontos = limparN(dados.descontos);
    const liquido = bruto - descontos;
    const idSalvar = dados.id || "FIN-" + new Date().getTime();

    const novaLinha = [
      idSalvar,
      dados.data, 
      dados.tipo, 
      dados.categoria,
      dados.descricao,
      bruto,
      descontos,
      liquido,
      dados.status || "PAGO",
      dados.origemId || "",
      dados.pagamento || "Não Informado", // GRAVA A FORMA DE PAGAMENTO
      dados.doutor || ""                  // GRAVA O DOUTOR
    ];

    const busca = sheet.getRange("A:A").createTextFinder(idSalvar).matchEntireCell(true).findNext();
    
    if (busca) {
      const linha = busca.getRow();
      sheet.getRange(linha, 1, 1, novaLinha.length).setValues([novaLinha]);
      return "OK_EDIT";
    } else {
      sheet.appendRow(novaLinha);
      return "OK_NOVO";
    }
  } catch (e) {
    return "ERRO: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// BUSCADOR AVANÇADO COM FILTRO DE DATAS, PAGAMENTOS E COMISSÕES
// ==========================================
function getDadosFinanceiroFiltrado(dtInicioStr, dtFimStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaFin = ss.getSheetByName("FINANCEIRO");
  const abaTotal = ss.getSheetByName("TOTAL");
  
  let result = { fluxo: [], pagamentos: {}, comissoes: {}, totalE: 0, totalS: 0 };

  // Tratamento Flexível de Datas
  const hoje = new Date();
  let inicio = dtInicioStr ? new Date(dtInicioStr + "T00:00:00") : new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  let fim = dtFimStr ? new Date(dtFimStr + "T23:59:59") : new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0, 23, 59, 59);

  // 1. DADOS DE FLUXO DE CAIXA E PAGAMENTOS
  if (abaFin) {
    const dadosFin = abaFin.getDataRange().getValues();
    for (let i = 1; i < dadosFin.length; i++) {
      const r = dadosFin[i];
      let dataFiltro = null;
      let dataFormatada = "";

      if (r[1] instanceof Date) {
        dataFiltro = r[1];
        dataFormatada = Utilities.formatDate(r[1], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
      } else if (typeof r[1] === 'string' && r[1].includes('-')) {
        const pts = r[1].split('-');
        dataFiltro = new Date(pts[0], pts[1]-1, pts[2]);
        dataFormatada = `${pts[2]}/${pts[1]}/${pts[0]}`;
      } else if (typeof r[1] === 'string' && r[1].includes('/')) {
        const pts = r[1].split('/');
        dataFiltro = new Date(pts[2], pts[1]-1, pts[0]);
        dataFormatada = r[1];
      }

      // Se estiver dentro do intervalo de tempo selecionado
      if (dataFiltro && dataFiltro >= inicio && dataFiltro <= fim) {
        const tipo = r[2];
        const liquido = Number(r[7]) || 0;
        const pagamento = r[10] || "Não Informado";
        
        result.fluxo.push({
          id: r[0], data: dataFormatada, tipo: tipo, categoria: r[3],
          descricao: r[4], bruto: r[5], descontos: r[6], liquido: liquido, status: r[8]
        });

        if (tipo === "ENTRADA") {
          result.totalE += liquido;
          if (r[3] === "PARTICULAR" || r[3] === "PLANO") {
             result.pagamentos[pagamento] = (result.pagamentos[pagamento] || 0) + liquido;
          }
        } else {
          result.totalS += liquido;
        }
      }
    }
  }

  // 2. DADOS DE COMISSÃO DE DOUTORES (Vindo direto da Produção)
  if (abaTotal) {
    const dadosTotal = abaTotal.getDataRange().getValues();
    for (let i = 2; i < dadosTotal.length; i++) {
      const r = dadosTotal[i];
      let dataFiltro = null;

      if (r[0] instanceof Date) {
        dataFiltro = r[0];
      } else if (typeof r[0] === 'string' && r[0].includes('/')) {
        const pts = r[0].split('/');
        dataFiltro = new Date(pts[2], pts[1]-1, pts[0]);
      }

      if (dataFiltro && dataFiltro >= inicio && dataFiltro <= fim) {
        const doutor = r[1];
        const comissao = (Number(r[10]) || 0) + (Number(r[11]) || 0); // Soma Parte Clínica/Cirurgia
        if (doutor && comissao > 0) {
          result.comissoes[doutor] = (result.comissoes[doutor] || 0) + comissao;
        }
      }
    }
  }

  result.fluxo.reverse(); // Coloca os mais recentes no topo
  return result;
}

/**
 * Busca os dados do financeiro para exibir no ERP
 */
function getDadosFinanceiroMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ABA_FINANCEIRO);
  if (!sheet) return [];

  const dados = sheet.getDataRange().getValues();
  if (dados.length < 2) return [];

  const hoje = new Date();
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();

  const resultado = [];

  // Pula o cabeçalho (i=1)
  for (let i = 1; i < dados.length; i++) {
    const r = dados[i];
    const dataRow = r[1]; // Coluna DATA
    
    // Tenta converter o valor da coluna B em uma data válida para filtrar o mês
    let dataFormatada = "";
    let dataFiltro = null;

    if (dataRow instanceof Date) {
      dataFiltro = dataRow;
      dataFormatada = Utilities.formatDate(dataRow, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    } else if (typeof dataRow === 'string' && dataRow.includes('-')) {
      const pts = dataRow.split('-');
      dataFiltro = new Date(pts[0], pts[1]-1, pts[2]);
      dataFormatada = `${pts[2]}/${pts[1]}/${pts[0]}`;
    } else {
      dataFormatada = dataRow.toString();
    }

    // Filtro Opcional: Se quiser mostrar apenas o mês atual na tabela
    // if (dataFiltro && (dataFiltro.getMonth() !== mesAtual || dataFiltro.getFullYear() !== anoAtual)) continue;

    resultado.push({
      id: r[0],
      data: dataFormatada,
      tipo: r[2],
      categoria: r[3],
      descricao: r[4],
      bruto: r[5],
      descontos: r[6],
      liquido: r[7],
      status: r[8]
    });
  }

  // Retorna os lançamentos mais recentes primeiro
  return resultado.reverse();
}

function getResumoDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaFin = ss.getSheetByName("FINANCEIRO");
  const abaDados = ss.getSheetByName("DADOS");
  
  let resumo = { totalLab: 0, totalFixas: 0, totalEntradas: 0 };
  
  if (abaFin) {
    const dadosFin = abaFin.getDataRange().getValues();
    for(let i=1; i<dadosFin.length; i++) {
      if(dadosFin[i][2] === "ENTRADA") resumo.totalEntradas += Number(dadosFin[i][7]);
      if(dadosFin[i][2] === "SAÍDA") resumo.totalFixas += Number(dadosFin[i][7]);
    }
  }
  
  if (abaDados) {
    const dadosProd = abaDados.getDataRange().getValues();
    for(let i=1; i<dadosProd.length; i++) {
      resumo.totalLab += Number(dadosProd[i][9]); // Coluna J (Valor Lab)
    }
  }
  
  return resumo;
}

// ==========================================
// MOTOR DE RECORRÊNCIA E PARCELAMENTO
// ==========================================

/**
 * Função que o Google vai rodar automaticamente todo dia
 */
function gatilhoProcessarRecorrencias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaConfig = ss.getSheetByName("CONFIG_FIXOS");
  const hoje = new Date();
  const diaHoje = hoje.getDate();
  const mesAnoAtual = Utilities.formatDate(hoje, ss.getSpreadsheetTimeZone(), "MM/yyyy");

  if (!abaConfig) return;
  const dados = abaConfig.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    let [id, cat, desc, valor, dia, parcelas, ultimo] = dados[i];
    
    // Se o dia chegou e ainda não foi lançado este mês
    if (diaHoje >= dia && ultimo !== mesAnoAtual && (parcelas > 0 || parcelas === -1)) {
      
      // 1. Lança no Financeiro
      const dadosLancamento = {
        tipo: "SAÍDA",
        data: Utilities.formatDate(hoje, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd"),
        categoria: cat,
        descricao: parcelas > 0 ? `${desc} (Parc. ${parcelas} restantes)` : desc,
        valorBruto: valor,
        descontos: 0,
        status: "PENDENTE",
        origemId: "FIXO-" + id
      };
      
      apiSalvarMovimentoFinanceiro(dadosLancamento);

      // 2. Atualiza a regra de recorrência
      if (parcelas > 0) parcelas--; // Se for parcelado, subtrai uma
      
      abaConfig.getRange(i + 1, 6).setValue(parcelas); // Atualiza parcelas restantes
      abaConfig.getRange(i + 1, 7).setValue(mesAnoAtual); // Marca como lançado no mês
    }
  }
}

/**
 * Cadastra uma nova regra (Fixo ou Parcelado)
 */
function salvarRegraFixa(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("CONFIG_FIXOS");
  const id = "REG-" + new Date().getTime();
  
  // dados.parcelas: -1 para fixo, X para parcelado
  aba.appendRow([id, dados.categoria, dados.descricao, dados.valor, dados.dia, dados.parcelas, ""]);
  return "Regra configurada com sucesso!";
}

/**
 * Busca todas as regras de custos fixos/parcelados
 */
function getRegrasFixas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName("CONFIG_FIXOS");
  
  // Se a aba não existir, cria para não dar erro de "null"
  if (!aba) {
    aba = ss.insertSheet("CONFIG_FIXOS");
    aba.appendRow(["ID", "CATEGORIA", "DESCRICAO", "VALOR", "DIA", "PARCELAS_RESTANTES", "ULTIMO_LANCAMENTO"]);
    return [];
  }
  
  const dados = aba.getDataRange().getValues();
  if (dados.length < 2) return [];

  // Mapeamento explícito das colunas para evitar erro de índice
  return dados.slice(1).map(r => {
    return {
      id: r[0] ? r[0].toString() : "",
      categoria: r[1] ? r[1].toString() : "",
      descricao: r[2] ? r[2].toString() : "",
      valor: r[3] ? Number(r[3]) : 0,
      dia: r[4] ? r[4].toString() : "",
      parcelas: r[5] !== "" ? Number(r[5]) : -1,
      ultimo: r[6] ? r[6].toString() : ""
    };
  }).filter(item => item.id !== ""); // Remove linhas vazias acidentais
}

function getRegistroFinanceiroPorId(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("FINANCEIRO");
  const dados = sheet.getDataRange().getValues();
  
  for(let i=1; i<dados.length; i++) {
    if(dados[i][0] == id) {
      let dataOriginal = dados[i][1];
      let dataISO = dataOriginal instanceof Date ? Utilities.formatDate(dataOriginal, "GMT-3", "yyyy-MM-dd") : "";
      
      return {
        id: dados[i][0],
        dataISO: dataISO,
        tipo: dados[i][2],
        categoria: dados[i][3],
        descricao: dados[i][4],
        valorBR: dados[i][5].toLocaleString('pt-BR', {minimumFractionDigits: 2})
      };
    }
  }
  return null;
}

/**
 * Remove uma regra de recorrência
 */
function excluirRegraFixa(idRegra) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("CONFIG_FIXOS");
  const busca = aba.getRange("A:A").createTextFinder(idRegra).matchEntireCell(true).findNext();
  
  if (busca) {
    aba.deleteRow(busca.getRow());
    return "Regra removida!";
  }
  return "Erro: Regra não encontrada.";
}

/**
 * Exclui um lançamento do Fluxo de Caixa (Aba FINANCEIRO)
 */
function excluirRegistroFinanceiroBD(idRegistro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("FINANCEIRO");
  if (!sheet) return "Erro: Aba Financeiro não encontrada.";
  
  const busca = sheet.getRange("A:A").createTextFinder(idRegistro.toString()).matchEntireCell(true).findNext();
  if (busca) {
    sheet.deleteRow(busca.getRow());
    SpreadsheetApp.flush();
    return "Lançamento excluído com sucesso!";
  }
  return "Erro: Registro não encontrado.";
}

/**
 * Edita uma regra de recorrência (Aba CONFIG_FIXOS)
 */
function editarRegraFixaBD(id, desc, valor, dia) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("CONFIG_FIXOS");
    
    const busca = aba.getRange("A:A").createTextFinder(id.toString()).matchEntireCell(true).findNext();
    
    if (busca) {
      const linha = busca.getRow();
      const valorNum = Number(valor.toString().replace(/\./g, '').replace(',', '.')) || 0;
      
      aba.getRange(linha, 3).setValue(desc);    // Atualiza Descrição
      aba.getRange(linha, 4).setValue(valorNum); // Atualiza Valor
      aba.getRange(linha, 5).setValue(dia);     // Atualiza Dia do Vencimento
      
      SpreadsheetApp.flush();
      return "✅ Recorrência atualizada!";
    }
    return "❌ Erro: Regra não encontrada.";
  } finally {
    lock.releaseLock();
  }
}
