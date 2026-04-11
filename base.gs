const NOME_ABA_DADOS = "TOTAL"; 


function doGet(e) {
  var pagina = e?.parameter?.page || 'ModuloERP';

  // Mapeamento para o portal do fornecedor
  if (pagina === 'cotacao') { pagina = 'CotacaoFornecedor'; }

  try {
    var template = HtmlService.createTemplateFromFile(pagina);
    return template.evaluate()
      .setTitle('Clínica Ramos')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (erro) {
    console.error("Erro ao carregar página: " + pagina, erro);
    return HtmlService.createTemplateFromFile('ModuloERP').evaluate()
      .setTitle('Clínica Ramos - Erro de Carregamento')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
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
    const sheet = ss.getSheetByName(NOME_ABA_DADOS);
    if (!sheet) return "ERRO: Aba '" + NOME_ABA_DADOS + "' não encontrada";
    
    let dataObj = new Date();
    if (dados.data && dados.data.includes('/')) {
      const partes = dados.data.split("/");
      dataObj = new Date(partes[2], partes[1] - 1, partes[0], 12, 0, 0);
    }

    const valorParticular = Number(dados.particular) || 0;
    const valorPlano = Number(dados.vplano) || 0;
    const valorLab = Number(dados.vlab) || 0;
    
    // EDIÇÃO
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
        return "Registro atualizado com sucesso!";
      }
      return "Registro não encontrado para edição";
    }
    
    // NOVO REGISTRO
    const ultimaLinha = sheet.getLastRow();
    let linhaDestino = ultimaLinha + 1;
    const inicioBusca = Math.max(3, ultimaLinha - 100);
    
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
    
    sheet.getRange(linhaDestino, 1, 1, 10).setValues(novaLinhaDados);
    sheet.getRange(linhaDestino, 14).setValue(novoId);
    
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

// ===== EXCLUIR REGISTRO =====
function excluirLancamento(idLinha) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(NOME_ABA_DADOS);
    const sheetLog = ss.getSheetByName("LOG_EXCLUSOES") || ss.insertSheet("LOG_EXCLUSOES");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 2; i < data.length; i++) {
      if (data[i][13] == idLinha) {
        const linha = i + 1;
        sheetLog.appendRow(data[i]);
        sheet.getRange(linha, 1, 1, 10).clearContent();
        sheet.getRange(linha, 14).clearContent();
        return "Registro excluído!";
      }
    }
    return "Registro não encontrado.";
  } catch (erro) {
    return "Erro ao excluir: " + erro.toString();
  } finally {
    lock.releaseLock();
  }
}

function desfazerUltimaExclusao() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLog = ss.getSheetByName("LOG_EXCLUSOES");
    const sheetDestino = ss.getSheetByName(NOME_ABA_DADOS);
    if (!sheetLog || sheetLog.getLastRow() < 1) return "Nada para desfazer.";
    const lastRow = sheetLog.getLastRow();
    const dados = sheetLog.getRange(lastRow, 1, 1, 14).getValues()[0];
    const linhaDestino = encontrarPrimeiraLinhaVazia(sheetDestino);
    for (let j = 0; j < 14; j++) {
      sheetDestino.getRange(linhaDestino, j + 1).setValue(dados[j]);
    }
    if (linhaDestino > 3) {
      copiarFormatacaoEFormulas(sheetDestino, linhaDestino - 1, linhaDestino);
    }
    sheetLog.deleteRow(lastRow);
    return "Última exclusão recuperada!";
  } catch (erro) {
    return "Erro ao desfazer: " + erro.toString();
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
  // Coloque aqui o nome exato da sua aba de pacientes (ex: "PACIENTES" ou "CRM")
  const aba = ss.getSheetByName("PACIENTES"); 
  if (!aba) return [];

  const dados = aba.getDataRange().getValues();
  if (dados.length < 2) return []; // Se só tiver o cabeçalho, retorna vazio

  // Lê o cabeçalho e transforma tudo em minúsculo para não dar erro de digitação
  const cabecalho = dados[0].map(c => c.toString().toLowerCase().trim());

  // O sistema "caça" a posição exata de cada coluna pelo nome dela!
  const iNome = cabecalho.indexOf("nome");
  const iCpf = cabecalho.indexOf("cpf");
  const iRg = cabecalho.indexOf("rg");
  const iNascimento = cabecalho.indexOf("nascimento");
  const iCelular = cabecalho.indexOf("celular");
  const iTelefone = cabecalho.indexOf("telefone");
  const iEmail = cabecalho.indexOf("email");
  const iConvenio = cabecalho.indexOf("convenio");
  const iEndereco = cabecalho.indexOf("endereco");
  const iProfissao = cabecalho.indexOf("profissao");
  const iSexo = cabecalho.indexOf("sexo");
  const iResponsavel = cabecalho.indexOf("responsavel");
  
  // Caça o Estado Civil (seja como 'estado civil' ou 'estadocivil')
  let iEstadoCivil = cabecalho.indexOf("estado civil");
  if (iEstadoCivil === -1) iEstadoCivil = cabecalho.indexOf("estadocivil");

  const pacientes = [];
  
  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    if (!linha[iNome]) continue; // Pula linhas em branco

    // Tratamento especial para a data de nascimento (evita que a data quebre o sistema)
    let dtNasc = linha[iNascimento];
    if (dtNasc instanceof Date) {
       dtNasc = Utilities.formatDate(dtNasc, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    }

    pacientes.push({
      nome: linha[iNome] ? linha[iNome].toString() : "",
      cpf: iCpf >= 0 ? linha[iCpf].toString() : "",
      rg: iRg >= 0 ? linha[iRg].toString() : "",
      nascimento: dtNasc ? dtNasc.toString() : "",
      celular: iCelular >= 0 ? linha[iCelular].toString() : "",
      telefone: iTelefone >= 0 ? linha[iTelefone].toString() : "",
      email: iEmail >= 0 ? linha[iEmail].toString() : "",
      convenio: iConvenio >= 0 ? linha[iConvenio].toString() : "Particular",
      endereco: iEndereco >= 0 ? linha[iEndereco].toString() : "",
      profissao: iProfissao >= 0 ? linha[iProfissao].toString() : "",
      sexo: iSexo >= 0 ? linha[iSexo].toString() : "",
      responsavel: iResponsavel >= 0 ? linha[iResponsavel].toString() : "",
      estado_civil: iEstadoCivil >= 0 ? linha[iEstadoCivil].toString() : ""
    });
  }
  
  return pacientes;
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
// SALVAR AGENDAMENTO COM TRAVA
// ==========================================
function salvarAgendamentoNaPlanilha(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    // === A MÁGICA DA DATA BRASILEIRA AQUI ===
    if (d.data && d.data.includes('-')) {
        const [ano, mes, dia] = d.data.split('-');
        d.data = `${dia}/${mes}/${ano}`;
    }
    // ========================================

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("AGENDA");
    
    // Se não tem ID, é novo. Se tem, é edição.
    if (!d.id) {
      aba.appendRow([Date.now(), d.data, d.hora, d.paciente, "", d.doutor, d.procedimento, d.status, d.obs]);
    } else {
      // Lógica de edição
      const dados = aba.getDataRange().getValues();
      for (let i = 1; i < dados.length; i++) {
        if (dados[i][0] == d.id) {
          aba.getRange(i + 1, 1, 1, 9).setValues([[d.id, d.data, d.hora, d.paciente, "", d.doutor, d.procedimento, d.status, d.obs]]);
          break;
        }
      }
    }
    
    SpreadsheetApp.flush(); // Garante o salvamento físico
    return "OK";
    
  } catch (e) {
    throw new Error("A agenda está sendo atualizada por outra pessoa. Tente novamente em alguns segundos.");
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
