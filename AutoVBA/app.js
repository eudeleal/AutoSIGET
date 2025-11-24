function rangeVazio(sheet, rangeA1) { 
  const valores = sheet.getRange(rangeA1).getValues();

  // Percorre todas as linhas e colunas
  for (let i = 0; i < valores.length; i++) {
    for (let j = 0; j < valores[i].length; j++) {
      if (valores[i][j] === "" || valores[i][j] === null) {
        return true; // Encontrou alguma célula vazia
      }
    }
  }
  return false; // Todas as células estão preenchidas
}

function ADD_Linha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaCSV = ss.getSheetByName("CSV");
  const abaAdd_Linhas = ss.getSheetByName("ADD_Linhas");
  
  // Cabeçalho fixo desejado
  const cabecalhoCSV = [
    ["FaixaInicio", "FaixaFinal", "Intervalo", "Percurso", "TempTerm", "Frota", "Linha", "Dia", "Sentido", "Oso", "LinhaOso"]
  ];

  // Sempre sobrescreve o cabeçalho
  abaCSV.getRange(1, 1, 1, cabecalhoCSV[0].length).setValues(cabecalhoCSV);

  // Verifica campos obrigatórios
  if (rangeVazio(abaAdd_Linhas, "H2:J2")) {
    SpreadsheetApp.getUi().alert("Faltam informações das faixas.");
    return;
  }

  const dados = abaAdd_Linhas.getDataRange().getValues();
  if (dados.length < 2) {
    SpreadsheetApp.getUi().alert("Nenhuma faixa preenchida.");
    return;
  }

  const linhas = dados.slice(1); // Remove o cabeçalho original

  // Lê dados fixos da primeira linha
  const linhaBase = linhas[0];
  const linhaValor = linhaBase[7];   // Coluna H
  const diaValor = linhaBase[8];     // Coluna I
  const sentidoBruto = linhaBase[9]; // Coluna J

  // Converte SENTIDO para número
  let sentidoValor = "";
  if (typeof sentidoBruto === "string") {
    if (sentidoBruto.includes("0")) sentidoValor = "0";
    else if (sentidoBruto.includes("1")) sentidoValor = "1";
    else if (sentidoBruto.includes("2")) sentidoValor = "2";
    else sentidoValor = sentidoBruto.trim();
  } else {
    sentidoValor = sentidoBruto;
  }

  // Função para formatar hora como HHMM
  function formatarHora(valor) {
    if (!valor) return "";

    if (Object.prototype.toString.call(valor) === "[object Date]") {
      return Utilities.formatDate(valor, ss.getSpreadsheetTimeZone(), "HHmm");
    }

    const str = String(valor).trim();
    const regex = /^(\d{1,2}):(\d{2})/;
    const match = str.match(regex);
    if (match) {
      return match[1].padStart(2, "0") + match[2];
    }

    if (!isNaN(valor)) {
      return String(valor).padStart(4, "0");
    }

    return "";
  }

  // Monta as novas linhas
  const novasLinhas = linhas
    .filter(l => l[0] && l[1]) // Ignora linhas vazias
    .map(l => [
      formatarHora(l[0]), // FaixaInicio
      formatarHora(l[1]), // FaixaFinal
      l[2],               // Intervalo
      l[3],               // Percurso
      l[4],               // TempTerm
      l[5],               // Frota
      linhaValor,         // Linha
      diaValor,           // Dia
      sentidoValor,       // Sentido
      "",                 // Oso
      "",                 // LinhaOso
    ]);

  if (novasLinhas.length === 0) {
    SpreadsheetApp.getUi().alert("Nenhuma linha válida encontrada.");
    return;
  }

  // Adiciona após o cabeçalho
  const ultimaLinha = abaCSV.getLastRow();
  abaCSV.getRange(ultimaLinha + 1, 1, novasLinhas.length, novasLinhas[0].length).setValues(novasLinhas);

  // Limpa a aba ADD (mantendo o cabeçalho)
  abaAdd_Linhas.getRange(2, 1, abaAdd_Linhas.getLastRow() - 1, abaAdd_Linhas.getLastColumn()).clearContent();

  SpreadsheetApp.getUi().alert("✅ Linhas adicionadas à aba CSV com sucesso!");
}

function ADD_OSO() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaAdd = ss.getSheetByName("ADD_OSOs");
  const abaCSV = ss.getSheetByName("CSV");

  if (!abaAdd || !abaCSV) {
    SpreadsheetApp.getUi().alert("❌ As abas 'ADD_OSOs' ou 'CSV' não foram encontradas.");
    return;
  }

  // Lê valores da aba ADD_OSOs (ignorando o cabeçalho)
  const dados = abaAdd.getRange(2, 1, abaAdd.getLastRow() - 1, 2).getValues();

  // Filtra linhas válidas
  const novos = dados.filter(r => r[0] && r[1]); // OSO e Linha preenchidos
  if (novos.length === 0) {
    SpreadsheetApp.getUi().alert("⚠️ Nenhum dado válido encontrado para adicionar.");
    return;
  }

  // Validação de formato (OSO = 6 dígitos | Linha = 4 ou 6 dígitos)
  const formatoOSO = /^\d{6}$/;
  const formatoLinha = /^\d{4}(\d{2})?$/;

  for (let i = 0; i < novos.length; i++) {
    const [oso, linha] = novos[i];

    if (!formatoOSO.test(oso)) {
      SpreadsheetApp.getUi().alert(`❌ OSO inválido na linha ${i + 2}: "${oso}" (esperado 6 dígitos)`);
      return;
    }

    if (!formatoLinha.test(linha)) {
      SpreadsheetApp.getUi().alert(`❌ Linha inválida na linha ${i + 2}: "${linha}" (esperado 4 ou 6 dígitos)`);
      return;
    }
  }

  // Localiza a última linha com OSO existente na aba CSV
  const ultimaLinhaCSV = abaCSV.getLastRow();
  let inicioNova = 2; // após cabeçalho

  // Se já existem OSOs, adiciona abaixo
  if (ultimaLinhaCSV > 1) {
    const valoresCSV = abaCSV.getRange(2, 10, ultimaLinhaCSV - 1, 2).getValues();
    const ultimasNaoVazias = valoresCSV.filter(r => r[0] || r[1]);
    inicioNova = ultimasNaoVazias.length + 2;
  }

  // Escreve os novos dados nas colunas J e K (10 e 11)
  abaCSV.getRange(inicioNova, 10, novos.length, 2).setValues(novos);

  // Limpa apenas o conteúdo digitado em ADD_OSOs (mantém cabeçalho)
  abaAdd.getRange(2, 1, abaAdd.getLastRow() - 1, 2).clearContent();

  SpreadsheetApp.getUi().alert(`✅ ${novos.length} OSO(s) adicionada(s) com sucesso à aba CSV!`);
}

function baixarCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("CSV");

  if (!aba) {
    SpreadsheetApp.getUi().alert("A aba 'CSV' não existe!");
    return;
  }

  const dados = aba.getDataRange().getValues();
  if (dados.length < 2) {
    SpreadsheetApp.getUi().alert("Aba CSV está vazia.");
    return;
  }

  // IDs necessários
  const spreadsheetId = ss.getId();
  const sheetId = aba.getSheetId();

  // URL para exportar diretamente como CSV
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=${sheetId}`;

  // Abre o link de download em nova janela
  const html = HtmlService.createHtmlOutput(
    `<script>window.open("${url}");google.script.host.close();</script>`
  );
  SpreadsheetApp.getUi().showModalDialog(html, "Exportando CSV...");
}

function Limpar_OSO() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaCSV = ss.getSheetByName("CSV");

  if (!abaCSV) {
    SpreadsheetApp.getUi().alert("❌ A aba 'CSV' não foi encontrada.");
    return;
  }

  const ultimaLinha = abaCSV.getLastRow();
  if (ultimaLinha < 2) {
    SpreadsheetApp.getUi().alert("A aba CSV não contém dados além do cabeçalho.");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const resposta = ui.alert(
    "Confirmação",
    "Deseja realmente limpar as OSOs?",
    ui.ButtonSet.YES_NO
  );

  if (resposta !== ui.Button.YES) {
    ui.alert("Operação cancelada.");
    return;
  }

  // Limpa apenas colunas 10 e 11 (Oso e LinhaOso), mantendo o cabeçalho
  abaCSV.getRange(2, 10, ultimaLinha - 1, 2).clearContent();

  ui.alert("✅ Colunas OSO e LinhaOso limpas com sucesso!");
}

function Limpar_Linha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaCSV = ss.getSheetByName("CSV");

  if (!abaCSV) {
    SpreadsheetApp.getUi().alert("❌ A aba 'CSV' não foi encontrada.");
    return;
  }

  const ultimaLinha = abaCSV.getLastRow();
  if (ultimaLinha < 2) {
    SpreadsheetApp.getUi().alert("A aba CSV não contém dados além do cabeçalho.");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const resposta = ui.alert(
    "Confirmação",
    "Deseja realmente limpar as FAIXAS e demais dados?",
    ui.ButtonSet.YES_NO
  );

  if (resposta !== ui.Button.YES) {
    ui.alert("Operação cancelada.");
    return;
  }

  // Limpa colunas 1 a 9 (Faixas, Intervalo, Frota, etc.), mantendo OSO e LinhaOso
  abaCSV.getRange(2, 1, ultimaLinha - 1, 9).clearContent();

  ui.alert("✅ Colunas de Linhas e Faixas limpas com sucesso!");
}