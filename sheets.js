function adicionarAoCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaAdd = ss.getSheetByName("ADD");
  const abaCSV = ss.getSheetByName("CSV");
  
  const dados = abaAdd.getDataRange().getValues();
  if (dados.length < 2) {
    SpreadsheetApp.getUi().alert("Nenhum dado encontrado na aba ADD!");
    return;
  }

  const linhas = dados.slice(1); // remove o cabeçalho original

  // Lê dados fixos da primeira linha
  const linhaBase = linhas[0];
  const linhaValor = linhaBase[7];   // LINHA (coluna H)
  const diaValor = linhaBase[8];     // DIA (coluna I)
  const sentidoBruto = linhaBase[9]; // SENTIDO (coluna J)

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

  // Função robusta para extrair hora como HHMM
  function formatarHora(valor) {
    if (!valor) return "";

    // Caso seja um objeto Date
    if (Object.prototype.toString.call(valor) === "[object Date]") {
      return Utilities.formatDate(valor, ss.getSpreadsheetTimeZone(), "HHmm");
    }

    // Caso seja texto (tipo "08:00:00" ou "8:00")
    let str = String(valor).trim();
    const regex = /^(\d{1,2}):(\d{2})/;
    const match = str.match(regex);
    if (match) {
      return match[1].padStart(2, "0") + match[2];
    }

    // Caso já esteja em número (tipo 800 ou 0800)
    if (!isNaN(valor)) {
      return String(valor).padStart(4, "0");
    }

    return "";
  }

  // Monta as linhas finais
  const novasLinhas = linhas
    .filter(l => l[0] && l[1]) // ignora linhas vazias
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
    ]);

  if (novasLinhas.length === 0) {
    SpreadsheetApp.getUi().alert("Nenhuma linha válida encontrada.");
    return;
  }

  // Cabeçalho desejado
  const cabecalhoCSV = [
    ["FaixaInicio", "FaixaFinal", "Intervalo", "Percurso", "TempTerm", "Frota", "Linha", "Dia", "Sentido"]
  ];

  // Se a aba CSV estiver vazia, adiciona o cabeçalho
  if (abaCSV.getLastRow() === 0) {
    abaCSV.getRange(1, 1, 1, cabecalhoCSV[0].length).setValues(cabecalhoCSV);
  }

  // Adiciona os dados após o cabeçalho existente
  const ultimaLinha = abaCSV.getLastRow();
  abaCSV.getRange(ultimaLinha + 1, 1, novasLinhas.length, novasLinhas[0].length).setValues(novasLinhas);

  // Limpa a aba ADD (mantendo o cabeçalho)
  abaAdd.getRange(2, 1, abaAdd.getLastRow() - 1, abaAdd.getLastColumn()).clearContent();

  SpreadsheetApp.getUi().alert("✅ Dados adicionados à aba CSV com sucesso!");
}