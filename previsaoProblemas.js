function gerarPoisson(lambda) {
  let L = Math.exp(-lambda);
  let p = 1;
  let k = 0;

  do {
    k++;
    p *= Math.random();
  } while (p > L);

  return k - 1;
}

function calcularMonteCarlo() {
  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("RelatórioFichas");

  const dados = sheet.getRange("K12:M18").getValues();
  // K = problema | M = média semanal (λ)

  const semanasFuturas = 4;
  const simulacoes = 10000;
  const limiteCritico = 15;

  let resultado = [];

  dados.forEach(linha => {
    const problema = linha[0];
    const lambda = Number(linha[2]); // coluna M

    if (!problema || !lambda) return;

    let ultrapassou = 0;

    for (let i = 0; i < simulacoes; i++) {
      let totalPeriodo = 0;

      for (let s = 0; s < semanasFuturas; s++) {
        totalPeriodo += gerarPoisson(lambda);
      }

      if (totalPeriodo >= limiteCritico) ultrapassou++;
    }

    resultado.push({
      problema: problema,
      probabilidade: ultrapassou / simulacoes
    });
  });

  return resultado; // 🔥 ESSENCIAL
}


function abrirPopupMonteCarlo() {
  const html = HtmlService
    .createHtmlOutputFromFile("popupMonteCarlo")
    .setWidth(650)
    .setHeight(500);

  SpreadsheetApp.getUi()
    .showModalDialog(html, "Previsão de Risco ");
}
