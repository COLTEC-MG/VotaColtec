function gravaCliques(emailEleitor,linkPlanilha) {
 
  // Abre a planilha e a página com os dados de votação. Procura a linha correspondente ao e-mail do usuário.
  var sheet = SpreadsheetApp.openByUrl(linkPlanilha);
  var sheetPage = SpreadsheetApp.openByUrl(linkPlanilha).getSheetByName('DadosCliques');
  var procurar = sheetPage.createTextFinder(emailEleitor);
  var linha = procurar.findNext();
  var ultimaLinha = sheetPage.getLastRow();
  
  // Caso encontre a linha do usuário.
  if (linha != null) {
    var linhaUser = linha.getRow();
    var nCliques = sheetPage.getRange(linhaUser,4).getValue(); // Pega o número de cliques já realizados.
    var newCliques = nCliques + 1; // Soma mais um clique.
    sheetPage.getRange(linhaUser,4).setValue(newCliques); // Anota o número de cliques realizados por ele.
  }
  // Caso não se encontre a linha do usuário.
  else {
    sheetPage.getRange(ultimaLinha+1,2).setValue(emailEleitor); // Grava o e-mail do usuário na primeira linha sem dados.
    var nCliques = sheetPage.getRange(ultimaLinha+1,4).getValue();
    var newCliques = nCliques + 1;
    sheetPage.getRange(ultimaLinha+1,4).setValue(newCliques);
  } 
  return newCliques
}
