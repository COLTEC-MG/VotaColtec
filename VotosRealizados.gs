function votosRealizados(emailEleitor,linkPlanilha) {
  
  // Recupera a senha das propriedades do script  
  var cache = CacheService.getScriptCache();
  var SENHA = cache.get('Word');
  var cipher = new cCryptoGS.Cipher(SENHA, 'aes');
  
  // Abre a planilha e a página com os dados de votação. Procura a linha correspondente ao e-mail do usuário.
  var sheet = SpreadsheetApp.openByUrl(linkPlanilha);
  var sheetPage = SpreadsheetApp.openByUrl(linkPlanilha).getSheetByName('DadosVotação');
  var procurar = sheetPage.createTextFinder(emailEleitor);
  var linha = procurar.findNext();
  var ultimaLinha = sheetPage.getLastRow();
  
  // Caso encontre a linha do usuário, verifica quantas opções de voto já foram realizadas.
  if (linha != null) {
    var linhaUser = linha.getRow();
    var opcoes = sheetPage.getRange(linhaUser,3).getValue();
    if (opcoes == 0) { // Caso não haja nada, retorna 0
      return 0
    }
    else {
      var linhaUser = linha.getRow();
      var opcoes = sheetPage.getRange(linhaUser,3).getValue();
      opcoes = cipher.decrypt(opcoes);
      var opcoesTotais = opcoes.split('$');
      var tam = opcoesTotais.length;
      return tam // Caso hajam votos, retorna o tamanho do vetor de votos realizados.
    }
  }
  else {
    return 0
  }
}

