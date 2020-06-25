// Linha da planilha: Nome; E-mail; Opcoes Escolhidas; Nº de Cliques

function gravaDados(nomeEleitor,emailEleitor,opcao,linkPlanilha) {

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
  
  // Caso encontre o usuário.
  if (linha != null) {
    var linhaUser = linha.getRow();
    var opcoes = sheetPage.getRange(linhaUser,3).getValue(); // Pega a cédula dos votos já realizados do usuário.
    opcoes = cipher.decrypt(opcoes); // Descriptografa os dados.
    if (opcoes == 0 || opcao == 'nulo' || opcao == 'branco') { // Caso não haja votos ou o usuário deseja votar branco ou nulo
      var opcao_aux = cipher.encrypt(opcao); // Criptografa a opção de voto.
      sheetPage.getRange(linhaUser,3).setValue(opcao_aux); // Grava a opção de voto criptografada.
      sheetPage.getRange(linhaUser,1).setValue(nomeEleitor); // Grava o nome do eleitor.
    }
    // Caso o usuário queira fazer um outro voto se permitido mais opções.
    else {
      var linhaUser = linha.getRow(); // Pega a linha do usuário.
      var opcoes = sheetPage.getRange(linhaUser,3).getValue(); // Pega a cédula dos votos já realizados do usuário.
      opcoes = cipher.decrypt(opcoes); // Descriptografa os dados.
      var opcoesTotais = opcoes.split('$'); // Separa a string dos votos.
      opcoesTotais.push(opcao); // Adiciona o novo voto.
      var opcoesTotais2 = opcoesTotais.join(['$']); // Junta a string de votos novamente.
      var opcoesTotais3 = cipher.encrypt(opcoesTotais2); // Criptografa os votos.
      sheetPage.getRange(linhaUser,3).setValue(opcoesTotais3); // Grava a string de votos.
      sheetPage.getRange(linhaUser,1).setValue(nomeEleitor); // Grava o nome do eleitor.
    }
  }
  // Caso o usuário não tenha uma linha própria.
  else {
    // Grava na última linha sem dados o nome, o e-mail e as opções criptografadas.
    var opcao_aux = cipher.encrypt(opcao);
    sheetPage.getRange(ultimaLinha+1,1).setValue(nomeEleitor);
    sheetPage.getRange(ultimaLinha+1,2).setValue(emailEleitor);
    sheetPage.getRange(ultimaLinha+1,3).setValue(opcao_aux);
  }
  return
}
