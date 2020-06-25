function buscaVoto(emailEleitor,opcao, linkPlanilha) {
  
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
  
  // Caso encontre o usuário, verifica se o mesmo já realizou voto naquela opção.
  if (linha != null) {
    var linhaUser = linha.getRow(); // Busca o número da linha do usuário.
    var opcoes = sheetPage.getRange(linhaUser,3).getValue(); // Pega a cédula dos votos já realizados do usuário.
    opcoes = cipher.decrypt(opcoes); // Descriptografa os dados.
    opcoes = opcoes.toString(); // Transforma em string.
    var opcoesTotais = opcoes.split('$'); // Separa a string em um vetor.
    var tam = opcoesTotais.length; // Pega o tamanho do vetor.
    for (var i = 0; i < tam; i++) {
      if (opcoesTotais[i] == opcao) { // Verifica se o usuário já votou naquela opção a partir das opções gravadas dele.
        return true // Caso positivo, retorna true
      }
    }
  }
  // Caso não se encontre a linha do usuário, retorna falso.
  else {
    return false
  } 
  return false
}
