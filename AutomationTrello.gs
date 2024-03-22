/* Autor do código: Matheus Bueno Faria
Desafio técnico - Seazone*/

function lerDadosDaPlanilha() {

    var planilhaAtiva = SpreadsheetApp.openById('1fF-NZG4_U2-cHeTnDPxwoC6JUVXjeqEbRSN6AXGRNRM')

    var abaObjetivo = planilhaAtiva.getSheetByName('Respostas ao formulário 3')
    
    var ultimaLinha = abaObjetivo.getLastRow();

    //notação A1
    var intervaloA1 = abaObjetivo.getRange('B2:E' + ultimaLinha)
    //forma de linhas e colunas
    //var intervaloA2 = abaObjetivo.getRange(2,1,2,4)
    Logger.log(intervaloA1)

    //pega os valores 
    var dados =  intervaloA1.getValues()
    Logger.log(dados)
    return dados;
}

function processarDados() {

    var planilhaAtiva = SpreadsheetApp.openById('1fF-NZG4_U2-cHeTnDPxwoC6JUVXjeqEbRSN6AXGRNRM')

    var abaObjetivo = planilhaAtiva.getSheetByName('Respostas ao formulário 3')
    var dados = lerDadosDaPlanilha();
        for (var i = 0; i < dados.length; i++) {
            var enviadoParaTrello = dados[i][3]; // Quarta coluna (coluna 'E') para marcar se os dados foram enviados para o Trello
            if (enviadoParaTrello !== 'X') {
                var nomeCartao = dados[i][0]; // 1ª coluna
                var descricaoCartao = dados[i][1]; // 2ª coluna
                var listaSelecionada = dados[i][2]; // 3ª coluna

                var chaveTrello = '8720ebb8c1757d415ef4dc782d597783';
                var tokenTrello = 'ATTA191c55de21d90f385ac573e9aa48417988203edb830fe4d34bce726e9c29827819EF07D5';
                var idBoardTrello = 'inDnB5KN';
                var url = 'https://api.trello.com/1/cards?key=' + chaveTrello + '&token=' + tokenTrello;
                
                var payload = {
                  method: 'POST',
                  muteHttpExceptions: true,
                  payload: {
                    name: nomeCartao,
                    desc: descricaoCartao,
                    idList: obterIdListaTrello(listaSelecionada, chaveTrello, tokenTrello, idBoardTrello)
                  }
                };
                
                var response = UrlFetchApp.fetch(url, payload);
                Logger.log(response.getContentText());

              }
            Logger.log(nomeCartao);
    }
}

function obterIdListaTrello(nomeLista, chave, token, idBoard) {
  
    var url = 'https://api.trello.com/1/boards/' + idBoard + '/lists?key=' + chave + '&token=' + token;
    var response = UrlFetchApp.fetch(url);
    var listas = JSON.parse(response.getContentText());
    
    for (var i = 0; i < listas.length; i++) {
      if (listas[i].name === nomeLista) {
        return listas[i].id;
      }
    }
    return null;
}