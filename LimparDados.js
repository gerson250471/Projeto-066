function limparDados(){
    let planilha = SpreadsheetApp.getActive();
    let aba1= planilha.getSheetByName("01");
    let aba2 = planilha.getSheetByName("02");
    
    aba1.getRange("B4:I20").clearContent();
    aba2.getRange("B4:I20").clearContent();
    aba1.getRange("B38:I55").clearContent();
    aba2.getRange("B38:I55").clearContent();
  }