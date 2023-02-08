function arquivar() {
  // chamando as variáveis
  let planilha = SpreadsheetApp.getActive();
  let aba1= planilha.getSheetByName("01");
  let aba2 = planilha.getSheetByName("02");
  let aba3 = planilha.getSheetByName("BD_TAREFAS");
  // colocando a informação da planilha
  let rng1 = aba1.getRange("B4:I20").getValues();
  let rng2 = aba2.getRange("B4:I20").getValues();
  let arqLf = aba3.getLastRow()+1;
  // colocando a informação das pessoas responsáveis
  let r1 = aba1.getRange("B2").getValue();
  let r2 = aba2.getRange("B2").getValue();

  for (var  a = 0; a < 17; a ++){
    if (rng1[a][0] != "") {
      aba3.getRange(arqLf,1).setValue(rng1[a][0]);
      aba3.getRange(arqLf,2).setValue(r1);
      aba3.getRange(arqLf,3).setValue(rng1[a][5]);
      aba3.getRange(arqLf,4).setValue(rng1[a][6]);
      aba3.getRange(arqLf,5).setValue(rng1[a][7]);
      if (rng1[a][7]!=""){
        aba3.getRange(arqLf,9).setValue(1);
      }else{
        aba3.getRange(arqLf,9).setValue(0);
      }

      var dt =  new Date(rng1[a][6]);
      var dt1=dt.getMonth()+1;
      var dt2=dt.getFullYear();
    
      aba3.getRange(arqLf,6).setFormula('=WEEKNUM(R[0]C[-2])');
      aba3.getRange(arqLf,7).setValue(dt1);
      aba3.getRange(arqLf,8).setValue(dt2);
      arqLf = aba3.getLastRow()+1;
    }
}