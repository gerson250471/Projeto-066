function rel90dias(){
    let planilha = SpreadsheetApp.getActive();
    let aba1= planilha.getSheetByName("BD_TAREFAS");
    let aba2 = planilha.getSheetByName("RelDet");
  
    let rng1 = aba1.getRange("A2:I").getValues();
    let rng2 = aba2.getRange("A1:E").getValues();
  
    let dt1 = new Date();
    dt1.setDate(dt1.getDate()-90);
    let lbd = aba1.getLastRow()-1;
  
    aba2.getRange("A2:E").clearContent();
    let lult = aba2.getLastRow()+1;
  
    for (var a=0; a < lbd;a ++){
      if(rng1[a][3]>dt1){
        aba2.getRange(lult,1).setValue(rng1[a][0]);     //  Tarefa
        aba2.getRange(lult,2).setValue(rng1[a][1]);     //  Responsável
        aba2.getRange(lult,3).setValue(rng1[a][2]);     //  Horario
        aba2.getRange(lult,4).setValue(rng1[a][3]);     //  Data
        aba2.getRange(lult,5).setValue(rng1[a][4]);     //  Status
      }
      lult = aba2.getLastRow()+1;
    }
    Browser.msgBox("Relatório Gerado com Sucesso");
  }