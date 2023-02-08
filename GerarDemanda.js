function gerarDemanda(){
  let planilha = SpreadsheetApp.getActive();
  let aba1 = planilha.getSheetByName("LISTA DE TAREFAS");
  let aba2= planilha.getSheetByName("01");
  let aba3 = planilha.getSheetByName("02");
  
  // Obtendo as informações da planilha
  let rng1 = aba1.getRange("A3:F").getValues();
  let lbd = aba1.getLastRow()+1;
  let laba2=4;
  let laba3=4;
  let resp1=aba1.getRange("R3").getValue();
  let resp2=aba1.getRange("R4").getValue();

  let dt = new Date();
  let dtnome 

  if (dt.getDay()==0){
    dtnome="Domingo"
  }

  if (dt.getDay()==1){
    dtnome="Segunda-feira"
  }

  if (dt.getDay()==2){
    dtnome="Terça-feira"
  }

  if (dt.getDay()==3){
    dtnome="Quarta-feira"
  }

  if (dt.getDay()==4){
    dtnome="Quinta-feira"
  }

  if (dt.getDay()==5){
    dtnome="Sexta-feira"
  }

  if (dt.getDay()==6){
    dtnome="Sábado"
  }

  // Obter as Tarefas do dia da Semana do primeiro Profissional
  for (var a=0; a < lbd; a ++){
    if (rng1[a][2]==resp1){
        if (rng1[a][5]=="Semana"){
          if (rng1[a][4]==dtnome){
          aba2.getRange(laba2,2).setValue(rng1[a][0]);
          aba2.getRange(laba2,7).setValue(rng1[a][3]);
          aba2.getRange(laba2,8).setValue(dt);
          laba2=laba2+1;
        }
      }
    }
  }
  
    // Obter as Tarefas do dia do primeiro Profissional

    for (var a=0; a < lbd; a ++){
    if (rng1[a][2]==resp1){
        if (rng1[a][5]=="Mês"){
          if (rng1[a][4]==dtnome){
          aba2.getRange(laba2,2).setValue(rng1[a][0]);
          aba2.getRange(laba2,7).setValue(rng1[a][3]);
          aba2.getRange(laba2,8).setValue(dt);
          laba2=laba2+1;
        }
      }
    }
  }

  laba2=38;
    // Obter as Próximas Tarefas do mês do primeiro profissional
  for (var a=0; a < lbd; a ++){
    if (rng1[a][2]==resp1){
        if (rng1[a][5]=="Mês"){
          if (rng1[a][4]>dt.getDate()){
          aba2.getRange(laba2,2).setValue(rng1[a][0]);
          if (dt.getMonth()==2){
              if (rng1[a][5]>28){
                aba2.getRange(laba2,9).setValue(28);    
              }
          } else {
            aba2.getRange(laba2,9).setValue(rng1[a][4]);
            laba2=laba2+1;
          }
        }
      }
    }
  }

  // Obter as Tarefas do dia da Semana do Segundo Profissional
  for (var a=0; a < lbd; a ++){
    if (rng1[a][2]==resp2){
        if (rng1[a][5]=="Semana"){
          if (rng1[a][4]==dtnome){
          aba3.getRange(laba3,2).setValue(rng1[a][0]);
          aba3.getRange(laba3,7).setValue(rng1[a][3]);
          aba3.getRange(laba3,8).setValue(dt);
          laba3=laba3+1;
        }
      }
    }
  }
  
    // Obter as Tarefas do dia do Segundo Profissional
  for (var a=0; a < lbd; a ++){
    if (rng1[a][2]==resp2){
        if (rng1[a][5]=="Mês"){
          if (rng1[a][4]==dtnome){
          aba3.getRange(laba3,2).setValue(rng1[a][0]);
          aba3.getRange(laba3,7).setValue(rng1[a][3]);
          aba3.getRange(laba3,8).setValue(dt);
          laba3=laba3+1;
        }
      }
    }
}
