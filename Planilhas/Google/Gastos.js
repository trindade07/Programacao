function DistribuirParcelas() {

   //var Parcelas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Parcelas');
   var PlanilhaExibir = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1RazL5Mbte7IwV3SKhnBxGBzk5eg3xKSmaDZawhOYWdQ/edit?gid=0#gid=0'); 
   var Parcelas = PlanilhaExibir.getSheetByName('Parcelas');
   var Formulario = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form');
   var Historico  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historico');

   Historico.getRange('A1').activate();
   var PrimeiralinhaH = Historico.getRange('A2').getRow();
   var UltimalinhaH = Historico.getLastRow();
   Logger.log('Primeira linha H = ' + PrimeiralinhaH);
    Logger.log('Ultima linha H =  ' + UltimalinhaH);
    
   //Historico.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
   //var LinhaHistorico = Historico.getCurrentCell().getRow();
   var LinhaHistorico = Historico.getLastRow();

   var Primeiralinha = Formulario.getRange('A2').getRow();
   var Ultimalinha = Formulario.getLastRow();
   var UltimaLinhaParcela = Parcelas.getLastRow();
   var UltimaColunaParcela = Parcelas.getLastColumn();

    Logger.log('Primeira linha = ' + Primeiralinha);
    Logger.log('Ultima linha = ' + Ultimalinha);
    Parcelas.setFrozenRows(0);

    if (Parcelas.getLastRow() > 1 )
        {   Parcelas.deleteRows(1, Parcelas.getLastRow())};
    if (Parcelas.getLastColumn() > 1 )
        {  Parcelas.deleteColumns(1,Parcelas.getLastColumn()) };
  
    LinhaParcela = 1;
    // DataInicial = Formulario.getRange(i,2).getValue();
    // 1o param => ano, 2o param =>  o mes de Jan = 0 ate dez = 11, 3o param => dia 
    // DataInicial = new Date(2024,0,426);
    DiaInicialFatura = 26;
    DataInicialFatura = new Date(new Date().getFullYear(), new Date().getMonth(),DiaInicialFatura);
    DataFinalFatura = new Date(new Date().getFullYear(), new Date().getMonth() + 1,DiaInicialFatura);
    UltimaDataParcela = DataInicialFatura;
    PrimeiraDataParcela = DataInicialFatura;
    

    Parcelas.getRange(1,1).setValue('Data Compra');
    Parcelas.getRange(1,2).setValue('Descrição Compra');
    Parcelas.getRange(1,3).setValue('Responsável');
    Parcelas.getRange(1,4).setValue('Pagas/Total');
    Parcelas.getRange(1,5).setValue('Valor Compra');
    Parcelas.getRange(1,6).setValue('Valor Pago');


    for (var i = Primeiralinha; i <= Ultimalinha; i++)
    {
       Logger.log('Linha = ' + i);
       if (Formulario.getRange(i,1).getValue() == "" || Formulario.getRange(i,8).getValue() == "S")
       {
           Logger.log('linha = ' + i + ' sem valor ou encerrada = ' + Formulario.getRange(i,8).getValue());
       }
       else
       {

           

           DataCompra = Formulario.getRange(i,2).getValue();
           MesDataCompra = new Date(DataCompra).getMonth();
           QtdParcelas = Formulario.getRange(i,4).getValue();
           DiaParcelaFinal = new Date(DataCompra).getDate();
           // Mes proxima parcela = mes da compra sempre começa no mes seguinte. 
           MesProximaParcela = MesDataCompra + 1; 
           // Mes da ultima parcela = soma a qtd parcelas
           MesParcelaFinal = MesDataCompra + QtdParcelas;
           AnoParcelaFinal = new Date(DataCompra).getFullYear();

           DataCompraFinal = new Date(AnoParcelaFinal, MesParcelaFinal, DiaParcelaFinal);
           DataProximaParcela = new Date(AnoParcelaFinal, MesProximaParcela, DiaParcelaFinal);
      
           // o mes de Jan = 0 e dez = 11
            Logger.log(' Data Incial Fatura = ' + Utilities.formatDate(new Date(DataInicialFatura),Session.getScriptTimeZone(), 'dd/MM/yyyy'));
            
            Logger.log(' Data Final Fatura = ' + Utilities.formatDate(new Date(DataFinalFatura),Session.getScriptTimeZone(), 'dd/MM/yyyy'));
            Logger.log(' Data Compra Inicial = ' + Utilities.formatDate(new Date(DataCompra),Session.getScriptTimeZone(), 'dd/MM/yyyy'));
            Logger.log(' Data Compra Final = ' + Utilities.formatDate(new Date(DataCompraFinal),Session.getScriptTimeZone(), 'dd/MM/yyyy'));
            Logger.log(' Data Prox Parcela 1 = ' + Utilities.formatDate(new Date(DataProximaParcela),Session.getScriptTimeZone(), 'dd/MM/yyyy'));
            
           
           if (DataCompraFinal < DataInicialFatura)
           {

              Historico.getRange(LinhaHistorico,1).activate();
              Historico.getActiveCell().offset(1,7).activate();
              LinhaHistorico = LinhaHistorico + 1;
              
              Historico.getRange(LinhaHistorico, 1).setValue(Formulario.getRange(i,1).getValue());
              Historico.getRange(LinhaHistorico, 2).setValue(Formulario.getRange(i,2).getValue());
              Historico.getRange(LinhaHistorico, 3).setValue(Formulario.getRange(i,3).getValue());
              Historico.getRange(LinhaHistorico, 4).setValue(Formulario.getRange(i,4).getValue());
              Historico.getRange(LinhaHistorico, 5).setValue(Formulario.getRange(i,5).getValue());
              Historico.getRange(LinhaHistorico, 6).setValue(Formulario.getRange(i,6).getValue());
              Historico.getRange(LinhaHistorico, 7).setValue(Formulario.getRange(i,7).getValue());

              Formulario.getRange(i,8).setValue('S');
             
              Logger.log(' Compra parcelada terminou. Data ultima parcela antes da data atual');
              
            
           }
           else
           {
             
            if (UltimaDataParcela < DataCompraFinal)
             {
                Logger.log('entrou ultima data parcela menor que Data Compra final');
                UltimaDataParcela = DataCompraFinal;
             }

              Parcelas.getRange(LinhaParcela,1).activate();
              Parcelas.getActiveCell().offset(1,7).activate();
              
              LinhaParcela = LinhaParcela + 1;
              Logger.log('Linha parcela = ' + LinhaParcela);
          
              Logger.log('Data compra = ' + Formulario.getRange(i,2).getValue());
              Parcelas.getRange(LinhaParcela,1).setValue(Formulario.getRange(i,2).getValue());

              Logger.log('Nome compra = ' + Formulario.getRange(i,6).getValue());
              Parcelas.getRange(LinhaParcela,2).setValue(Formulario.getRange(i,6).getValue());

              Logger.log('Resp = ' + Formulario.getRange(i,7).getValue());
              Parcelas.getRange(LinhaParcela,3).setValue(Formulario.getRange(i,7).getValue());

              Logger.log('Qtd Parcelas = ' + Formulario.getRange(i,4).getValue());             
              ValorParcela = Formulario.getRange(i,3).getValue() / Formulario.getRange(i,4).getValue();
              Logger.log('Valor Parcela = ' + ValorParcela );

              Logger.log('Vl Compra = ' + Formulario.getRange(i, 3).getValue());
              Parcelas.getRange(LinhaParcela,5).setValue(Formulario.getRange(i,3).getValue());
            
              // exibir valor das parcelas a vencer
                QtdParcelaPaga = 1;
                Nrocoluna = 6;
                
                for (var j = 1; j <= Formulario.getRange(i,4).getValue(); j++)
                {
                    
                    Logger.log('Data Parcela (' + j + ') = ' + Utilities.formatDate(new Date(DataProximaParcela),Session.getScriptTimeZone(), 'dd/MM/yyyy'));

                     // não exibir parcela vencida
                      if (DataProximaParcela < DataInicialFatura)
                      {
                          QtdParcelaPaga = j;
                          Parcelas.getRange(LinhaParcela,6).setValue(QtdParcelaPaga * ValorParcela);
                                                  
                      }
                      
                      else
                      {
                         //if (new Date(DataProximaParcela).getTime() == new Date(DataAtual).getTime())
                         if (DataProximaParcela > DataInicialFatura && DataProximaParcela < DataFinalFatura)
                         {
                            
                            QtdParcelaPaga = j;
                            Logger.log('data dentro do mes da Fatura = ' + QtdParcelaPaga);
                         }
                         else
                         {
                            Logger.log('data Proximo mes da fatura = ' + DataProximaParcela );
                         }
                           
                        Nrocoluna = Nrocoluna + 1;  
                        Parcelas.getRange(LinhaParcela,Nrocoluna).setValue(ValorParcela);
                        Logger.log('Valor = ' + ValorParcela + ' Nrocoluna = ' + Nrocoluna )
                                                   
                        /*MesInicial = MesInicial % 12;
                        if ((MesInicial % 12 ) == 0 )
                        {
                            MesInicial = 12;
                        }
                        Parcelas.getRange(1,j + 2).setValue(MesInicial + '/' + AnoInicial);
                        
                        MesInicial = MesInicial + 1;
                        Logger.log('Proximo mes = ' + (MesInicial % 12 ));
                        if ((MesInicial % 12 ) == 1 )
                        {
                            AnoInicial = AnoInicial + 1;
                        }
                        */
                      }
                      MesColuna = new Date(DataProximaParcela).getMonth();
                      MesColuna = MesColuna + 1;
                      AnoColuna = new Date(DataProximaParcela).getFullYear();
                      DiaColuna = new Date(DataProximaParcela).getDate();

                      DataProximaParcela = new Date(AnoColuna, MesColuna, DiaColuna);
                      
                     
                }
                     Parcelas.getRange(LinhaParcela,4).setValue(QtdParcelaPaga + ' / ' + Formulario.getRange(i,4).getValue());
                    
        
                                
           }

       }

    }

     Logger.log(' Primeira Data todas compras = ' + Utilities.formatDate(new Date(PrimeiraDataParcela),Session.getScriptTimeZone(), 'dd/MM/yyyy'));
      Logger.log(' Ultima Data todas compras = ' + Utilities.formatDate(new Date(UltimaDataParcela),Session.getScriptTimeZone(), 'dd/MM/yyyy'));

      Parcelas.getRange(LinhaParcela + 3 ,1).activate();
      Parcelas.getActiveCell().offset(1,7).activate();
              
      // coluna do Valor Total Parcela
      Nrocoluna = 5;
      CelulaPrimeira = Parcelas.getRange(2,Nrocoluna).getA1Notation();
      CelulaUltima = Parcelas.getRange(LinhaParcela + 1, Nrocoluna).getA1Notation();    
      Parcelas.getRange(LinhaParcela + 3, Nrocoluna).setFormula('=SUBTOTAL(109;' + CelulaPrimeira + ':' + CelulaUltima + ')');
      
      // coluna do Valor Pago
      Nrocoluna = 6;
      CelulaPrimeira = Parcelas.getRange(2,Nrocoluna).getA1Notation();
      CelulaUltima = Parcelas.getRange(LinhaParcela + 1, Nrocoluna).getA1Notation();    
      Parcelas.getRange(LinhaParcela + 2, Nrocoluna).setFormula('=SUBTOTAL(109;' + CelulaPrimeira + ':' + CelulaUltima + ')');
    
      CelulaSubTotal = Parcelas.getRange(LinhaParcela + 2, Nrocoluna).getA1Notation();
      CelulaValorTotal = Parcelas.getRange(LinhaParcela + 3, Nrocoluna - 1).getA1Notation();    
      Parcelas.getRange(LinhaParcela + 3, Nrocoluna).setFormula('= ' + CelulaValorTotal + ' - ' + CelulaSubTotal);

      // inicia na coluna das PARCELAS
      Nrocoluna = 7;
      while (PrimeiraDataParcela <= UltimaDataParcela)
      {
            Logger.log('Data na planilha = ' + Utilities.formatDate(new Date(PrimeiraDataParcela),Session.getScriptTimeZone(), 'MM/yyyy'));
            Parcelas.getRange(1,Nrocoluna).setValue(Utilities.formatDate(new Date(PrimeiraDataParcela),Session.getScriptTimeZone(), 'MM/yyyy'));
            MesColuna = new Date(PrimeiraDataParcela).getMonth() + 1;
            AnoColuna = new Date(PrimeiraDataParcela).getFullYear();
            DiaColuna = new Date(PrimeiraDataParcela).getDate();
            PrimeiraDataParcela = new Date(AnoColuna, MesColuna, DiaColuna);
            CelulaPrimeira = Parcelas.getRange(2,Nrocoluna).getA1Notation();
            Logger.log('Celula Primeira = ' + CelulaPrimeira);

            CelulaUltima = Parcelas.getRange(LinhaParcela + 1, Nrocoluna).getA1Notation();
            Logger.log('Celula ultima = ' + CelulaUltima);
            Parcelas.getRange(LinhaParcela + 2, Nrocoluna).setFormula('=SUBTOTAL(109;' + CelulaPrimeira + ':' + CelulaUltima + ')');
            CelulaSubTotal = Parcelas.getRange(LinhaParcela + 2, Nrocoluna).getA1Notation();
            CelulaValorTotal = Parcelas.getRange(LinhaParcela + 3, Nrocoluna - 1).getA1Notation();    

            Parcelas.getRange(LinhaParcela + 3, Nrocoluna).setFormula('= ' + CelulaValorTotal + ' - ' + CelulaSubTotal);

            Nrocoluna = Nrocoluna + 1;

            

      }
      Parcelas.getRange(LinhaParcela + 4 ,1).activate();
      Parcelas.getActiveCell().offset(1,7).activate();

      //Criar um filtro
      CelulaPrimeira = Parcelas.getRange(1,1).getA1Notation();
      CelulaUltima = Parcelas.getRange(LinhaParcela, Nrocoluna).getA1Notation();

      Parcelas.getRange('' + CelulaPrimeira + ':' + CelulaUltima + '').createFilter();
     
      // Ordenar pela coluna 1 => Data compra  
      Parcelas.getRange('' + CelulaPrimeira + ':' + CelulaUltima + '').activate();
      Parcelas.getActiveRange().offset(1, 0, Parcelas.getActiveRange().getNumRows() - 1).sort({column: 1, ascending: true});

      //Formatar celulas com valor com duas casas decimais. Celula primeira => Linha 2 coluna 5
      
      CelulaPrimeira = Parcelas.getRange(2,5).getA1Notation();
      Parcelas.getRange('' + CelulaPrimeira + ':' + CelulaUltima + '').activate();
      Parcelas.getActiveRange().setNumberFormat('0.00');

      // Congelas a 1a linha
      Parcelas.setFrozenRows(1);

}

function LimpezaParcelas() {

   var Formulario = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form');
   
   var Linha = Formulario.getRange('A2').getRow();
   var Ultimalinha = Formulario.getLastRow();

   while (Linha <= Ultimalinha)
   {
        Logger.log('Linha = ' + Linha);
       if (Formulario.getRange(Linha,1).getValue() == "" || Formulario.getRange(Linha,8).getValue() == "S")
       {
           Logger.log('linha = ' + Linha + ' sem valor ou encerrada = ' + Formulario.getRange(Linha,8).getValue());
           Formulario.deleteRow(Linha);
           Ultimalinha = Ultimalinha -1;
       }
       else
       {
        Linha = Linha + 1;
       }
 
 
   }      
  
}

function TesteCelula ()
{
  var Parcelas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Parcelas');
  Parcelas.getRange(2,5).activate();
  CelulaPrimeira = Parcelas.getCurrentCell().getA1Notation();
  Logger.log('Celula Primeira = ' + CelulaPrimeira);
  Parcelas.getRange(10 + 2, 5).activate();
  CelulaUltima = Parcelas.getCurrentCell().getA1Notation();
  //Celulatextoultima = Parcelas.getDataRange()
  Logger.log('Celula ultima = ' + CelulaUltima);
}


function TesteSomaData ()
{
  // definir como 01 de dezembro de 2024
    DataAtual = new Date(2024, 11, 1);
    // o mes de Jan = 0 e dez = 11
    Logger.log(' Data Atual = ' + Utilities.formatDate(new Date(DataAtual),Session.getScriptTimeZone(), 'dd/MM/yyyy'));

// diminuir 3 dias do dia 01 de dezembro de 2024
    DataSoma = new Date(2024, 11, (1 - 4));
    Logger.log(' Data Somada = ' + Utilities.formatDate(new Date(DataSoma),Session.getScriptTimeZone(), 'dd/MM/yyyy'));

// definir o dia 26 do mes passado do dia 01 de dezembro de 2024
    DataSoma = new Date(2024, (11 - 1), 26);
    Logger.log(' Data Somada = ' + Utilities.formatDate(new Date(DataSoma),Session.getScriptTimeZone(), 'dd/MM/yyyy'));

}
