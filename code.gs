// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Cryto Menu')
      .addItem('Update Data','UpdateData').addItem("Update History", "UpdateHistory")
      .addToUi();
  
}

/*
Fucntion : UpdateData 
Parameter : 
Return Value : Object  
Description : 
*/
function UpdateData ()
{
  // Initialisation of 
  var coinsIds = MyCoinsAre();
  
  for(var index in coinsIds ){
    
    coinInfo = GetcoinmarketcapInfo(coinsIds[index]);
   
    UpdatePortfolio (coinInfo);
    UpdateBook (coinInfo);
    //  Logger.log("CoinsInfo ="+ coinInfo.symbol);  
  }
}


/*
Fucntion : MyCoinsAre 
Parameter : #N/A
Return Value : Object  
Description : Read the Master Data Spreadsheet to extract the coins in your possession. 
*/
function MyCoinsAre()
{
   var WorkBook = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = SpreadsheetApp.setActiveSheet(WorkBook.getSheetByName("MasterData"));
   
   var lastRow = sheet.getLastRow()-1;
   
   // Get all coins IDs for the marketcoincap API
   var coinsIds = sheet.getRange(2,2,lastRow).getValues();
   
   return coinsIds;
}

/*
Fucntion : GetCoinData 
Parameter : 
Return Value : Object  
Description : 
*/

function GetcoinmarketcapInfo(coin) {
  
  //var URI = "https://api.coinmarketcap.com/v1/ticker/?limit=10"; 
  var URI = "https://api.coinmarketcap.com/v1/ticker/"+ coin +"/?convert=EUR";
  // Call the Numbers API for random math fact
  var response = UrlFetchApp.fetch(URI);
  //Logger.log(response.getContentText());
  
  var JsonResponse = JSON.parse(response.getContentText());
  //var data = (JSON.parse(JsonResponse));
  
  
//  Logger.log(JsonResponse[0]);
  var CoinData = {symbol:String(JsonResponse[0]['symbol']),
              price_btc:parseFloat(JsonResponse[0]['price_btc']),
              usd: parseFloat(JsonResponse[0]['price_usd']),
              volume_usd_24h : parseFloat(JsonResponse[0]['24h_volume_usd']),
              eur: parseFloat(JsonResponse[0]['price_eur']),
              volume_eur_24h : parseFloat(JsonResponse[0]['24h_volume_eur']),
              change_1h: Number(JsonResponse[0]['percent_change_1h']),
              change_1d : Number(JsonResponse[0]['percent_change_24h']),
              change_7d :Number(JsonResponse[0]['percent_change_7d']),
              rank : parseInt(JsonResponse[0]['rank'])
             };
  return CoinData;

}


function UpdatePortfolio (coin)
{
   var EUR_COLUMN=3;
   var BTC_COLUMN = 5;
   
   var WorkBook = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = SpreadsheetApp.setActiveSheet(WorkBook.getSheetByName("Portfolio"));
   var dataRange = sheet.getDataRange();
   var values = dataRange.getValues();
  
   var insertRow =-1;
   
   //remove the header from the 
   values.shift();
   Logger.log(values[0][0]);
   Logger.log(coin.symbol.trim().toLowerCase());
   for (var index in values)
   {
     if (values[index][0] == coin.symbol.trim().toUpperCase())
     {
       insertRow = Number(index)+2;
       Logger.log("    Index : "+ insertRow);
       
       sheet.getRange(insertRow,EUR_COLUMN).setFontColor(ColorEvolution(sheet.getRange(insertRow,EUR_COLUMN).getValue(),coin.eur));
       sheet.getRange(insertRow,EUR_COLUMN).setValue(Number(coin.eur));
       
       sheet.getRange(insertRow,BTC_COLUMN).setFontColor(ColorEvolution(sheet.getRange(insertRow,BTC_COLUMN).getValue(),coin.price_btc));
       sheet.getRange(insertRow,BTC_COLUMN).setValue(Number(coin.price_btc));
       //sheet.setActiveRange()
     }
     else 
     {
       Logger.log("No coin referenced as : "+ values[index][0] + " extract from the market datas");
     }
   }


 
}

function UpdateHistory()
{
  /*
  * Recuperation de la liste des tokens suivi 
  */
  var coinsIds = MyCoinsAre();
  var coin;
  
  var WorkBook = SpreadsheetApp.getActiveSpreadsheet();
  
  //var sheet = SpreadsheetApp.setActiveSheet(WorkBook.getSheetByName("history"));
  //var dataRange = sheet.getDataRange().getLastRow();
  var Now = Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd' 'HH:mm")
  
    for(var index in coinsIds ){
    coin = GetcoinmarketcapInfo(coinsIds[index]);
    var sheetName = "history_"+coin.symbol ;
    
    
    try {
       sheet = SpreadsheetApp.setActiveSheet(WorkBook.getSheetByName(sheetName));
    }
    catch(err) {
       sheet = createHistoryCoinSheet (sheetName);
    } 
     
    var dataRange = sheet.getDataRange().getLastRow();
    var Now = Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd' 'HH:mm")
    sheet.appendRow([Now,coin.symbol,
                     coin.price_btc,
                     coin.usd,
                     coin.volume_usd_24h,
                     coin.eur,
                     coin.volume_eur_24h,
                     coin.change_1h/100,
                     coin.change_1d/100,
                     coin.change_7d/100,
                     coin.rank])
  }
   Logger.log(Now);
  
}


function createHistoryCoinSheet (sheetName)
{
   var WorkBook = SpreadsheetApp.getActiveSpreadsheet();  
   WorkBook.insertSheet(sheetName);
  
   var sheet = SpreadsheetApp.setActiveSheet(WorkBook.getSheetByName(sheetName));
   sheet.appendRow(['DateTime',
                    'symbol',
                    'price_btc',
                    'usd',
                    'volume_usd_24h',
                    'eur',
                    'volume_eur_24h',
                    'change_1h',
                    'change_1d',
                    'change_7d',
                    'rank']);
  return sheet;
}

function UpdateBook(coin)
{
   var COURS_COLUMN = 11;
   var EVOL_COLUMN = 12;
   var RECO_COLUMN = 13;
   
   var WorkBook = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = SpreadsheetApp.setActiveSheet(WorkBook.getSheetByName("Book"));
   var dataRange = sheet.getDataRange();
   var values = dataRange.getValues();
  
   var insertRow =-1;
  
   var cours_Achat = -1;
   var Evolution = 0;
   var recommandation = 0;
  
   //remove the header from the 
   values.shift();
   Logger.log(values[0][0]);
   Logger.log(coin.symbol.trim().toLowerCase());
   for (var index in values)
   {
     if (values[index][2] == coin.symbol.trim().toUpperCase())
     {
       insertRow = Number(index)+2;
       Logger.log("    Index : "+ insertRow);
       cours_Achat = sheet.getRange(insertRow,6).getValue();
       sheet.getRange(insertRow,COURS_COLUMN).setFontColor(ColorEvolution(cours_Achat,coin.eur));
       sheet.getRange(insertRow,COURS_COLUMN).setValue(Number(coin.eur));
       
       Evolution = (coin.eur - cours_Achat )/ cours_Achat;
       sheet.getRange(insertRow,EVOL_COLUMN).setFontColor(ColorEvolution(0,Evolution));
       sheet.getRange(insertRow,EVOL_COLUMN).setValue(Number(Evolution));
  
     }
     else 
     {
       Logger.log("No coin referenced as : "+ values[index][0] + " extract from the market datas");
     }
   }

}


/*
Fucntion : ColorEvolution 
Parameter :  
Return Value : Color <String> : Font color to be set to the cell 
Description : Compare the two parameters Values to get the Color reprensenting positive or negative price Evolution 
*/

function ColorEvolution (OldValue, NewValue)
{
  var RED = "#c53929";
  var GREEN = "#0b8043" ;
  var BLACK = "#000000";
  if (OldValue < NewValue )
  {
    return GREEN ;
  }
  else if (OldValue > NewValue)
  {
    return RED;
  }
  
  return BLACK;
}


/*
               ------------------------------
               ANCIENNE FONCTION D'HISTORIQUE
               ------------------------------
               Abandonner car ne creait que une seule feuille la rendqnt illisible. 


function UpdateHistory ()
{
  
  // Recuperqtion de la liste des tokens suivi 
  
  
  var coinsIds = MyCoinsAre();
  var coin;
  
  
  var WorkBook = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.setActiveSheet(WorkBook.getSheetByName("history"));
  var dataRange = sheet.getDataRange().getLastRow();
  var Now = Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd' 'HH:mm")
  
  
  for(var index in coinsIds ){
    coin = GetcoinmarketcapInfo(coinsIds[index]);
    sheet.appendRow([Now,coin.symbol,
                    coin.price_btc,
                    coin.usd,
                    coin.volume_usd_24h,
                    coin.eur,
                    coin.volume_eur_24h,
                    coin.change_1h/100,
                    coin.change_1d/100,
                    coin.change_7d/100,
                    coin.rank])
  }
   Logger.log(Now);
  
}

*/

