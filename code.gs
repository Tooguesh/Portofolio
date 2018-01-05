// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Cryto Menu')
      .addItem('Update Data','UpdateData')
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
    //UpdateHistory(coinInfo);
    UpdatePortfolio (coinInfo);
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
   var coinInfo ;
   var lastRow = sheet.getLastRow()-1;
  
   var coinsIds = sheet.getRange(2,2,lastRow).getValues();
   //Logger.log(coinsIds);
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

function UpdateHistory ()
{
   var WorkBook = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = SpreadsheetApp.setActiveSheet(WorkBook.getSheetByName("history"));
   var dataRange = sheet.getDataRange().getLastRow();
   var Now = Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd' 'HH:mm")
   
      var coinsIds = MyCoinsAre();
  var coin;
  
  for(var index in coinsIds ){
    coin = GetcoinmarketcapInfo(coinsIds[index]);
    sheet.appendRow([Now,coin.symbol,
                    coin.price_btc,
                    coin.usd,
                    coin.volume_usd_24h,
                    coin.eur,
                    coin.volume_eur_24h,
                    coin.change_1h,
                    coin.change_1d,
                    coin.change_7d,
                    coin.rank])
  }
   Logger.log(Now);
  
}

function UpdatePortfolio (coin)
{
   var EUR_COLUMN=3;
   var BTC_COLUMN = 5;
   var WorkBook = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = SpreadsheetApp.setActiveSheet(WorkBook.getSheetByName("Portfolio"));
   var dataRange = sheet.getDataRange();
   var values = dataRange.getValues();
   
   //remove the header from the 
   values.shift();
   Logger.log(values[0][0]);
   Logger.log(coin.symbol.trim().toLowerCase());
   for (var index in values)
   {
     if (values[index][0] == coin.symbol.trim().toUpperCase())
     {
       var insertRow = Number(index)+2;
       Logger.log("    Index : "+ insertRow);
       sheet.getRange(insertRow,EUR_COLUMN).setValue(Number(coin.eur));
       sheet.getRange(insertRow,BTC_COLUMN).setValue(Number(coin.price_btc));
       //sheet.setActiveRange()
     }
     else 
     {
       Logger.log("No coin referenced as : "+ values[index][0] + " extract from the market datas");
     }
   }

 
}

