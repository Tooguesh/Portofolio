function GetcoinbaseInfo() {
  var coin = "EUR";
  
  var URI = "https://api.coinbase.com/v2/exchange-rates?currency="+coin;
  //var URI = "https://api.coinbase.com/v2/exchange-rates";
  
  Logger.log(URI);
  // Call the Numbers API for random math fact
  var response = UrlFetchApp.fetch(URI);
  Logger.log(response.getContentText());
  
  var JsonResponse = JSON.parse(response.getContentText());
  Logger.log("End Script");
}
