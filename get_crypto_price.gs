// custom menu function
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Refresh crypto price','getCryptoPrice')
      .addToUi();
}

function getCryptoPrice() {
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("API CoinMarketCap");
  var url='https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest?limit=2500'
  var requestOptions = {
  method: 'GET',
  uri: 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest',
  qs: {
    start: 1,
    limit: 1000,
    convert: 'USD'
  },
  headers: {
    'X-CMC_PRO_API_KEY': '00000000-7777-4444-5555-111111111111' //insert API key here
  },
    json: true,
    gzip: true
  };

  var httpRequest= UrlFetchApp.fetch(url, requestOptions);
  var json= httpRequest.getContentText();
  var dataSet = JSON.parse(json);
  var rows = [];

  rows.push([
      "name",
      "symbol",
      "cmc_rank",
      "circulating_supply",
      "market_cap",
      "volume_24h",
      "price",
      "percent_change_1h",
      "percent_change_24h",
      "percent_change_7d",
      "percent_change_30d",
      "percent_change_60d",
      "percent_change_90d",
      "last_updated"
  ]);

  for (i = 0; i < dataSet["data"].length; i++) {
    rows.push([
      dataSet["data"][i]["symbol"],
      dataSet["data"][i]["name"],
      dataSet["data"][i]["cmc_rank"],
      dataSet["data"][i]["circulating_supply"],
      dataSet["data"][i]["quote"]["USD"]["market_cap"],
      dataSet["data"][i]["quote"]["USD"]["volume_24h"],
      dataSet["data"][i]["quote"]["USD"]["price"],
      dataSet["data"][i]["quote"]["USD"]["percent_change_1h"],
      dataSet["data"][i]["quote"]["USD"]["percent_change_24h"],
      dataSet["data"][i]["quote"]["USD"]["percent_change_7d"],
      dataSet["data"][i]["quote"]["USD"]["percent_change_30d"],
      dataSet["data"][i]["quote"]["USD"]["percent_change_60d"],
      dataSet["data"][i]["quote"]["USD"]["percent_change_90d"],
      dataSet["data"][i]["quote"]["USD"]["last_updated"]
    ]);
  }

  dataRange = sheet.getRange(1, 1, rows.length, 14);
  dataRange.setValues(rows);
}
