// KimonoのAPI IDとAPI Key
var API_ID  = "YourApiId";
var API_KEY = "YourApiKey";

// データ取得のためのURL
var FETCH_URL = "https://www.kimonolabs.com/api/" + API_ID + "?apikey=" + API_KEY;

// クローリング開始のためのURL
var START_CRAWL_URL = "https://www.kimonolabs.com/kimonoapis/" + API_ID + "/startcrawl";

// ランキングを記録していくシートの名前
var SHEET_NAME = "価格推移";


function ranking2Sheet(){
  // Kimono APIへのリクエストを発行して結果を加工
  var response = request();
  var ranking  = process(response.crawlResults);


  // 価格情報(価格+送料)
  var prices = ranking.map(function(item) {
    return [item.price + "\n+ 送料:" + item.postage];
  });

  // 在庫状況
  var stocks = ranking.map(function(item) {
    return ["在庫:" + item.stock];
  });

  // 在庫状況はセルの背景色でも表現
  var bgColors = ranking.map(function(item) {
    return [item.stockColor];
  });

  // 店舗情報(店名+商品ページへのリンク)
  var shopInfos = ranking.map(function(item) {
    return ['=HYPERLINK("' + item.shopUrl + '","' + item.shopName + '")'];
  });


  // シートの選択
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    // 対象のシートが見つからなかったら新規作成して名前を設定
    sheet = activeSpreadsheet.insertSheet().setName(SHEET_NAME);
  }


  // B列の左に2列挿入
  sheet.insertColumns(2, 2);

  // 追加した列(B,C列)の背景色をデフォルトの"#fff"にで塗りつぶし
  sheet.getRange("B:C")
    .setBackground("#fff");

  // ヘッダーの設定
  sheet.getRange("B1:C1")
    .merge()                          // セルを結合して
    .setFontWeight("bold")            // 太字にして
    .setHorizontalAlignment("center") // 中央揃えにして
    .setValue(new Date())             // 値に現在時刻を設定
    .setNumberFormat("MM月dd日");     // 表示形式を"MM月dd日"に設定

  // 価格情報の列を設定
  sheet.getRange(2, 2, prices.length, prices[0].length)
    .setBackgrounds(bgColors)         // 背景色を設定して
    .setNotes(stocks)                 // メモを挿入して
    .setValues(prices);               // 値に「価格+送料」を設定

  // 店舗情報の列を設定
  // "HYPERLINK()"関数でセルのテキストにリンクを貼ります
  sheet.getRange(2, 3, shopInfos.length, shopInfos[0].length)
    .setFormulas(shopInfos);          // 値「店舗名」とリンク(hyperlink)の設定
}


// Kimono API側でクローニングを開始させる
function startCrawl() {
  // Kimono API側の仕様としてクローニングの開始はPOSTでリクエストを飛ばす決まりなので
  // API Keyはクエリ文字列ではなくpayload(e.g. POST body)として渡す必要がある
  var payload = {
    apikey: API_KEY
  };

  return UrlFetchApp.fetch(START_CRAWL_URL, {
    method: "post",
    contentType: "application/json", // JSON形式で、これもKimono API側の仕様
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
      // そのままUrlFetchApp.fetch()に渡すとkey/value mapに変換されてしまうので
      // 事前にJSON.stringify()する
  });
}


// Kimono APIからのレスポンスを加工
function process(crawlResults) {
  return crawlResults["価格ランキング"].map(function(item, index) {

    // 在庫状況は2行に渡る場合がある、
    // 2行目があれば括弧で括って1行にまとめる
    var stockRows  = (item["在庫有無"].text || item["在庫有無"]).split(/\r\n|\r|\n/) ;
    var stock      = stockRows[0] + (stockRows[1] != null ? "(" + stockRows[1] + ")" : "");

    // 在庫状況によって背景色を色分けする
    // 在庫状況は "有" or "問合せ" or "～○営業日"
    var stockColor = (function(str) {
      switch (true) {
        case /^有/.test(str):
          return "#b6d7a8"; // 緑
        case /^問合せ/.test(str):
          return "#ea9999"; // 赤
        default:
          return "#f9cb9c"; // オレンジ
      }
    })(stockRows);

    // その他、必要な情報をオブジェクトにまとめて返す
    return {
      rank       : index + 1,
      price      : item["価格"],
      postage    : item["送料"],
      shopName   : item["店名"].text,
      shopUrl    : item["リンク"].href,
      stock      : stock,
      stockColor : stockColor,
    };
  });
};


// Kimono APIへのリクエストを組み立て、発行する
function request() {
  var fetchOptions = {
    method             : "get",
    muteHttpExceptions : true,  // HTTPステータスコードでエラーが返っても例外を投げない
  };

  var response = UrlFetchApp.fetch(FETCH_URL, fetchOptions);

  try {
    var body = JSON.parse(response.getContentText());

    return {
      responseCode      : response.getResponseCode(),
      parseError        : false,
      body              : body,
      bodyText          : response.getContentText(),
      thisVersionRun    : body.thisversionrun,    // 最後にKimono APIを走らせた日時
      thisVersionStatus : body.thisversionstatus, // 最後にKimono APIを走らせた際のステータス
      crawlResults      : body.results,           // クロール結果
    };
  } catch(e) {
    return {
      responseCode      : response.getResponseCode(),
      parseError        : true,
      body              : null,
      bodyText          : response.getContentText(),
      thisVersionRun    : null,
      thisVersionStatus : null,
      crawlResults      : null,
    };
  }
}
