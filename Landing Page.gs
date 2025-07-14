function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('MCD Memo Routing')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
