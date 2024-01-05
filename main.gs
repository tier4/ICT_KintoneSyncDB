function updateSheet() {
  const sheetID   = "1RJBRmKYP6Axg7Yt8ZHbFnKaIDbweuuxNct50xcAb-ME";
  const srcSheet  = SpreadsheetApp.openById(sheetID);
  const srcTab    = srcSheet.getSheetByName("AppSync");
  const reqAppList  = srcTab.getRange("B3:C").getValues();

  ICT_HRDB.ClsCommonFunc.updateStatsIndicator(sheetID,"AppSync","B1","C1",true);
  for(var i in reqAppList) {
    const lineItem = reqAppList[i];
    if (lineItem.length == 2) {
      const appID = lineItem[0];
      const page  = lineItem[1];
      if (appID > 0 && page.toLocaleString().length > 4) {
        const targetTab = srcSheet.getSheetByName(page);
        const result    = ICT_HRDB.ClsCybozu.getDataArrayFromApp(appID);
        const dataRange = targetTab.getDataRange();
        targetTab.getRange(2,1,dataRange.getLastRow(),dataRange.getLastColumn()).clear(
          {
            contentsOnly:true,
            skipFilteredRows:false
          }
        );
        if (result.length > 0) {
          targetTab.getRange(1,1,result.length,result[0].length).setValues(result);
        }
      }
    }
  }
  ICT_HRDB.ClsCommonFunc.updateStatsIndicator(sheetID,"AppSync","B1","C1",false);
}

