function main() {
  DetermineDebugMode();

  var newRecord = GetNewRecordMeta();

 

  var formSub = FormApp.openById(FormId);
  var formResponses = formSub.getResponses();
  var formLength = formResponses.length;

  //tables are 0 based, getting the lastest repsonse instead of using all of them
  var latestResponse = formResponses[formLength - 1];
  var items = latestResponse.getItemResponses();

  //Looping through the questions, get the title and the answer
  for (var i = 0; i < items.length; i++) {
    var currentResonseItem = items[i];
    var title = currentResonseItem.getItem().getTitle();
    var answer = currentResonseItem.getResponse();

    newRecord[title] = answer;
  };

  // newRecord = {
  //     Name:    "John Doe",
  //     Date:    "2025-08-11",
  //     DayPart: "Dinner",
  //     Nuggets:  10, 
  //     Strips:   10,
  //     Filets:   10,
  //     Spicy:    10
  // }

  print(newRecord);

  var DateSplit = newRecord.Date.split('-')
  var DateSplitMonth = DateSplit[1]
  var DateSplitPosition = DateSplit[2]

  print(`Position: ${Number(DateSplitPosition)}`)


  MonthIndex.forEach(function(Value) {
    let targetMonth = `${Value.Month}`
    if (Value.Month < 10) { targetMonth = `0${Value.Month}` }

    if (DateSplitMonth != targetMonth) { return }

    print(`Month: ${Value.Name}`)
    NameOfYourSheet = Value.Name


    let sheetInstance = GetSheet()
    int_SheetId = sheetInstance.getSheetId();

    CacheImportantData()
    let TargetCellPosition = AdvanceToPosition(Number(DateSplitPosition))

    let DayPartOffset = DonationCellMeta[`${newRecord.DayPart}DayPartOffset`]
    
    TargetCellPosition.startColumnIndex = DayPartOffset
    TargetCellPosition.endColumnIndex = DayPartOffset + 1


    let Requests = []

    for (var row = 0; row < DonationCellMeta.productOrder.length; row++) {

      let CurrentProduct = DonationCellMeta.productOrder[row]

      let PositionToUpdate = BU_MakeUpdateCells()
      PositionToUpdate.range.sheetId = int_SheetId
      PositionToUpdate.range.startRowIndex = TargetCellPosition.startRowIndex + row
      PositionToUpdate.range.endRowIndex = TargetCellPosition.endRowIndex + row - 1

      PositionToUpdate.range.startColumnIndex = TargetCellPosition.startColumnIndex
      PositionToUpdate.range.endColumnIndex = TargetCellPosition.endColumnIndex


      PositionToUpdate.rows = [{
        values: [{
          userEnteredValue: {numberValue: Number(newRecord[CurrentProduct])}
        }]
      }]

      PositionToUpdate.fields = "userEnteredValue"

      Requests.push({updateCells: PositionToUpdate})
    }

    print(Requests)
    SheetsAPI.batchUpdate({requests: Requests}, CurrentSheetIdToUpdate)
  })
}








































