var DonationSheetCache = {
  StartingCellPositions: [],
  StartingCellPositionsInA1: [],
  CellValues: [],
};

function GetMonth(NameOfMonth) {
  let MonthIndexToReturn = null;

  MonthIndex.forEach(function (Value) {
    if (Value.Name != NameOfMonth) {
      return;
    }
    MonthIndexToReturn = Value;
  });

  print(`Getting Month: ${NameOfMonth}`);

  return MonthIndexToReturn;
}

function GetDayOfTheWeek(DayByNumberOrString) {
  let indexToReturn = null;

  DayIndex.forEach(function (Value) {
    if (DayByNumberOrString == Value.Day) {
      indexToReturn = Value;
    }
    if (DayByNumberOrString == Value.Name) {
      indexToReturn = Value;
    }
  });

  print(`Getting day of the week: ${indexToReturn.Name}`);

  return indexToReturn;
}

function AddWeeksToWeekTotaling(Week) {
  for (var i = 1; i <= 7; i++) {
    let CurrentLoopDay = GetDayOfTheWeek(i);
    ProductTotalsByWeekDayAndProduct[CurrentWeek][CurrentLoopDay.Name] = {
      Nuggets: 0,
      Strips: 0,
      Filets: 0,
      Spicy: 0,
    };
  }

  print(`Adding week: ${Week}`);

  //print(ProductTotalsByWeekDayAndProduct)
}

//This Caches all the position to index the cells on the sheet. This Also has to run before "AdvanceToPosition" can be used
function CacheImportantData() {
  print("Trying to Cache A1 positions");

  let startingPosition = U_MakeA1Range();
  startingPosition.startRow = 4;
  startingPosition.endRow = 4;
  startingPosition.startColumn = 0;
  startingPosition.endColumn = 0;

  for (;;) {
    let ToA1 = ConvertToA1Notation(startingPosition);
    DonationSheetCache.StartingCellPositionsInA1.push(ToA1);

    if (DonationSheetCache.StartingCellPositionsInA1.length > 33) {
      break;
    }

    DonationSheetCache.StartingCellPositions.push({
      startRow: startingPosition.startRow,
      endRow: startingPosition.endRow,
      startColumn: startingPosition.startColumn,
      endColumn: startingPosition.endColumn,
    });

    startingPosition.endRow += DonationCellMeta.CellOffset;
    startingPosition.startRow += DonationCellMeta.CellOffset;
  }

  print(`Cell Positions`);
  print(DonationSheetCache.StartingCellPositions);

  DonationSheetCache.CellValues = SheetsAPI.Values.batchGet(
    CurrentSheetIdToUpdate,
    { ranges: DonationSheetCache.StartingCellPositionsInA1 }
  ).valueRanges;
  if (DonationSheetCache.CellValues != null) {
    print(`Cached A1 Positions: ${DonationSheetCache.CellValues}`);
  }
}

function CreateCellPositions() {
  print("Creating Cell Positions");

  let requestCache = [];
  let CurrentOffset = 0;

  for (var i = 0; i < AmountOfCellsToMake; i++) {
    let NewPosition = BU_MakeUpdateCells();
    NewPosition.range.sheetId = int_SheetId;
    NewPosition.range.startRowIndex =
      DonationCellMeta.CellIndexOrigin.startRowIndex + CurrentOffset;
    NewPosition.range.endRowIndex =
      DonationCellMeta.CellIndexOrigin.endRowIndex + CurrentOffset;
    NewPosition.range.startColumnIndex =
      DonationCellMeta.CellIndexOrigin.startColumnIndex;
    NewPosition.range.endColumnIndex =
      DonationCellMeta.CellIndexOrigin.endColumnIndex;

    let stringOutput = `${i + 1}`;
    //`
    //

    NewPosition.rows = [{
        values: [{
            userEnteredValue: { stringValue: stringOutput },
            userEnteredFormat: {
              textFormat: {
                foregroundColor: {red: 1, green: 1, blue: 1}
              }
            }
          },
        ],
      },
    ];

    NewPosition.fields = "userEnteredValue,userEnteredFormat.textFormat";

    requestCache.push({ updateCells: NewPosition });

    CurrentOffset += 4;
  }
  SheetsAPI.batchUpdate({ requests: requestCache }, CurrentSheetIdToUpdate);
  print("Cell Positions Created");
}

//Position is the number index next to the cell
function AdvanceToPosition(Position) {
  print(`Trying to advance to: ${Position}`);

  let PositionCache = DonationSheetCache.StartingCellPositions;
  let PositionToReturn;

  print(`Cell Value  Length: ${DonationSheetCache.CellValues.length}`);

  for (var i = 0; i < DonationSheetCache.CellValues.length; i++) {
    //print(DonationSheetCache.CellValues[i].values)

    if (DonationSheetCache.CellValues[i].values != Position) {
      continue;
    }

    // print(Position);
    // print(PositionCache[i]);

    PositionToReturn = SS_Vector();
    PositionToReturn.startRowIndex = PositionCache[i].startRow - 1;
    PositionToReturn.endRowIndex = PositionCache[i].endRow + 1;
    PositionToReturn.startColumnIndex = PositionCache[i].startColumn; //this is staying the same becauase the positions are located in the A column
    PositionToReturn.endColumnIndex = PositionCache[i].endColumn + 2;

    //print(PositionToReturn)

    break;
  }

  if (PositionToReturn != null) {
    print(`Advanced to: ${Position}`);
  }

  //Returning like this so it can be use in the batch update
  //This is the range of the leftmost donation header: Date
  return PositionToReturn;
}
