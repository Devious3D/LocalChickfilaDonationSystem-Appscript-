1//This script reads the current month and fills in the results


var int_SheetId = 0;

var SheetsAPI = Sheets.Spreadsheets;
var CurrentSheetName = "";
var AmountOfCellsToMake = 0;
var MonthAsNumber = 0;
var CurrentWeek = 1;
var CurrentDayOfTheWeek = 1;
var functionStandaloneMode = true



//When the sheet is created, it offsets with this value
var DonationMergeMeta = {

  DonationTitleLocation: {
    "startRow": 1,
    "startColumn": 1,
    "endRow": 2,
    "endColumn": 2
  },

  DonationHeadersLocation: {
    startRow: 2,
    endRow: 2,
    startColumn: 2,
    endColumn: 3
  },

  BackgroundColor: {
    "red": .6,
    "green": .6,
    "blue": .6
  },

  BorderColor: {
    "red": 0,
    "green": 0,
    "blue": 0
  },

  DonationHeaderBackgroundColor: {
    red: .5,
    green: .3,
    blue: .3
  },

  DonationHeaderTextColor:  {
    red: 1,
    green: .4,
    blue: 0
  },

  DonationHeaderOrder: ["Date", "Product", "Morning", "Afternoon", "Dinner", "Total"],
  DonationHeadersRange: {
    startRowIndex: 2,
    endRowIndex: 3,
    startColumnIndex: 1,
    endColumnIndex: 2
  },

  HorizontalOffset : 1,
  columnOffset: 1,
  endColumnIndex: 6
}

var DonationCellMeta = {

  productOrder: ["Nuggets", "Filets", "Spicy", "Strips"],
  dayPartOrder: ["Morning", "Afternoon", "Dinner"],
  TotalsForegroundColor: {red: .7, green: .4, blue: 0},

   CellIndexOrigin: {
    startRowIndex: 3,
    endRowIndex: 5,
    startColumnIndex: 0,
    endColumnIndex: 2,
   },

  backgroundColorForTotals: {red: 1, green: 1, blue: .72},

  productTitlesTextColor: {red: 0, green: .3, blue: .3},
  productTitleBackgroundColor: {red: 0.78, green: 0.94, blue: 0.81},


  backgroundColorWhenDayIsSunday: {red: 1, green: 0.78, blue: 0.81},
  foregroundColorWhenDayIsSunday: {red: 1, green: 0.78, blue: 0.81},

  //The Length of the cell by rows
  CellOffset: 4,

  //The Amount of columns to get to the mentioned Section
  DateOffset: 1,
  MorningDayPartOffset: 3,
  AfternoonDayPartOffset: 4,
  DinnerDayPartOffset: 5,
  TotalsOffest: 6,
}

var DonationTotalsByWeekMeta = {

  StartingPosition: {
    startRow: 1,
    endRow: 2,
    startColumn: 8,
    endColumn: 9,
  },

  RowsInOrder: [
    "Products",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Totals",
  ],


  LengthOfSection: 8,
  WidthOfSection: 4,
  SectionOffest: 8
}

//This Stores Cell Ranges so when the totals are set up, the functions reads this
var ProductCellPositionsOrderedByDayPart = {

  Nuggets: {
    Morning: [],
    Afternoon: [],
    Dinner: [],
  },

  Filets: {
    Morning: [],
    Afternoon: [],
    Dinner: [],
  },

  Spicy: {
    Morning: [],
    Afternoon: [],
    Dinner: [],
  },

  Strips: {
    Morning: [],
    Afternoon: [],
    Dinner: [],
  },

}

//The index is the week of the month
//Using Current Week
var ProductTotalsByWeekDayAndProduct = {
  1: {},
  2: {},
  3: {},
  4: {},
  5: {}
}

var ConditionalFormattingRules = null
//=========================================================================================
function GenerateConditionalFormatting() {


  let OriginCell = AdvanceToPosition(1);
  OriginCell.startColumnIndex += DonationCellMeta.MorningDayPartOffset;
  OriginCell.endColumnIndex += DonationCellMeta.MorningDayPartOffset;

  let OriginCellA1 = `${Letters[OriginCell.startColumnIndex]}${OriginCell.startRowIndex + 1}`;

  let LastPosition  = AdvanceToPosition(AmountOfCellsToMake - 1);
  LastPosition.startColumnIndex += DonationCellMeta.DinnerDayPartOffset;
  LastPosition.endColumnIndex += DonationCellMeta.DinnerDayPartOffset;
  LastPosition.startRowIndex += 3;
  LastPosition.endColumnIndex += 4;

  let LastPositionToA1 = `${Letters[LastPosition.startColumnIndex]}${LastPosition.startRowIndex + 1}`;

  print(`${OriginCellA1}:${LastPositionToA1}`);

  var Sheet = GetSheet();
  const Ranges = Sheet.getRange(`${OriginCellA1}:${LastPositionToA1}`);

  Sheet.clearConditionalFormatRules();

  var CellConditionalFormatting = SpreadsheetApp.newConditionalFormatRule();
  CellConditionalFormatting.setRanges([Ranges]);
  CellConditionalFormatting.whenNumberBetween(0, 6);
  CellConditionalFormatting.setGradientMidpointWithValue("#bbd780", SpreadsheetApp.InterpolationType.PERCENTILE, "0");
  CellConditionalFormatting.setGradientMidpointWithValue("#f6b26b",SpreadsheetApp.InterpolationType.PERCENTILE,"4");
  CellConditionalFormatting.setGradientMaxpointWithValue("#e06666", SpreadsheetApp.InterpolationType.PERCENTILE, "6");
  CellConditionalFormatting.build();


  var rules = Sheet.getConditionalFormatRules();
  rules.push(CellConditionalFormatting);

  Sheet.setConditionalFormatRules(rules);  

  print("Applying New Rules");  
}



function CreateDonationTitles() {
  
  print(`Trying to create Main sheet titles`)

  let fullRequests = [];

  //-----------------------------------------------------------------
  //Donation title and formatting
  let DonationTitleMerge = BU_MakeMergeCells();
  DonationTitleMerge.range.sheetId = int_SheetId;
  DonationTitleMerge.range.startRowIndex = 1;
  DonationTitleMerge.range.endRowIndex = 1 + DonationMergeMeta.HorizontalOffset;
  DonationTitleMerge.range.startColumnIndex = DonationMergeMeta.columnOffset;
  DonationTitleMerge.range.endColumnIndex = DonationMergeMeta.endColumnIndex + DonationMergeMeta.columnOffset;
  fullRequests.push({"mergeCells": DonationTitleMerge});

  let SetDonationTitle = BU_MakeUpdateCells();
  SetDonationTitle.range.sheetId = int_SheetId;
  SetDonationTitle.range.startRowIndex = DonationMergeMeta.DonationTitleLocation.startRow;
  SetDonationTitle.range.endRowIndex = DonationMergeMeta.DonationTitleLocation.endRow;
  SetDonationTitle.range.startColumnIndex = DonationMergeMeta.DonationTitleLocation.startColumn;
  SetDonationTitle.range.endColumnIndex = DonationMergeMeta.DonationTitleLocation.endColumn;

  SetDonationTitle.rows = [{
     "values": [{
      "userEnteredValue": {"stringValue": "Donations"}, 
      "userEnteredFormat": {
        "horizontalAlignment": "CENTER", 
        "verticalAlignment": "MIDDLE",
        "textFormat": {"bold": "true"},
        "backgroundColor": DonationMergeMeta.BackgroundColor
        }
      }]
  }]

  SetDonationTitle.fields =  "userEnteredValue,";
  SetDonationTitle.fields = SetDonationTitle.fields + "userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment,";
  SetDonationTitle.fields = SetDonationTitle.fields + "userEnteredFormat.textFormat.bold,";
  SetDonationTitle.fields =  SetDonationTitle.fields + "userEnteredFormat.backgroundColor";
  fullRequests.push({"updateCells": SetDonationTitle});

  let DonationsBorder = BU_MakeUpdateBorder();
  DonationsBorder.range.sheetId = int_SheetId;
  DonationsBorder.range.startRowIndex = 1;
  DonationsBorder.range.endRowIndex = 1 + DonationMergeMeta.HorizontalOffset;
  DonationsBorder.range.startColumnIndex = DonationMergeMeta.columnOffset;
  DonationsBorder.range.endColumnIndex = DonationMergeMeta.endColumnIndex + DonationMergeMeta.columnOffset;
  DonationsBorder.top.color = DonationMergeMeta.BorderColor;
  DonationsBorder.bottom.color = DonationMergeMeta.BorderColor;
  DonationsBorder.left.color = DonationMergeMeta.BorderColor;
  DonationsBorder.right.color = DonationMergeMeta.BorderColor;
  fullRequests.push({"updateBorders": DonationsBorder});

  print("Pushed Borders")
  //-----------------------------------------------------------------


  let DonationHeadersRequestCache = [];

  //Section Headers and Formatting
  for (var i = 0; i < 6; i++) {

    let DonationHeadersRequest = BU_MakeUpdateCells();
    DonationHeadersRequest.range.sheetId = int_SheetId;
    DonationHeadersRequest.range.startRowIndex = DonationMergeMeta.DonationHeadersRange.startRowIndex;
    DonationHeadersRequest.range.endRowIndex = DonationMergeMeta.DonationHeadersRange.endRowIndex;
    DonationHeadersRequest.range.startColumnIndex = DonationMergeMeta.DonationHeadersRange.startColumnIndex + i;
    DonationHeadersRequest.range.endColumnIndex = DonationMergeMeta.DonationHeadersRange.endColumnIndex + i;

    DonationHeadersRequest.rows = [{
      "values": [{
        "userEnteredValue": {"stringValue": DonationMergeMeta.DonationHeaderOrder[i]},
        "userEnteredFormat": {
          "textFormat": { 
            "bold": false,
            "foregroundColor": DonationMergeMeta.DonationHeaderTextColor
          },
          "backgroundColor": DonationMergeMeta.DonationHeaderBackgroundColor,
          
        }
      }]
    }];

    DonationHeadersRequest.fields = "userEnteredValue,userEnteredFormat.textFormat,userEnteredFormat.backgroundColor";

    //Format the headers next

    fullRequests.push({"updateCells": DonationHeadersRequest});
  }

  print("Pushed titles");
  return fullRequests;
}



function CreateNewDonationCell(sheetInstance, Position) {-
  print(`Creating Cell At Position: ${Position}`)

  //let sheetInstance = GetSheet();
  let requestCache = [];

  //Positioning
  //The position is in the row before the date row
  //It will contain a number indicating the position of the cell
  let CellPosition = AdvanceToPosition(Position);
  print(CellPosition)

  //-------------------------------------------------------------------
  //Date
  //Storing the current position
  //Offsetting it by one column to get to to the date section
  let PositionForDate = CellPosition;
  PositionForDate.startColumnIndex += 1;
  PositionForDate.endRowIndex += 2;

  let DateCellMerge = BU_MakeMergeCells()
  DateCellMerge.range.sheetId = int_SheetId;
  DateCellMerge.range.startRowIndex = PositionForDate.startRowIndex;
  DateCellMerge.range.endRowIndex = PositionForDate.endRowIndex;
  DateCellMerge.range.startColumnIndex = PositionForDate.startColumnIndex;
  DateCellMerge.range.endColumnIndex = PositionForDate.endColumnIndex;
  requestCache.push({ mergeCells: DateCellMerge });

  let DateCellUpdate = BU_MakeUpdateCells();
  DateCellUpdate.range.sheetId = int_SheetId;
  DateCellUpdate.range.startRowIndex = PositionForDate.startRowIndex;
  DateCellUpdate.range.endRowIndex =PositionForDate.endRowIndex;
  DateCellUpdate.range.startColumnIndex = PositionForDate.startColumnIndex;
  DateCellUpdate.range.endColumnIndex = PositionForDate.endColumnIndex;

  let dateToInput = `${MonthAsNumber}/${Position}/${Year}`;
  if (MonthAsNumber < 10) {
    dateToInput = "0" + dateToInput;
  }

  DateCellUpdate.rows = [
    {
      values: [
        {
          userEnteredValue: { stringValue: dateToInput },
          userEnteredFormat: {
            horizontalAlignment: "CENTER",
            textFormat: {
              bold: true,
            },
          },
        },
      ],
    },
  ];

  DateCellUpdate.fields = "userEnteredValue,userEnteredFormat.horizontalAlignment,userEnteredFormat.textFormat";
  requestCache.push({ updateCells: DateCellUpdate });

  if (Position != AmountOfCellsToMake) {
    let CellBorders = BU_MakeUpdateBorder();
    CellBorders.range.sheetId = int_SheetId;
    CellBorders.range.startRowIndex = CellPosition.startRowIndex;
    CellBorders.range.endRowIndex =
    CellPosition.endRowIndex + DonationCellMeta.CellOffset;
    CellBorders.range.startColumnIndex = CellPosition.startColumnIndex;
    CellBorders.range.endColumnIndex =
    CellPosition.endColumnIndex + DonationCellMeta.DinnerDayPartOffset;
    CellBorders.top.color = DonationMergeMeta.BorderColor;
    CellBorders.bottom.color = DonationMergeMeta.BorderColor;
    CellBorders.left.color = DonationMergeMeta.BorderColor;
    CellBorders.right.color = DonationMergeMeta.BorderColor;
    requestCache.push({ updateBorders: CellBorders });

    print(CellBorders);

  }

  print("Pushing Dates");

  
  //-------------------------------------------------------------------

  //-------------------------------------------------------------------
  //Product
  let RangeForProduct = CellPosition;
  RangeForProduct.startColumnIndex += 1;
  RangeForProduct.endColumnIndex += 2;


  for (var i = 0; i < 4; i++) {

    let ProductCellUdpate = BU_MakeUpdateCells();
    ProductCellUdpate.range.sheetId = int_SheetId;
    ProductCellUdpate.range.startRowIndex = RangeForProduct.startRowIndex + i;
    ProductCellUdpate.range.endRowIndex = RangeForProduct.endRowIndex + i;
    ProductCellUdpate.range.startColumnIndex = RangeForProduct.startColumnIndex;
    ProductCellUdpate.range.endColumnIndex = RangeForProduct.endColumnIndex;

    ProductCellUdpate.rows = [{
      values: [{
        userEnteredValue: {stringValue: DonationCellMeta.productOrder[i]},
        userEnteredFormat: {
          horizontalAlignment: "CENTER",
          textFormat: {
            bold: true,
            foregroundColor: DonationCellMeta.productTitlesTextColor
          },
          backgroundColor: DonationCellMeta.productTitleBackgroundColor
        }
      }]
    }];

    ProductCellUdpate.fields = "userEnteredValue,userEnteredFormat.horizontalAlignment,userEnteredFormat.textFormat,userEnteredFormat.backgroundColor";

    requestCache.push({updateCells: ProductCellUdpate});    
  }

  print("Pushed Product Titles")

  //-------------------------------------------------------------------
  //Totals for product row
  let OriginProductTotalLocation = CellPosition;
  OriginProductTotalLocation.startColumnIndex += 1;
  OriginProductTotalLocation.endColumnIndex += 2;

  let IsSunday = (GetDayOfTheWeek(CurrentDayOfTheWeek).Name) == "Sunday"

  //print(OriginProductTotalLocation)

  for (var row = 0; row < 4; row++) {

    let RangesForTotal = [];

    let LetterIndex = null
    let RowIndex = null
    
    for (var column = 0; column < 3; column++) {

      //Pushed for later. To Calculate the sum of the product row
      LetterIndex = Letters[(OriginProductTotalLocation.endColumnIndex - OriginProductTotalLocation.startColumnIndex) + column];
      RowIndex = OriginProductTotalLocation.startRowIndex + (row + 1);

      RangesForTotal.push(`${LetterIndex}${RowIndex}`);

      //I want the for loop to contiue over everything else except the cell at the bottom of the table
      //However, I still need the loop to store the cells to calculate the total of the product row

      let ProductTotalCellUpdate = BU_MakeUpdateCells();
      ProductTotalCellUpdate.range.sheetId = int_SheetId;
      ProductTotalCellUpdate.range.startRowIndex = OriginProductTotalLocation.startRowIndex + row;
      ProductTotalCellUpdate.range.endRowIndex = OriginProductTotalLocation.endRowIndex + row;
      ProductTotalCellUpdate.range.startColumnIndex = OriginProductTotalLocation.startColumnIndex + column;
      ProductTotalCellUpdate.range.endColumnIndex = OriginProductTotalLocation.endColumnIndex + column;

      ProductTotalCellUpdate.rows = [{
        values: [{
            userEnteredValue: {stringValue: ""}
        }]
      }];

      ProductTotalCellUpdate.fields = "userEnteredValue";

      //Turning the whole cell red to indicate sunday on the sheet
      if (IsSunday) {

        ProductTotalCellUpdate.rows = [{
          values: [{
            userEnteredValue: {numberValue: "0"},
            userEnteredFormat: {
              backgroundColor: DonationCellMeta.backgroundColorWhenDayIsSunday,
              textFormat: {
                foregroundColor: DonationCellMeta.foregroundColorWhenDayIsSunday
              }
            }
          }]
        }]

        ProductTotalCellUpdate.fields = "userEnteredValue, userEnteredFormat.backgroundColor, userEnteredFormat.textFormat "
      }

      requestCache.push({updateCells: ProductTotalCellUpdate});


      if (Position == AmountOfCellsToMake || IsSunday) {continue;}


      //-------------------------------------------------------------
      //-------------------------------------------------------------
      //Get CellPositions and Convert to A1; store them to count towards totaling at the bottom of the table
      let ProductType = DonationCellMeta.productOrder[row] 
      let dayPart = DonationCellMeta.dayPartOrder[column]

      ProductCellPositionsOrderedByDayPart[ProductType][dayPart].push(`${LetterIndex}${RowIndex}`)
      //print(`${LetterIndex}${RowIndex}`)
      //-------------------------------------------------------------
      //-------------------------------------------------------------
    }


    //Setting up the =SUM() function in the G column
    let SumCell = BU_MakeUpdateCells();
    SumCell.range.sheetId = int_SheetId;
    SumCell.range.startRowIndex = OriginProductTotalLocation.startRowIndex + row;
    SumCell.range.endRowIndex = OriginProductTotalLocation.endRowIndex + row;
    SumCell.range.startColumnIndex = OriginProductTotalLocation.startColumnIndex + 3;
    SumCell.range.endColumnIndex = OriginProductTotalLocation.startColumnIndex + 4;

    let stringToInput = `${RangesForTotal[0]}:${RangesForTotal[RangesForTotal.length - 1]}`

    //---------------------------------------------------------------------------------------------
    let NameOfCurrentDay = GetDayOfTheWeek(CurrentDayOfTheWeek).Name
    let ProductType = DonationCellMeta.productOrder[row] 
    if (Position != AmountOfCellsToMake) { ProductTotalsByWeekDayAndProduct[CurrentWeek][NameOfCurrentDay][ProductType] = stringToInput }
    //---------------------------------------------------------------------------------------------


    //print(GetDayOfTheWeek(CurrentDayOfTheWeek))

    SumCell.rows = [{
      values: [{
        userEnteredValue: {formulaValue: `=SUM(${stringToInput})`}, // Take this sum and save according to what product it is and its day of the week
        userEnteredFormat: {
          textFormat: {
            foregroundColor: {red: .7, green: .4, blue: 0}
          },

          backgroundColor: DonationCellMeta.backgroundColorForTotals
        }
      }]
    }];

    SumCell.fields = "userEnteredValue,userEnteredFormat.textFormat, userEnteredFormat.backgroundColor";
    
    requestCache.push({updateCells: SumCell});
    //print(RangesForTotal);
  }

  print("Pushed Product Totals")

  return requestCache;
}


//Sets up the totals for every product on the sheet
//Totals are at the bottom of the main table
function CreateFinalTotals() {
  print("Creating Final Totals")

  let FinalTotalsLocation = AdvanceToPosition(AmountOfCellsToMake);
  let requests = [];

  //Title
  let TitleText = BU_MakeUpdateCells();
  TitleText.range.sheetId = int_SheetId;
  TitleText.range.startRowIndex = FinalTotalsLocation.startRowIndex;
  TitleText.range.endRowIndex = FinalTotalsLocation.endRowIndex
  TitleText.range.startColumnIndex = 1;
  TitleText.range.endColumnIndex = 2;

  TitleText.rows = [{
    values: [{userEnteredValue: {stringValue: "Total"}}
    ]
  }]

  TitleText.fields = "userEnteredValue";
  requests.push({updateCells: TitleText});

  print("Pushed Titles")
  //------------------------------------------


  //Setting Up Sums
  for (var row = 0; row < 4; row++) {

    for (var column = 0; column < 3; column++) {

      let ProductType = DonationCellMeta.productOrder[row]
      let dayPart = DonationCellMeta.dayPartOrder[column]

      let LocationToUpdate = BU_MakeUpdateCells()
      LocationToUpdate.range.sheetId = int_SheetId
      LocationToUpdate.range.startRowIndex = FinalTotalsLocation.startRowIndex + row
      LocationToUpdate.range.endRowIndex = FinalTotalsLocation.endRowIndex + row
      LocationToUpdate.range.startColumnIndex =  3 + column
      LocationToUpdate.range.endColumnIndex = 4 + column

      LocationToUpdate.rows = [{
        values: [{
            userEnteredValue: {formulaValue: `=SUM(${ProductCellPositionsOrderedByDayPart[ProductType][dayPart]})`},
            userEnteredFormat: {
              textFormat: {
                foregroundColor: DonationCellMeta.TotalsForegroundColor,
              },

              backgroundColor: {red: 1, green: 1, blue: .72}
            }
          },
    
        ]
      }]

      LocationToUpdate.fields = "userEnteredValue, userEnteredFormat.textFormat, userEnteredFormat.backgroundColor"
      requests.push({updateCells: LocationToUpdate})

    }

  }

  print("Pushed Sums")

  return requests;
}

//This creates all the sections and sets up the formulas
function CreateProductTotalsByWeek() {
  print("Creating Weekly Product totals")

  let Requests = []
  let StartingPosition = DonationTotalsByWeekMeta.StartingPosition 
  let AmountOfWeeks = GetMonth(CurrentMonth).Weeks
  let CachedCellLocationsThatAreTotals = []

  for (var Weeks = 1; Weeks <= AmountOfWeeks; Weeks++) {

    //Title
    let MergeRow = BU_MakeMergeCells()
    MergeRow.range.sheetId = int_SheetId
    MergeRow.range.startRowIndex = StartingPosition.startRow 
    MergeRow.range.endRowIndex = StartingPosition.endRow
    MergeRow.range.startColumnIndex = StartingPosition.startColumn 
    MergeRow.range.endColumnIndex = StartingPosition.endColumn + 7
    MergeRow.mergeType = "MERGE_ALL"

    let TitleText = BU_MakeUpdateCells()
    TitleText.range.sheetId = int_SheetId
    TitleText.range.startRowIndex = StartingPosition.startRow
    TitleText.range.endRowIndex = StartingPosition.endRow
    TitleText.range.startColumnIndex = StartingPosition.startColumn
    TitleText.range.endColumnIndex = StartingPosition.endColumn

    TitleText.rows = [{
      values: [{
          userEnteredValue: {stringValue: `Week ${Weeks}`},
          userEnteredFormat: {

            horizontalAlignment: "CENTER",
            backgroundColor: {red: .6, green: .6, blue: .6},
            textFormat: {
              bold: true,
            }
          }
          
      }],
    }]

    TitleText.fields = "userEnteredValue, userEnteredFormat.horizontalAlignment, userEnteredFormat.textFormat, userEnteredFormat.backgroundColor"

    Requests.push({mergeCells: MergeRow})
    Requests.push({updateCells: TitleText})

    print("Pushed Title")

    //Row Titles
    for (var Idx_RowTitles = 0; Idx_RowTitles < DonationTotalsByWeekMeta.LengthOfSection; Idx_RowTitles++) {

      let RowTitlesStartingPosition = BU_MakeUpdateCells()
      RowTitlesStartingPosition.range.sheetId = int_SheetId
      RowTitlesStartingPosition.range.startRowIndex = StartingPosition.startRow + 1
      RowTitlesStartingPosition.range.endRowIndex = StartingPosition.endRow + 1
      RowTitlesStartingPosition.range.startColumnIndex = StartingPosition.startColumn + Idx_RowTitles
      RowTitlesStartingPosition.range.endColumnIndex = StartingPosition.endColumn + Idx_RowTitles

      RowTitlesStartingPosition.rows = [{
        values: [{
          userEnteredValue: {stringValue: DonationTotalsByWeekMeta.RowsInOrder[Idx_RowTitles]},
          userEnteredFormat: {

            horizontalAlignment: "CENTER",
            backgroundColor: DonationMergeMeta.DonationHeaderBackgroundColor,
            textFormat: {
              bold: true,
              foregroundColor: DonationMergeMeta.DonationHeaderTextColor,
            }
          }
        }]
      }]

      RowTitlesStartingPosition.fields = "userEnteredValue, userEnteredFormat.horizontalAlignment, userEnteredFormat.backgroundColor, userEnteredFormat.textFormat"
      Requests.push({updateCells: RowTitlesStartingPosition})
    }

    print("Pushed Day of the week Titles")

    //ProductTitles
    for (var Idx_ProductTitles = 0; Idx_ProductTitles <= 3; Idx_ProductTitles++) {

      let ProductTitlesStartingPosition = BU_MakeUpdateCells()
      ProductTitlesStartingPosition.range.sheetId = int_SheetId
      ProductTitlesStartingPosition.range.startRowIndex = StartingPosition.startRow + 2 + Idx_ProductTitles
      ProductTitlesStartingPosition.range.endRowIndex = StartingPosition.endRow + 2 + Idx_ProductTitles
      ProductTitlesStartingPosition.range.startColumnIndex = StartingPosition.startColumn 
      ProductTitlesStartingPosition.range.endColumnIndex = StartingPosition.endColumn

      ProductTitlesStartingPosition.rows = [{
        values: [{
          userEnteredValue: {stringValue: DonationCellMeta.productOrder[Idx_ProductTitles]},
          userEnteredFormat: {

            horizontalAlignment: "CENTER",
            backgroundColor: DonationCellMeta.productTitleBackgroundColor,
            textFormat: {
              bold: true,
              foregroundColor: DonationCellMeta.productTitlesTextColor
            }
          }
        }]
      }]

      ProductTitlesStartingPosition.fields = "userEnteredValue, userEnteredFormat.horizontalAlignment, userEnteredFormat.backgroundColor, userEnteredFormat.textFormat"
      Requests.push({updateCells: ProductTitlesStartingPosition})
    }

    print("Pushed Product Titles")


    //Formatting Totals Row
    for (var Idx_TotalsRowFormatting = 0; Idx_TotalsRowFormatting < 4; Idx_TotalsRowFormatting++) {
      
      let StartingPositionForTotalsFormatting = BU_MakeUpdateCells()
      StartingPositionForTotalsFormatting.range.sheetId = int_SheetId
      StartingPositionForTotalsFormatting.range.startRowIndex = StartingPosition.startRow + 2 + Idx_TotalsRowFormatting
      StartingPositionForTotalsFormatting.range.endRowIndex = StartingPosition.endRow + 2 + Idx_TotalsRowFormatting
      StartingPositionForTotalsFormatting.range.startColumnIndex = StartingPosition.startColumn + (DonationTotalsByWeekMeta.LengthOfSection - 1)
      StartingPositionForTotalsFormatting.range.endColumnIndex = StartingPosition.endColumn + (DonationTotalsByWeekMeta.LengthOfSection - 1)

      StartingPositionForTotalsFormatting.rows = [{
        values: [{
          userEnteredValue: {stringValue: "0"},
          userEnteredFormat: {
            horizontalAlignment: "RIGHT",
            backgroundColor: DonationCellMeta.backgroundColorForTotals,
            textFormat: {
              bold: true,
                foregroundColor: DonationCellMeta.TotalsForegroundColor
            }
          }
        }]
      }]

      StartingPositionForTotalsFormatting.fields = "userEnteredValue, userEnteredFormat.horizontalAlignment, userEnteredFormat.backgroundColor, userEnteredFormat.textFormat"

      Requests.push({updateCells: StartingPositionForTotalsFormatting})
    }
    print("Formatting Totals")

    //Setting up the =SUM Macro

    //6 days 4 rows
    //Collecting sums form main table
    //This loop goes down then across

    //This loops from 1 - 7 to be usable to GetDayOfTheWeek functions
    //it does not take 0 as an input
    let productCount = 0

    for (var days = 1; days < 7; days++) {
      for (var rows = 1; rows < 5; rows++) {
        
        //Save these Positions to Total at the end of the week table
        let PositionToAffect = BU_MakeUpdateCells()
        PositionToAffect.range.sheetId = int_SheetId
        PositionToAffect.range.startRowIndex = (StartingPosition.startRow + 1 ) + rows
        PositionToAffect.range.endRowIndex = (StartingPosition.endRow + 1 ) + rows
        PositionToAffect.range.startColumnIndex = (StartingPosition.startColumn) + days
        PositionToAffect.range.endColumnIndex = (StartingPosition.endColumn) + days


        let CurrentDay = GetDayOfTheWeek(days)
        
        //Using rows to index from this table, they both so happen to range from 0-4
        //If this errors, that means theres a week with no information in it or theres too many weeks
        let TargetWeekday = ProductTotalsByWeekDayAndProduct[Weeks][CurrentDay.Name][DonationCellMeta.productOrder[rows - 1]]
        productCount += 1
        

        switch (typeof(TargetWeekday) == "number") {
          case true: 
          {
            PositionToAffect.rows = [{
              values: [{
                userEnteredValue: {numberValue: 0},
                userEnteredFormat: {
                  horizontalAlignment: "RIGHT",
                  backgroundColor: DonationCellMeta.backgroundColorWhenDayIsSunday
                }
              }]
            }]
            
            PositionToAffect.fields = "userEnteredValue, userEnteredFormat.backgroundColor, userEnteredFormat.horizontalAlignment"

          } 

          case false:
          {
            PositionToAffect.rows = [{
              values: [{
                userEnteredValue: {formulaValue: `=SUM(${TargetWeekday})`},
                userEnteredFormat: {
                  horizontalAlignment: "RIGHT",
                }
              }]
            }]

            PositionToAffect.fields = "userEnteredValue, userEnteredFormat.horizontalAlignment"
          }
        }


        //Creating the Location of the Cell In A1
        //Starting column is I

        
        Requests.push({updateCells: PositionToAffect})  
      }     
    }

    print("Inputted Totals from Main Sheet")

    //this loop goes across then down
    for (var row = 1; row < 5; row++) {

      let RangeData = {
        LastPosition: BU_MakeUpdateCells(), // When this is inputed, I increment the colum to get the totals rowo
        Ranges: []
      }


      for (var days = 1; days < 7; days++) {
      
        let PositionToAffect = BU_MakeUpdateCells()
        PositionToAffect.range.sheetId = int_SheetId
        PositionToAffect.range.startRowIndex = (StartingPosition.startRow + 1) + row 
        PositionToAffect.range.endRowIndex = (StartingPosition.endRow + 1) + row 
        PositionToAffect.range.startColumnIndex = (StartingPosition.startColumn) + days
        PositionToAffect.range.endColumnIndex = (StartingPosition.endColumn) + days

        if (days == 1 || days == 6) {
          print(days)
          
          let PossibleLocation = 8 + days
          let PosibleRow = PositionToAffect.range.startRowIndex + 1
          RangeData.Ranges[RangeData.Ranges.length] = `${Letters[PossibleLocation]}${PosibleRow}`
          
          //Im expecting to overwrite this value when days == 6
          RangeData.LastPosition = PositionToAffect
        }
      }

      RangeData.LastPosition.range.startColumnIndex += 1
      RangeData.LastPosition.range.endColumnIndex += 1

      RangeData.LastPosition.rows = [{
        values: [{
          userEnteredValue: {formulaValue: `=SUM(${RangeData.Ranges[0]}:${RangeData.Ranges[1]})`}
        }]
      }]

      Requests.push({updateCells: RangeData.LastPosition})

    }

    //Summing Up the totals


    //Offsetting for next Loop
    StartingPosition.startRow += DonationTotalsByWeekMeta.SectionOffest
    StartingPosition.endRow += DonationTotalsByWeekMeta.SectionOffest
  }

  print("Pushed Totals for Days of the Week")

  return Requests
}


//=========================================================================================
//=========================================================================================

function SheetDisplayMain() { 

  let TargetMonth = GetMonth(CurrentMonth)
  AmountOfCellsToMake = TargetMonth.Days + 1
  CurrentDayOfTheWeek = GetDayOfTheWeek(TargetMonth.StartsOn).Day
  MonthAsNumber = TargetMonth.Month
  CurrentWeek = 1

  //Creating the Titles
  //Get the Current month
 
  //print(CurrentDayOfTheWeek)

  let FinalDonationRequests = [];
  let sheetInstance = GetSheet();
  int_SheetId = sheetInstance.getSheetId();

  //Title of the Donation Table
  FinalDonationRequests.push(CreateDonationTitles(sheetInstance)); 
  CreateCellPositions();
  CacheImportantData();
  AddWeeksToWeekTotaling();

  //This local is mainly used to count the weeks for the WeekTotaling function to do its job
  //Some Months dont start on a monday which is indexed as 1.
  //So I need a way to count how many days have past
  //When it reached 7 that means a week past
  let DaysCounted = 1

  //Some Months dont  start on monday, 
  //so this makes the for loop below wait until the current week ends to start counting
  //to then increment the weeks that have past
  let StartCountingDays = false

  for (var i = 1; i <= AmountOfCellsToMake; i++) {
    FinalDonationRequests.push(CreateNewDonationCell(sheetInstance, i));

   // print(`${DayIndex[CurrentDayOfTheWeek]}`)

    switch (CurrentDayOfTheWeek >= 7) {
      case true: 
      {
        CurrentDayOfTheWeek = 1
        StartCountingDays = true
        print(StartCountingDays)
        continue
      }

      case false: 
      {
        CurrentDayOfTheWeek += 1

        if (StartCountingDays) { DaysCounted += 1}    
      }
    }

    if  (DaysCounted == 7) {
        DaysCounted = 1

        CurrentWeek += 1
        AddWeeksToWeekTotaling()
        print("Adding Week")
    }
 

    if (i == AmountOfCellsToMake) {
      FinalDonationRequests.push(CreateFinalTotals());
      break;
    }
  }

  FinalDonationRequests.push(CreateProductTotalsByWeek())
  print(ProductCellPositionsOrderedByDayPart)

  SheetsAPI.batchUpdate({requests: FinalDonationRequests}, CurrentSheetIdToUpdate);

   //GenerateConditionalFormatting();
}

























































































