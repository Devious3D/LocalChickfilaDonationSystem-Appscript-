//----------------------------------------------
//General
var Letters = [
   "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", 
   "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
];

var DayIndex = [

Monday= {
    Day: 1,
    Name: "Monday"
  },
  Tuesday= {
    Day: 2,
    Name: "Tuesday"
  },
  Wednesday= {
    Day: 3,
    Name: "Wednesday"
  },
  Thursday= {
    Day: 4,
    Name: "Thursday"
  },
  Friday= {
    Day: 5,
    Name: "Friday"
  },
  Saturday= {
    Day: 6,
    Name: "Saturday"
  },
  Sunday= {
    Day: 7,
    Name: "Sunday"
  }

]

var MonthIndex = [
  January = {
    Name: "January",
    Month: 1,
    Days: 31,
    Weeks: 4,
    StartsOn: "Wednesday" 
  },
  February = {
    Name: "February",
    Month: 2,
    Days: 28,
    Weeks: 4,
    StartsOn: "Saturday"
  },
  March = {
    Name: "March",
    Month: 3,
    Days: 31,
    Weeks: 5,
    StartsOn: "Saturday"
  },
  April = {
    Name: "April",
    Month: 4,
    Days: 30,
    Weeks: 4,
    StartsOn: "Tuesday"
  },
  May = {
    Name: "May",
    Month: 5,
    Days: 31,
    Weeks: 5,
    StartsOn: "Thursday"
  },
  June = {
    Name: "June",
    Month: 6,
    Days: 30,
    Weeks: 5,
    StartsOn: "Sunday"
  },
  July = {
    Name: "July",
    Month: 7,
    Days: 31,
    Weeks: 3,
    StartsOn: "Tuesday"
  },
  August = {
    Name: "August",
    Month: 8,
    Days: 31,
    Weeks: 5,
    StartsOn: "Friday"
  },
  September = {
    Name: "September",
    Month: 9,
    Days: 30,
    Weeks: 4,
    StartsOn: "Monday"
  },
  October = {
    Name: "October",
    Month: 10,
    Days: 31,
    Weeks: 4,
    StartsOn: "Wednesday"
  },
  November = {
    Name: "November",
    Month: 11,
    Days: 30,
    Weeks: 5,
    StartsOn: "Saturday"
  },
  December = {
    Name: "December",
    Month: 12,
    Days: 31,
    Weeks: 4,
    StartsOn: "Monday"
  }
];


function print(text) {
  Logger.log(text);
}

//SpreadSheet_Vector 
function SS_Vector() {
  return {
    "startRowIndex": 1,
    "startColumnIndex": 1,
    "endRowIndex": 1,
    "endColumnIndex": 1
  }
}

//This is only for converting ranges  
function ConvertToA1Notation(Range) {

  let start_Range = `${Letters[Range.startColumn]}${Range.startRow}`;
  let end_Range = `${Letters[Range.endColumn]}${Range.endRow}`;
  let final_Range = `${NameOfYourSheet}!${start_Range}:${end_Range}`;

  return final_Range;
}

//----------------------------------------------
//Chick-Fil-A related
function GetNewRecordMeta() {
  return {
    Name:    "",
    Date:    "",
    DayPart: "",
    Nuggets: "",
    Strips:  "",
    Filets:  "",
    Spicy:   "",
    Feedback: "",
  }
}
//----------------------------------------------

//----------------------------------------------
//SpreadSheet related
function U_MakeA1Range() {
  return {
    startRow: 0,
    endRow: 0,
    startColumn: 0,
    endColumn: 0
  }
}


function BU_MakeRange() {
  return {
      sheetId: 0,
      startRowIndex: 0,
      endRowIndex: 0,
      startColumnIndex: 0,
      endColumnIndex: 0,
  };
}


function BU_MakeUpdateCells() {
  return {
    "range": BU_MakeRange(),
    "rows": {},
    "fields": "userEnteredValue"

  };
};

function BU_MakeMergeCells() {
  return {
    "range": BU_MakeRange(),
    "mergeType": "MERGE_ALL"
  };
}

function BU_MakeRGB() {
  return {
    red: 0,
    green: 0,
    blue: 0
  }
}

function BU_MakeUpdateBorder() {
  return {
    "range": BU_MakeRange(),
    "top": {"style": "SOLID", "width": 1, "color": BU_MakeRGB()},
    "bottom": {"style": "SOLID", "width": 1, "color": BU_MakeRGB()},
    "left": {"style": "SOLID", "width": 1, "color": BU_MakeRGB()},
    "right": {"style": "SOLID", "width": 1, "color": BU_MakeRGB()},
  }
}



































