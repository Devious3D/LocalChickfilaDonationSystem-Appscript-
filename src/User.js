var FormURL = "https://docs.google.com/forms/d/1Ko66mD38c5BtYVi_NiT0VOPcHfOBcn_m8LpG7TgC5p8/edit"
var FormId = "1Ko66mD38c5BtYVi_NiT0VOPcHfOBcn_m8LpG7TgC5p8"
var FormBackLogLink = "https://docs.google.com/spreadsheets/d/1FtcFbB02-jMV_KDMh2mt2J42sFdxnt04pLTw-7uyEfM/edit?gid=1285455829#gid=1285455829"

var SheetURL = "https://docs.google.com/spreadsheets/d/1cWfpuiEEmTjq0WEjAEfVqxMV2qXoe2ixjWbXUKVM2Iw/edit?gid=343483624#gid=343483624" // Original Sheet
var SheetId = "1cWfpuiEEmTjq0WEjAEfVqxMV2qXoe2ixjWbXUKVM2Iw"

var TestSheetEnvironmentURL = "https://docs.google.com/spreadsheets/d/1dYhjhzjTOuKHqwTW_6GHiFouDipb_VxHsBYCeklMtc8/edit"
var TestSheetEnvironmentId = "1dYhjhzjTOuKHqwTW_6GHiFouDipb_VxHsBYCeklMtc8"
var DebugMode = false

var CurrentSheetURLToUpdate = ""
var CurrentSheetIdToUpdate = ""

//------------------------------------------
var Year = 2025;
var CurrentMonth = "September"; // Make sure to put "" around the name
var NameOfYourSheet = "September"; // Make sure to put "" around the name
//------------------------------------------

function DetermineDebugMode() {
  switch (DebugMode) {
      case true: 
      {
        CurrentSheetURLToUpdate = TestSheetEnvironmentURL
        CurrentSheetIdToUpdate  = TestSheetEnvironmentId
      }
      break;

      case false:
      {
        CurrentSheetURLToUpdate = SheetURL
        CurrentSheetIdToUpdate  = SheetId
      }
      break;
  }
}

function Start() {

  DetermineDebugMode();

  SheetDisplayMain();
}

