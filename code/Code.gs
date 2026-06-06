const MASTER_SPREADSHEET = "1CQY1p67hlvFJRT60_IaL665i7sp1aW2AgawplnoMhSk"
const USER_EMAIL = Session.getActiveUser().getEmail()

let version = "R"

//newly added -> instead of having to go to each line and change, 
//you only have to change the sheet data here
//bug fix due to new column added that push all data 1 column to the right


const NAME_COLUMNS = "B6:C"//must be for sync
const EMAIL_COLUMN_ATTENDANCE = "E:E6"//must be for sync

const DATA_ATTENDANCE_DATA_COLUMNS = "K6:T"
const DATA_ATTENDANCE_DATA_DATES_COLUMNS = "K5:T5"
const DATA_ATTENDANCE_DATA_FINAL_DATA_COLUMNS = "F6:I"

const DATA_MANDATORY_DATA_DATES_COLUMNS = "V5:AE5"
const DATA_MANDATORY_DATA_DATA_COLUMNS = "V6:AE"
const DATA_MANDATORY_DATA_FINAL_DATA_COLUMNS = "V4:AE4"

const DATA_SERVICE_DATA_DATES_COLUMNS = "L5:BB5"
const DATA_SERVICE_DATA_EVENT_NAME_COLUMNS = "L3:BB3"
const DATA_SERVICE_DATA_FINAL_DATA_COLUMNS = "F6:I"
const DATA_SERVICE_DATA_EVENT_TYPE_COLUMNS = "L4:BB4"
const DATA_SERVICE_DATA_EVENT_STATUS_COLUMNS = "L6:BB"
const DATA_SERVICE_DATA_SERVICE_NEEDED_COLUMNS = "F4:H4"



let ERROR_MESSAGE = ""

let USER_INDEX = -1
let USER_GRADE = "none"

let DATA_ATTENDANCE = []
let DATA_ATTENDANCE_PRESENT = []
let DATA_ATTENDANCE_EXCUSED = []
let DATA_ATTENDANCE_ABS = []
let DATA_ATTENDANCE_CUT = []
let DATA_ATTENDANCE_DATES = []

//meeting (attendance) also has the mandatory event names so no need for abs, cut, etc
let DATA_MANDATORY = []
let DATA_MANDATORY_NAMES = []
let DATA_MANDATORY_DATES = []

let DATA_SERVICE_EVENTS = []
let DATA_SERVICE_DATES = []
let DATA_SERVICE_STATUS = []
let DATA_SERVICE_EVENTS_STATUS = []
let DATA_SERVICE_EVENT_TYPE = []
let DATA_SERVICE_NEEDS = []


let HTMLVISUALS_DATA_ATTENDANCE_TABLE = []
let HTMLVISUALS_DATA_MANDATORY_TABLE = []
let HTMLVISUALS_DATA_SERVICE_TABLE = []

let IS_DONE = 0

function myFunction() {
  findUser(USER_EMAIL)
  console.log (USER_GRADE)

  if (USER_INDEX !== -1) {
      ERROR_MESSAGE = ""//"Success; Got Data"

      main()


  } else {
      ERROR_MESSAGE = "You email is not part of NHS or please use you DOE email (@nycstudents.net)"
  }


  IS_DONE = 1
  console.log(ERROR_MESSAGE)
}

function findUser() {
  let emails = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Senior Attendance!"+EMAIL_COLUMN_ATTENDANCE, majorDimension: 2}).valueRanges[0].values[0]
  let emails_ = emails.map((value) => {
    return value.toLowerCase()
  })
  USER_INDEX = emails_.indexOf(USER_EMAIL)
  if (USER_INDEX !== -1) {
    USER_GRADE = 12
    return
  }
  emails = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Junior Attendance!"+EMAIL_COLUMN_ATTENDANCE, majorDimension: 2}).valueRanges[0].values[0]
  emails_ = emails.map((value) => {
    return value.toLowerCase()
  })
  USER_INDEX = emails_.indexOf(USER_EMAIL)
  if (USER_INDEX !== -1) {
    USER_GRADE = 11
  }
}

function main() {
  getAttendanceData()
  getMandatoryData()
  getServiceData()
  

  preformHTMLVisuals()
}

function preformHTMLVisuals() {
//attendance
    for (i = 0; i < DATA_ATTENDANCE_DATES.length; i += 1) {
      HTMLVISUALS_DATA_ATTENDANCE_TABLE.push({
        Date: DATA_ATTENDANCE_DATES[i].toDateString(),
        Status: DATA_ATTENDANCE[i]
      })
    }
    try {
    for (i = 0; i < DATA_ATTENDANCE_DATES.length; i += 1) {
      if (DATA_MANDATORY_DATES[i] !== "" && DATA_MANDATORY[i] !== "" && DATA_MANDATORY_NAMES[i] !== "") {
        HTMLVISUALS_DATA_MANDATORY_TABLE.push({
          Date: DATA_MANDATORY_DATES[i].toDateString(),
          Status: DATA_MANDATORY[i],
          Name: DATA_MANDATORY_NAMES[i]
        })
      }
    }
    } catch (e) {
      console.log(e)
    }

    for (i = 0; i<DATA_SERVICE_EVENTS_STATUS.length; i++) {
      let date = DATA_SERVICE_DATES[i]
      let date_final = ""
      if (typeof(date) === "object") {
          let type = date[0]
          date_final = ""
          for (a = 1; a < date.length; a ++) {
            date_final += date[a].toDateString()
            if (a + 1 !== date.length) {
              if (type === "&") {
                date_final += " and "
              }
              if (type === "TO") {
                date_final += " to "
              }
            }
          }
      } 
      if (!isNaN(new Date(date))){
        try {
         date_final = date.toDateString()
        } catch (e) {
          date_final=date
        }
      }
      //console.log(date_final)
      if (DATA_SERVICE_EVENTS_STATUS[i] !== "") {
        HTMLVISUALS_DATA_SERVICE_TABLE.push({
          Name: DATA_SERVICE_EVENTS[i],
          Date: date_final,
          Type: DATA_SERVICE_EVENT_TYPE[i],
          Status: DATA_SERVICE_EVENTS_STATUS[i]
        })
      }
    }
    console.log(HTMLVISUALS_DATA_SERVICE_TABLE)
}

function getAttendanceData(){
  let final_data = []
  let data = []
  let dates = []
  if (USER_GRADE === 11) {
    dates = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Junior Attendance!"+DATA_ATTENDANCE_DATA_DATES_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
    data = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Junior Attendance!"+DATA_ATTENDANCE_DATA_COLUMNS, majorDimension: 1}).valueRanges[0].values[USER_INDEX]
    final_data = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Junior Attendance!"+DATA_ATTENDANCE_DATA_FINAL_DATA_COLUMNS, majorDimension: 1}).valueRanges[0].values[USER_INDEX]
  }
  if (USER_GRADE === 12) {
    dates = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Senior Attendance!"+DATA_ATTENDANCE_DATA_DATES_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
    data = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Senior Attendance!"+DATA_ATTENDANCE_DATA_COLUMNS, majorDimension: 1}).valueRanges[0].values[USER_INDEX]
    final_data = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Senior Attendance!"+DATA_ATTENDANCE_DATA_FINAL_DATA_COLUMNS, majorDimension: 1}).valueRanges[0].values[USER_INDEX]
  }

      for (i = 0; i < dates.length; i += 2) {
        //set the date as a class date
        if (!isNaN(new Date(dates[i] + "/" + ((Number(dates[i].split("/")[0]) <= 8 && new Date().getMonth() + 1 <= 8) || (Number(dates[i].split("/")[0]) > 8 && new Date().getMonth() + 1 > 8) ? new Date().getFullYear() : new Date().getFullYear() - 1)))) {
        DATA_ATTENDANCE_DATES.push(new Date(dates[i] + "/" + ((Number(dates[i].split("/")[0]) <= 8 && new Date().getMonth() + 1 <= 8) || (Number(dates[i].split("/")[0]) > 8 && new Date().getMonth() + 1 > 8) ? new Date().getFullYear() : new Date().getFullYear() - 1)))

        switch (data[i]) {
          case "P":
            DATA_ATTENDANCE.push("Present")
            break;
          case "A":
            DATA_ATTENDANCE.push("Absent")
            break;
          case "E":
            DATA_ATTENDANCE.push("Excused")
            break;
          case "X":
            DATA_ATTENDANCE.push("Cut")
            break;
          default:
            DATA_ATTENDANCE.push("Not Yet in System")
            break;
        }
        }
      }

      DATA_ATTENDANCE_PRESENT = final_data[0]
      DATA_ATTENDANCE_EXCUSED = final_data[1]
      DATA_ATTENDANCE_ABS = final_data[2]
      DATA_ATTENDANCE_CUT = final_data[3]
  
}
function getMandatoryData(){
  let final_data = [] //event name
  let data = []
  let dates = []
  try {
  if (USER_GRADE === 11) {
    dates = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Junior Attendance!"+DATA_MANDATORY_DATA_DATES_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
    data = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Junior Attendance!"+DATA_MANDATORY_DATA_DATA_COLUMNS, majorDimension: 1}).valueRanges[0].values[USER_INDEX]
    final_data = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Junior Attendance!"+DATA_MANDATORY_DATA_FINAL_DATA_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
  }
  if (USER_GRADE === 12) {
    dates = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Senior Attendance!"+DATA_MANDATORY_DATA_DATES_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
    data = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Senior Attendance!"+DATA_MANDATORY_DATA_DATA_COLUMNS, majorDimension: 1}).valueRanges[0].values[USER_INDEX]
    final_data = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Senior Attendance!"+DATA_MANDATORY_DATA_FINAL_DATA_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
  }
  } catch (e) {
    console.log(e)
    return
  }

      for (i = 0; i < dates.length; i += 2) {
        //set the date as a class date
        if (!isNaN(new Date(dates[i] + "/" + ((Number(dates[i].split("/")[0]) <= 8 && new Date().getMonth() + 1 <= 8) || (Number(dates[i].split("/")[0]) > 8 && new Date().getMonth() + 1 > 8) ? new Date().getFullYear() : new Date().getFullYear() - 1)))) {
        DATA_MANDATORY_DATES.push(new Date(dates[i] + "/" + ((Number(dates[i].split("/")[0]) <= 8 && new Date().getMonth() + 1 <= 8) || (Number(dates[i].split("/")[0]) > 8 && new Date().getMonth() + 1 > 8) ? new Date().getFullYear() : new Date().getFullYear() - 1)))

        switch (data[i]) {
          case "P":
            DATA_MANDATORY.push("Present")
            break;
          case "A":
            DATA_MANDATORY.push("Absent")
            break;
          case "E":
            DATA_MANDATORY.push("Excused")
            break;
          case "X":
            DATA_MANDATORY.push("Cut")
            break;
          case "club":
            DATA_MANDATORY.push("Club")
          default:
            DATA_MANDATORY.push("")
            break;
        }
        }
      }

      DATA_MANDATORY_NAMES = final_data
}
function getServiceData(){
  let final_data = []
  let event_name = []
  let dates = []
  let dates_final = []
  let event_status = []
  let event_type = []
  let service_need = []

  if (USER_GRADE === 11) {
    dates = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Junior Credits!"+DATA_SERVICE_DATA_DATES_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
    event_name = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Junior Credits!"+DATA_SERVICE_DATA_EVENT_NAME_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
    final_data = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Junior Credits!"+DATA_SERVICE_DATA_FINAL_DATA_COLUMNS, majorDimension: 1}).valueRanges[0].values[USER_INDEX]
    event_type = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Junior Credits!"+DATA_SERVICE_DATA_EVENT_TYPE_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
    event_status = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Junior Credits!"+DATA_SERVICE_DATA_EVENT_STATUS_COLUMNS, majorDimension: 1}).valueRanges[0].values[USER_INDEX]
    service_need = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Junior Credits!"+DATA_SERVICE_DATA_SERVICE_NEEDED_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
  }
  if (USER_GRADE === 12) {
    dates = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Senior Credits!"+DATA_SERVICE_DATA_DATES_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
    event_name = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Senior Credits!"+DATA_SERVICE_DATA_EVENT_NAME_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
    final_data = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Senior Credits!"+DATA_SERVICE_DATA_FINAL_DATA_COLUMNS, majorDimension: 1}).valueRanges[0].values[USER_INDEX]
    event_type = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Senior Credits!"+DATA_SERVICE_DATA_EVENT_TYPE_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
    event_status = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Senior Credits!"+DATA_SERVICE_DATA_EVENT_STATUS_COLUMNS, majorDimension: 1}).valueRanges[0].values[USER_INDEX]
    service_need = Sheets.Spreadsheets.Values.batchGet(MASTER_SPREADSHEET, {ranges: "Senior Credits!"+DATA_SERVICE_DATA_SERVICE_NEEDED_COLUMNS, majorDimension: 1}).valueRanges[0].values[0]
  }

  //console.log(dates)
  DATA_SERVICE_EVENTS = event_name
  DATA_SERVICE_STATUS = final_data
  DATA_SERVICE_EVENT_TYPE = event_type
  DATA_SERVICE_NEEDS = service_need
  for (let a = 0; a < dates.length; a++) {
    try {
      let date = dates[a]
      if ((!isNaN(new Date(date))) === true) {
        let newdate = new Date(date)
      newdate.setFullYear((new Date()).getFullYear() - ((Number(date.split("/")[0]) <= 8 && new Date().getMonth() + 1 <= 8) || (Number(dates.split("/")[0]) > 8 && new Date().getMonth() + 1 > 8) ? 0: 1))
      dates_final.push(newdate)
      }
      if (isNaN(new Date(date)) === true) {
        if (date === "") {
          let event_string = event_name[a]
          let event_date_splits = event_string.split("/")
          let final_date_array = ["&"]
        //console.log(event_date_splits)
          for (let c = 0; c < event_date_splits.length - 1; c++) {
            let tempstring = ""
            for (let d = 0; d < event_date_splits[c].length; d++) {
            //console.log(event_date_splits[c].at(event_date_splits[c].length - 1 - d))
              if (isNaN(Number(event_date_splits[c].at(event_date_splits[c].length - 1 - d)))) {
              tempstring += event_date_splits[c].substring(event_date_splits[c].length - 1 - d + 1)
              break;
              }
            }
            if (tempstring === "") {
             tempstring +=   event_date_splits[c]
            }
            //console.log("-------------------")
            tempstring += "/"
            for (let d = 0; d < event_date_splits[c+1].length; d++) {
           // console.log(d)
            //console.log(event_date_splits[c+1].at(d))
              if (isNaN(Number(event_date_splits[c+1].at(d)))) {
                tempstring += event_date_splits[c+1].substring(0, d)
                break;
             }
            }
            let newdate = new Date(tempstring)
            newdate.setFullYear((new Date()).getFullYear() - ((Number(tempstring.split("/")[0]) <= 8 && new Date().getMonth() + 1 <= 8) || (Number(tempstring.split("/")[0]) > 8 && new Date().getMonth() + 1 > 8) ? 0: 1))
            final_date_array.push(newdate)
          }
          dates_final.push(final_date_array)
        } else {
          if (date === "DATE") { break; }
         let event_string = dates[a]
          let event_date_splits = event_string.split("/")
          let final_date_array = ["TO"]
          //console.log(event_date_splits)
         for (let c = 0; c < event_date_splits.length - 1; c++) {
            let tempstring = ""
            for (let d = 0; d < event_date_splits[c].length; d++) {
              //console.log(event_date_splits[c].at(event_date_splits[c].length - 1 - d))
              if (isNaN(Number(event_date_splits[c].at(event_date_splits[c].length - 1 - d)))) {
              tempstring += event_date_splits[c].substring(event_date_splits[c].length - 1 - d + 1)
                break;
              }
            }
            if (tempstring === "") {
              tempstring += event_date_splits[c]
            }
            //console.log("-------------------")
            tempstring += "/"
            for (let d = 0; d < event_date_splits[c+1].length; d++) {
              // console.log(d)
              //console.log(event_date_splits[c+1].at(d))
              if (isNaN(Number(event_date_splits[c+1].at(d)))) {
              tempstring += event_date_splits[c+1].substring(0, d)
                break;
              }
            }
          
            let newdate = new Date(tempstring)
            newdate.setFullYear((new Date()).getFullYear() - ((Number(tempstring.split("/")[0]) <= 8 && new Date().getMonth() + 1 <= 8) || (Number(tempstring.split("/")[0]) > 8 && new Date().getMonth() + 1 > 8) ? 0: 1))
            final_date_array.push(newdate)
          }
          dates_final.push(final_date_array)
        }
      }
    } catch (e) {
      console.log(e)
      dates_final.push(dates[a])
      
    }
  }
  //console.log(dates_final)
  DATA_SERVICE_DATES = dates_final
  let event_status_final = []

  for (let e of event_status) {
    let num_S = (e.match(new RegExp("S", "g")) || []).length
    let num_P = (e.match(new RegExp("P", "g")) || []).length
    let num_T = (e.match(new RegExp("T", "g")) || []).length
    let num_X = (e.match(new RegExp("X", "g")) || []).length
    //https://www.geeksforgeeks.org/javascript/how-to-count-string-occurrence-in-string-using-javascript/
    let str = ""
    if (num_S > 0) {
      str += "Service: " + num_S + " "
    }
    if (num_P > 0) {
      str += "Project: " + num_P + " "
    }
    if (num_T > 0) {
      str += "Tutoring: " + num_T + " "
    }
    if (num_X > 0) {
      str += "Probation: " + num_X + " "
    }
    //console.log(str)
    event_status_final.push(str)
  }


  DATA_SERVICE_EVENTS_STATUS = event_status_final
}
function getVolenteerData(){}
function getProjectData(){}


function read() {

  let master_data = Sheets.Spreadsheets.Values.batchGet("1-Z0j_W-mfJrc5ikeNabsn9Yw6NvtszWokGSDP8ULi6s", {ranges: "Sheet1!A1", majorDimension: 1}).valueRanges[0].values

  let result = getData(master_data[0][0])
  console.log(result)
}

function getData (type) {
  let master_data = Sheets.Spreadsheets.Values.batchGet("1-Z0j_W-mfJrc5ikeNabsn9Yw6NvtszWokGSDP8ULi6s", {ranges: "Sheet1!A:A", majorDimension: 1}).valueRanges[0].values



  console.log(type)

  switch (Number(type)) {
    case 1:
      let data = "A"
      let currentDate = 0
      let values = []
      for (let i = 0; i < master_data.length; i++) {
        try {
          data = master_data[i][0]
          if (data === undefined) {
            continue;
          }
          let date = new Date(data)
          if (!isNaN(date)) {      
            currentDate = date
            console.log("------")
            console.log("Date Set: " + date)
            console.log("Data:")
            continue
          }
          if (currentDate === 0) {
            continue;
          }
          console.log(data)
          values.push(data)
          
        } catch (e) {
          console.error(e)
        }
      } 




      //writing:

      let valuerange = {
        range: "Sheet1!D:D",
        majorDimension: 2,
        values: [values] 
      }
      Sheets.Spreadsheets.Values.batchUpdate({
        valueInputOption: 2,
        data: [valuerange],

      }, DATA_SPREADSHEET)
      return 1
    default:
      return undefined
  }
}







function waitForScriptDone() {
  const max_wait = 10000; 
  const check_time= 500;
  let time = 0;

  while (IS_DONE === 0 && time < max_wait) {
    Utilities.sleep(check_time);
    time += check_time;
  }
}

function doGet(e) {
  myFunction();
  let green = "#61ea99"
  let red = "#ea9999"
  let dark_red = "#e06666"
  let yellow = "#ffd966"


  let html = HtmlService.createTemplateFromFile('Test')
  html.ERROR_MESSAGE = ERROR_MESSAGE;
  html.USER_EMAIL = USER_EMAIL

  html.ATTENDANCE_TABLE = HTMLVISUALS_DATA_ATTENDANCE_TABLE
  html.ATTENDANCE_PRESENT = DATA_ATTENDANCE_PRESENT
  html.ATTENDANCE_ABSENT = DATA_ATTENDANCE_ABS
  html.ATTENDANCE_EXCUSED = DATA_ATTENDANCE_EXCUSED
  html.ATTENDANCE_CUT = DATA_ATTENDANCE_CUT

  html.MANDATORY_TABLE = HTMLVISUALS_DATA_MANDATORY_TABLE

  html.SERVICE_TABLE = HTMLVISUALS_DATA_SERVICE_TABLE
  html.SERVICE_STATUS = DATA_SERVICE_STATUS
  html.SERVICE_NEEDS = DATA_SERVICE_NEEDS

  let needs_left = []
  let needs_left_colors = []
  DATA_SERVICE_STATUS.forEach((value, index, array) => {
    let valuein = (DATA_SERVICE_NEEDS[index] - DATA_SERVICE_STATUS[index]) >= 0 ? (DATA_SERVICE_NEEDS[index] - DATA_SERVICE_STATUS[index]) : 0
    needs_left.push(valuein)
    if (DATA_SERVICE_NEEDS[index] / 2 >= valuein) {
      if (0 === valuein) {
        needs_left_colors.push(green)
      } else {
        needs_left_colors.push(yellow)
      }
    } else {
      needs_left_colors.push(red)
    }
  })
  
  html.SERVICE_NEEDS_LEFT = needs_left
  html.SERVICE_NEEDS_LEFT_COLORS = needs_left_colors
  
  
//https://drive.google.com/file/d/12fwAnokFWgdO6KZKKMWKGQ3xuR9xzDHv/view?usp=drivesdk
  let finalHtml = html.evaluate()
    .setTitle("NHS Credit Tracker")
    .setFaviconUrl("https://drive.google.com/uc?id=12fwAnokFWgdO6KZKKMWKGQ3xuR9xzDHv&export=download&format=png")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    //.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    
  return finalHtml;
}
