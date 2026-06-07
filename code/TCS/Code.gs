//TCS

let SPREADSHEET = "1BypCq8YwxvVTfr-ifRqzy5nHMPMDAxp0Bh1KHdOtaTs"
let MASTER = "1cBXkbiJIiG4PyqoLgWRf20SBfjvO9_7f4NNaOm6vRvs"

function myFunction() {
  let times = Sheets.Spreadsheets.Values.batchGet(SPREADSHEET, {ranges: "Private Tutoring Match!G2:G", majorDimension: 2}).valueRanges[0].values[0]
  let types = Sheets.Spreadsheets.Values.batchGet(SPREADSHEET, {ranges: "Private Tutoring Match!G2:G", majorDimension: 2}).valueRanges[0].values[0]
  let a = []
  //console.log(times)
  times.forEach((values) => {
    a.push(Number(values).toFixed(0))
  })
  let names = Sheets.Spreadsheets.Values.batchGet(SPREADSHEET, {ranges: "Private Tutoring Match!B2:B", majorDimension: 2}).valueRanges[0].values[0]
  let b = []
  times.forEach((values) => {
    let g = ""
    for (i=1; i<= values; i++) {
      g+="T"
    }
    b.push(g)
  })
  //console.log(b)
  let fin = names
  let emailsS = Sheets.Spreadsheets.Values.batchGet(MASTER, {ranges: "Senior Credits!E6:E", majorDimension: 2}).valueRanges[0].values[0]
  let emailsJ = Sheets.Spreadsheets.Values.batchGet(MASTER, {ranges: "Junior Credits!E6:E", majorDimension: 2}).valueRanges[0].values[0]
  //console.log(emailsJ)
  //console.log(fin)
  
  let finfinS = []
  emailsS.forEach((v,i) => {
    finfinS.push("")
  })
  fin.forEach((values, index) => {
    try{
    let aawds = emailsS.indexOf(values)
    finfinS[aawds] = b[index]
    } catch (e) {
    }
  })

  let finfinJ = []
  emailsJ.forEach((v,i) => {
    finfinJ.push("")
  })
  fin.forEach((values, index) => {
    try{
    let aawds = emailsJ.indexOf(values)
    finfinJ[aawds] = b[index]
    } catch (e) {
    }
  })

  //console.log(finfinJ)


  let range = Sheets.newValueRange()
    range.range = "Senior Credits!BF6:BF"
    range.values = [finfinS]
    range.majorDimension = "COLUMNS"
    Sheets.Spreadsheets.Values.batchUpdate({
     "valueInputOption": "USER_ENTERED",
      "data": 
        
          range
        
      ,
  }, MASTER)

  range = Sheets.newValueRange()
    range.range = "Junior Credits!BF6:BF"
    range.values = [finfinJ]
    range.majorDimension = "COLUMNS"
    Sheets.Spreadsheets.Values.batchUpdate({
     "valueInputOption": "USER_ENTERED",
      "data": 
        
          range
        
      ,
  }, MASTER)

  console.log("TCS Done!")
}
