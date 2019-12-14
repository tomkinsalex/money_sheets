 var Section =  Object.create (null, { 
    build: {
      value: function (sheet) {
        this.entries = [];
        this.dateRanges = [];
        this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet);
        return this;
      }
    }
    
  });

var configInfo = { 
  sheet: "Config",
  range: "A1:G10"
}

var state = {
  biweekly: null,
  monthly: null,
  quarterly: null,
  lastQuarterlyRun: null,
  lastRun: null,
  currentRun: null,
  responses: [],
  catsDict: {},
  catsOrder: []
}

var SWITCH = {
  rangeToDict: rangeToDict,
  rangeToArray: rangeToArray,
  rangeToArrayofDicts: rangeToArrayofDicts,
  rangeToValue: rangeToValue,
  datesToArrayWithRow: datesToArrayWithRow
}

var COL_OFFSET = 4;

function mainv2() {
  
  var rn = new Date();
  var startOfMonth = new Date(Date.UTC(rn.getUTCFullYear(), rn.getUTCMonth(), 1,0,0,0));
  
  var config = SWITCH.rangeToDict(configInfo.sheet, configInfo.range);
  setup(config);
  
  var startForBM = startOfMonth.valueOf() < state.lastRun.valueOf() ? startOfMonth : new Date(Date.UTC(state.lastRun.getUTCFullYear(), state.lastRun.getUTCMonth(), 1,0,0,0));
  
  calculateBiweeklyMonthly(startForBM);
  outputResults(["biweekly", "monthly"]);
  
  tearDown(config);
  
}

function setup(config){
  for ( var key in config ){
    if ( config[key]["section_mapping"] == null || config[key]["section_mapping"] == "" ){
      state[config[key]["state_mapping"]] = SWITCH[config[key]["funct"]](config[key]["sheet"], config[key]["start"]+':'+config[key]["end"]);
    }else{
      if( state[config[key]["state_mapping"]] == null ){
        state[config[key]["state_mapping"]] = Object.create(Section).build(config[key]["sheet"]);
      }
      state[config[key]["state_mapping"]][config[key]["section_mapping"]] = SWITCH[config[key]["funct"]](config[key]["sheet"], config[key]["start"]+':'+config[key]["end"]);
    }
  }
  state.lastRun = new Date(Date.UTC(state.lastRun.getUTCFullYear(), state.lastRun.getUTCMonth(), state.lastRun.getUTCDate()));
}

function tearDown(config){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  sheet.getRange(config["lastRun"]["start"]+':'+config["lastRun"]["end"]).setValue(new Date());
}

function calculateBiweeklyMonthly(startDate){
  var filtered = [];
  state.responses.forEach(function(response){
    if(response["Timestamp"] != "" && response["Timestamp"] != null){
      response["Timestamp"] = new Date(Date.UTC(response["Timestamp"].getUTCFullYear(), response["Timestamp"].getUTCMonth(), response["Timestamp"].getUTCDate(),0,0,0));
      if(response["Timestamp"].valueOf() >= startDate.valueOf()){
        filtered.push(response);
      }
    }
  });
  state.responses = filtered;
  
  state.biweekly.dateRanges = state.biweekly.dateRanges.filter(function(range){
    return range[0].valueOf() >= startDate.valueOf(); 
  });
  
  state.monthly.dateRanges = state.monthly.dateRanges.filter(function(range){
    return range[0].valueOf() >= startDate.valueOf();
  });
  
  var bCounter = 0;
  var mCounter= 0 ;
  
  state.biweekly.entries.push(deepClone(state.catsDict));
  state.monthly.entries.push(deepClone(state.catsDict));
  
  state.responses.forEach(function(response){
    bCounter = addIfInRange("biweekly", response, bCounter);
    mCounter = addIfInRange("monthly", response, mCounter);
  });
}

function addIfInRange(section, response, sectionCounter){
  if( state[section].dateRanges.length > sectionCounter){
    var cat = !(response["biweekly"] == null || response["biweekly"] == "") ? response["biweekly"] : response["monthly"];
    if( state[section].dateRanges[sectionCounter][0].valueOf() <= response["Timestamp"].valueOf() && state[section].dateRanges[sectionCounter][1].valueOf() >= response["Timestamp"].valueOf()){
      state[section].entries[sectionCounter][cat]["total"] += response["amount"];
    }else{
      sectionCounter+=1;
      state[section].entries.push(deepClone(state.catsDict));
      state[section].entries[sectionCounter][cat]["total"] += response["amount"];
    }
  }
  return sectionCounter;
}

function outputResults(sections){
  sections.forEach(function(section){
    var colCount;
    for(var i = 0; i < state[section].dateRanges.length; i++){
      colCount=0;
      state.catsOrder.forEach(function(cat){
        if(state[section].entries[i][cat]["is"+section]){
          setValueToCell(state[section]["sheet"], colCount, state[section].dateRanges[i][2], state[section].entries[i][cat]["total"]);
          colCount += 1;
        }
      });
    }
  });
}

function setValueToCell(sheet, col, row, value){
  sheet.getRange(row, col + COL_OFFSET).setValue(value);
}

function deepClone(obj){
  return JSON.parse(JSON.stringify(obj));
}

function rangeToDict(sheetName, range) {
  var matrix = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(range).getValues();
  var columns = matrix.shift();
  var dict_data = {};
  
  matrix.forEach(function(row){ 
    for(var i in columns){
      if ( i == 0 ){
        dict_data[row[0]] = {};
      }else{
        dict_data[row[0]][columns[i]]= row[i];
      }
    }
  });
  return dict_data;
}

function rangeToArrayofDicts(sheetName, range) {
  var matrix = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(range).getValues();
  var columns = matrix.shift();
  var array_data = [];
  
  matrix.forEach(function(row){
    var dict_data = {};
    for(var i in columns){
      dict_data[columns[i]] = row[i];
    }
    array_data.push(dict_data);
  });
  return array_data;
}

function rangeToArray(sheetName, range) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(range).getValues();
}

function datesToArrayWithRow(sheetName, range) {
  var arr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(range).getValues();
  for (var i=0; i < arr.length; i++){
    arr[i][0] = new Date(Date.UTC(arr[i][0].getUTCFullYear(), arr[i][0].getUTCMonth(), arr[i][0].getUTCDate(),0,0,0));
    arr[i][1] = new Date(Date.UTC(arr[i][1].getUTCFullYear(), arr[i][1].getUTCMonth(), arr[i][1].getUTCDate(),0,0,0));
    arr[i].push(i + 2); // dates always start on second row
  }
  return arr;
}

function rangeToValue(sheetName, range) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(range).getValue();
}

function getStartOfQuarter(date){
  var month = date.getUTCMonth() + 1;
  switch(Math.ceil(month / 3)){
    case 1:
      return new Date(date.getYear(), 1,1);
     break;
    case 2:
      return new Date(date.getYear(),3,1);
    break;
    case 3:
      return new Date(date.getYear(),6,1);
    break;
    case 4:
      return new Date(date.getYear(),9,1);
    break;
    default:
      return new Date(date.getYear(), 1,1);
  }
}