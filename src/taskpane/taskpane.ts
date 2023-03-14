/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


/* global document, Office, Word */
var currentTime = "00:00:00";

var nextTableKey:number = 0;

var activeTableKey:number = -1;
var activeTableTotalTime:number = -1;
var activeTableWeightTime:number = -1;
var activeTableStartTime:number = -1;

interface HashMap {
  [key: number] : Array<studentConfig>;
} 
let tables: HashMap = {};

interface studentConfig {
  name: string;
  mentor: string;
  class: string;
  timeWeight: number;
  timeUsed: number;
  active: boolean;
  done: boolean;
}



export async function updateAllTables() {
  return Word.run(async (context) => {

    activeTableKey = -1;
    var table:Word.Table = context.document.body.tables.getFirstOrNullObject();
    var updatingTableData:boolean = true;
    //fetching the table
    while(updatingTableData) {
      table.load();
      await context.sync();
      //checks to see if there is a table key, if there is none it generates one and adds the student info to the student hashmap
      //otherwise it updates the info in the hashmap
      var tableKey: number;
      //gets all the values
      var tableContent = table.values;
      var splitValues = tableContent[0][0].split("  -  ");
      if (splitValues.length < 2) {
        tableKey = generateNewTableKey();
        table.getCell(0,0).body.insertText("  -  " + tableKey, Word.InsertLocation.end);
      } else {
        tableKey = +splitValues[1];
      }

      var studentCount:number = table.rowCount - 1;
      var studentTable:Array<studentConfig> = new Array;
      var studentClass:string = tableContent[0][3].slice(0,4);
      var totalTimeWeights:number = 0;
      var hasActiveStudent:boolean = false;

      for (var i = 1; i <= studentCount; i++) {
        totalTimeWeights += +tableContent[i][1]; 
        studentTable.push(
          generateStudentRecord({
            name: tableContent[i][2],
            mentor: tableContent[i][3],
            class: studentClass,
            timeWeight: +tableContent[i][1],
            timeUsed: 0,
            active: (tableContent[i][0].toLowerCase() == "a"),
            done: (tableContent[i][0].toLowerCase() == "v")
          })
        );
        if ((tableContent[i][0].toLowerCase() == "a")) hasActiveStudent = true;
      }
      if (hasActiveStudent) {
        activeTableKey = tableKey;
        activeTableTotalTime = +tableContent[0][3].slice(25,27);
        activeTableWeightTime = activeTableTotalTime/totalTimeWeights;
      }
      tables[tableKey] = studentTable;

      var nextTable:Word.Table = table.getNextOrNullObject();
      await context.sync();

      if (nextTable.isNullObject) {
        updatingTableData = false;
      } else {
        table = nextTable;
      }
    }
  });
}



function getTableKey(wordTable: Word.Table) : number {
  var table: string[][] = wordTable.values;
  var tableKey: number

  var splitValues = table[0][0].split("  -  ");
  if (splitValues.length < 2) {
    tableKey = generateNewTableKey();
    wordTable.getCell(0,0).body.insertText("  -  " + tableKey, Word.InsertLocation.end);
  } else {
    tableKey = +splitValues[1];
  }
  return tableKey;
}



function generateStudentRecord(config: studentConfig) : {name: string, mentor: string, class: string, timeWeight: number, timeUsed: number, active: boolean, done: boolean} {
  let newStudentRecord:studentConfig = {name: "geen", mentor: "geen", class: "geen", timeWeight: 0, timeUsed: 0, active: false, done: false};
  newStudentRecord.name = config.name;
  newStudentRecord.mentor = config.mentor;
  newStudentRecord.class = config.class;
  newStudentRecord.timeWeight = config.timeWeight;
  newStudentRecord.timeUsed = config.timeUsed;
  newStudentRecord.active = config.active;
  newStudentRecord.done = config.done;
  return newStudentRecord;
}



function generateNewTableKey() : number {
  nextTableKey++;
  return nextTableKey;
}



function updateHUD() {
  if (activeTableKey != -1) {
    var studentTable:studentConfig[] = tables[activeTableKey];
    studentTable.forEach(element => {
      if (element.active) {
        document.getElementById("schedule-grid-item-leerling_naam").textContent = element.name;
        document.getElementById("schedule-grid-item-mentor_klas").textContent = element.mentor + " | " + element.class;
        document.getElementById("schedule-grid-item-tijd_leerling").textContent = minutes100ToMinutes60(+element.timeUsed.toFixed(2)) + " (" + minutes100ToMinutes60(+(activeTableWeightTime*element.timeWeight).toFixed(2)) + ")";

        //document.getElementById("schedule-grid-item-onbesproken_leerlingen").textContent = element.name;
        document.getElementById("schedule-grid-item-basistijd_leerling").textContent = minutes100ToMinutes60(activeTableWeightTime);
        //document.getElementById("schedule-grid-item-eindtijd").textContent = "-6";
      }
    });
  }
}


function minutes100ToMinutes60(number: number) : string {
  var truncValue: number = Math.trunc(number);
  var remainder: number = number - truncValue;

  var remainderInMinutes: string = Math.round(remainder * 60).toString();
  if (remainderInMinutes.length < 2) {
    remainderInMinutes = "0" + remainderInMinutes;
  }
  var value: string = truncValue + ":" + remainderInMinutes;
  return value;
}



Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});



function Clock() {
  var curTime = new Date();
  var hours = curTime.getHours().toString();
  var minutes = curTime.getMinutes().toString();
  var seconds = curTime.getSeconds.toString();

  if (hours.length < 2) {
    hours = '0' + hours;
  }
  if (minutes.length < 2) {
    minutes = '0' + minutes;
  }
  if (seconds.length < 2) {
    seconds = '0' + seconds;
  }
  //sets the time to a variable
  currentTime = hours + ':' + minutes + ':' + seconds;

  updateAllTables();
  updateHUD();
}

Clock();
setInterval(Clock, 1000);

//start button action
/*const scheduleStartButton = document.getElementById('scheduleStartButton');
scheduleStartButton?.addEventListener('click', function handleClick(event) {
  startTime = currentTime;
  startTimeElement.textContent = startTime;

  scheduleStartButton.style.display = "none";
});

//end time setter button action
const scheduleEndTimeSetter = document.getElementById('scheduleEndTimeSetter');
scheduleEndTimeSetter?.addEventListener('click', function handleClick(event) {
  var temp = document.getElementById("endTimeSelector") as HTMLInputElement;
  endTime = temp.value;
  
  endTimeElement.textContent = endTime;
  scheduleEndTimeSetter.textContent = "Edit end time"
});*/

//button to load a table
/*const loadTableButton = document.getElementById('loadTableButton');
loadTableButton?.addEventListener('click', function handleClick(event) {
  updateTableHud();
});*/

/*function updateScheduleHud() {
  if (startTime != "00:00" && currentTime != "00:00" && endTime != "00:00") {

    //time items
    var startTimeInMinutes: number = getAmountOfMinutes(startTime);
    var currentTimeInMinutes: number = getAmountOfMinutes(currentTime);
    var endTimeInMinutes: number = getAmountOfMinutes(endTime);

    var passed_hud_time_content: number = (currentTimeInMinutes - startTimeInMinutes);
    var future_hud_time_content: number = (endTimeInMinutes - currentTimeInMinutes);

    var passed_hud_time = document.getElementById("passed_hud_time");
    passed_hud_time.textContent = passed_hud_time_content.toString();

    var passed_time_percent = document.getElementById("passed_time_percent");
    passed_time_percent.textContent = (Math.round(passed_hud_time_content/(passed_hud_time_content+future_hud_time_content)*100)).toString() + "% -";

    var future_hud_time = document.getElementById("future_hud_time");
    future_hud_time.textContent = future_hud_time_content.toString();

    //student items
    updateTableHud();
  }
}*/

function getAmountOfSeconds(time: String):number {

  const splittedTime = time.split(":");
  var hours: number = +splittedTime[0];
  var minutes: number = +splittedTime[1];
  var seconds: number = +splittedTime[2];

  seconds += (minutes*60 + hours*3600)
  return seconds;
}



export async function HelloWorld(debug: string) {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph(debug, Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
