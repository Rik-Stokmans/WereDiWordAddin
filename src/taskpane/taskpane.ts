/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


/* global document, Office, Word */
var startTime = "00:00";
var currentTime = "00:00";
var endTime = "00:00";




Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

var startTimeElement = document.getElementById("start_time");
var clock = document.getElementById('current_time');
var endTimeElement = document.getElementById("end_time");
var endTimeSelector = document.getElementById('endTimeSelector');


function Clock() {
  var curTime = new Date();
  var hours = curTime.getHours().toString();
  var minutes = curTime.getMinutes().toString();

  if (hours.length < 2) {
    hours = '0' + hours;
  }

  if (minutes.length < 2) {
    minutes = '0' + minutes;
  }

  //sets the time to a variable
  currentTime = hours + ':' + minutes;

  //checks if the timer is still on the same minute
  const prevTime = clock.textContent.split(":");
  if (prevTime[1] != (" "+minutes+" ")) {
    //updates the minimum value for the end time
    if (minutes == "59") endTimeSelector.setAttribute("min", ((hours+1) + ":00"));
    else {
      endTimeSelector.setAttribute("min", (hours + ":" + (minutes+1)));
    }
  }
  //updates the clock variable to the current time
  clock.textContent = currentTime;

  //updates the schedule hud
  updateScheduleHud();
}

Clock();
setInterval(Clock, 1000);

//start button action
const scheduleStartButton = document.getElementById('scheduleStartButton');
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
});

//button to load a table
const loadTableButton = document.getElementById('loadTableButton');
loadTableButton?.addEventListener('click', function handleClick(event) {
  updateTableHud();
});

function updateScheduleHud() {
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
}

function getAmountOfMinutes(time: String):number {

  const splittedTime = time.split(":");
  var hours: number = Number(splittedTime[0]);
  var minutes: number = Number(splittedTime[1]);

  minutes += hours*60;
  return minutes;
  
}

var studentCount: number = 0;
var markedStudentCount: number = 0;
var studentTimeWeightTotal: number = 0;
var studentTimeWeightDone: number = 0;

export async function getTableInfo() {
  return Word.run(async (context) => {
    //resets table data variables
    studentCount = 0;
    markedStudentCount = 0;
    studentTimeWeightTotal = 0;
    studentTimeWeightDone = 0;

    //fetching the table
    var StundentTable: Word.Table = context.document.body.tables.getFirst();
    StundentTable.load();

    await context.sync();

    studentCount = StundentTable.rowCount - 1;

    var values: string[][] = StundentTable.values;

    var iterator: number = studentCount;
    for (var i:number = 1; i <= iterator; i++) {
      var studentTimeWeight:number = +values[i][1];
      studentTimeWeightTotal += studentTimeWeight;

      if (values[i][0].toLowerCase() == "x") {
        markedStudentCount++;
        studentTimeWeightDone += studentTimeWeight;
      } 
      else if (values[i][0] == "-" || values[i][0] == "") {
        studentCount--;
      }
    }
  });
}

function updateTableHud() {
  getTableInfo();

  var passed_hud_students = document.getElementById("passed_hud_students");
  passed_hud_students.textContent = markedStudentCount.toString();

  var future_hud_students = document.getElementById("future_hud_students");
  future_hud_students.textContent = (studentCount - markedStudentCount).toString();
}


export async function HelloWorld(debug: String, debug2: String, debug3: String) {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph(debug + " - " + debug2 + " - " + debug3, Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
