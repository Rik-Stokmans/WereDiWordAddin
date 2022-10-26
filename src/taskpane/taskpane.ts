/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

var clock = document.getElementById('time');
var endTimeSelector = document.getElementById('timeSelector');


function Clock() {
  var time = new Date();
  var hours = time.getHours().toString();
  var minutes = time.getMinutes().toString();

  if (hours.length < 2) {
    hours = '0' + hours;
  }

  if (minutes.length < 2) {
    minutes = '0' + minutes;
  }

  //sets the time to a variable
  var clockStr = hours + ':' + minutes;

  //checks if the timer is still on the same minute
  const prevTime = clock.textContent.split(":");
  if (prevTime[1] != (" "+minutes+" ")) {
    //updates the minimum value for the end time
    if (minutes == " 59 ") endTimeSelector.setAttribute("min", ((hours+1) + ":00"));
    else {
      
    }
  }

  //updates the clock variable to the current time
  clock.textContent = clockStr;

}

Clock();
setInterval(Clock, 1000);

export async function HelloWorld(debug, debug2) {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph(debug + " - " + debug2, Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
