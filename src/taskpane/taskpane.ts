/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("HelloWorld").onclick = HelloWorld;
  }
});

var clock = document.getElementById('time');
var endTime = document.getElementById('TimeSelector');

function Clock() {
  var time = new Date();
  var hours = time.getHours().toString();
  var minutes = time.getMinutes().toString();
  var seconds = time.getSeconds().toString();


  if (hours.length < 2) {
    hours = '0' + hours;
  }

  if (minutes.length < 2) {
    minutes = '0' + minutes;
  }

  if (seconds.length < 2) {
    seconds = '0' + seconds;
  }

  var clockStr = hours + ' : ' + minutes + ' : ' + seconds;
  endTime.setAttribute("min", (hours + ":" + minutes));
  console.log(hours + ":" + minutes);

  clock.textContent = clockStr;

}

Clock();
setInterval(Clock, 1000);

export async function HelloWorld() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
