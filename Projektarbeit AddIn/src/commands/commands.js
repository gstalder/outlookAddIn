/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// net = require("net");

var item;

/* global global, Office, self, window */
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function insertText(event) {
  // Insert as plain text (CoercionType.Text)
  var textToInsert;

  textToInsert = getMessageBody();
  Office.context.mailbox.item.body.setSelectedDataAsync(
    '<b> Test </b>',
    { coercionType: Office.CoercionType.Text},
    function (asyncResult) {
      // Display the result to the user
      if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        app.showNotification("Success", "\"" + textToInsert + "\" inserted successfully.");
      }
      else {
        app.showNotification("Error", "Failed to insert \"" + textToInsert + "\": " + asyncResult.error.message);
      }
    });

  var today = new Date();
  var subject;

  // Customize the subject with today's date.
  subject = 'Synchronisiert mit Raumsystemsteuerung um ' + today.toLocaleDateString();

  Office.context.mailbox.item.subject.setAsync(
    subject,
    { asyncContext: { var1: 1, var2: 2 } },
    function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
      }
      else {
        // Successfully set the subject.
        // Do whatever appropriate for your scenario
        // using the arguments var1 and var2 as applicable.
      }
    }
  );

  event.completed();
}

function sendToServer(event) {
  var response = '';
  $.ajax({ type: "GET",   
           url: "http://192.168.178.100:80",   
           async: false,
           success : function(text)
           {
               response = text;
               console.log("Succes: "+ text)
           }
  });
}

// Write to a div with id='message' on the page.
function write(message) {
  document.getElementById('message').innerText += message;
}

function getMessageBody() {
  var html = `
  <i>Bitte angeben</i>
  <p>
  <table cellpadding="2" cellspacing="2" border="0"> 
  <tr>
    <td>Anzahl&nbsp;Personen
    <td><i>xx</i>
    </td>
  </tr>
  <tr>
    <td>Verpflegung
    <td><i>Ja/Nein</i>
    </td>
  </tr>
  <tr>
    <td>Audio
    <td><i>Ja/Nein</i>
    </td>
  </tr>
  <tr>
    <td>Video
    <td><i>Ja/Nein</i>
    </td>
  </tr>
  <tr>
    <td>Bestuhlung
    <td><i>Theater/U-Form/Schulung</i>
    </td>
  </tr>
</table>  
  `;
  return html;
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
        ? global
        : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;
g.insertText = insertText;
g.sendToServer = sendToServer;
