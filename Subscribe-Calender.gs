/** @module GCalUtils */


/**
 * Gets all events that occur within a given time range,
 * and that include the specified guest email in the
 * guest list.
 *
 * @param {Calendar} calen Calendar to search
 * @param {Date} start the start of the time range
 * @param {Date} end the end of the time range, non-inclusive
 * @param {String} guestEmail Guest email address to search for
 *
 * @return {CalendarEvent[]} the matching events
 *
 * @see {@link https://developers.google.com/apps-script/reference/calendar/calendar-app?hl=en#getEvents(Date,Date)|CalendarApp.getEvents(Date,Date)}
 * @see {@link https://developers.google.com/apps-script/reference/calendar/calendar-event|Class CalendarEvent}
 */

function myFunction() {
    // get spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("表單回應 1");
  
  // get the data from Google Sheet
  var data = sheet.getRange(sheet.getLastRow(),1,1,2).getValues();
 
  //  create variables
  var email = data[0][1];
 
  // get calendar
  //var masterCal = CalendarApp.getCalendarById("");
  shareCalendar("", email);
}

function shareCalendar( calId, user, role ) {
  role = role || "reader";

  var acl = null;

  // Check whether there is already a rule for this user
  try {
    var acl = Calendar.Acl.get(calId, "user:"+user);
  }
  catch (e) {
    // no existing acl record for this user - as expected. Carry on.
  }

  if (!acl) {
    // No existing rule - insert one.
    acl = {
      "scope": {
        "type": "user",
        "value": user
      },
      "role": role
    };
    var newRule = Calendar.Acl.insert(acl, calId);
  }
  else {
    // There was a rule for this user - update it.
    acl.role = role;
    newRule = Calendar.Acl.update(acl, calId, acl.id)
  }

  return newRule;
}
