function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu('Calendar Extension')
  .addItem('fetch calendar', 'fetchCalendar')
  .addToUi()
}

function getWeeklyEvents() {
  var currentTime = new Date();
  var startTime = new Date();
  startTime.setDate(currentTime.getDate() - currentTime.getDay() + 1);
  startTime.setHours(0);
  startTime.setMinutes(0);
  startTime.setSeconds(0);
  
  var endTime = new Date();
  endTime.setDate(currentTime.getDate() - currentTime.getDay() + 7);
  endTime.setHours(23);
  endTime.setMinutes(59);
  endTime.setSeconds(59);
  
  Logger.log(endTime)
  
  var events = CalendarApp.getDefaultCalendar().getEvents(startTime, endTime);
  return events.filter(function(event) {
    return (!event.isAllDayEvent() && (event.getMyStatus() == 'OWNER' || event.getMyStatus() == 'YES'));
  });
}

function fetchCalendar() {
  var spreadsheet = SpreadsheetApp.getActive();
  
  var colors = spreadsheet.getRange('A4:A7').getValues();
  for( var i = 0; i < colors.length; i++ ) {
    colors[i] = eval("CalendarApp.EventColor." + colors[i][0]);
  }
    
  var values = [
    [0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0]
  ]
  
  var events = getWeeklyEvents();
  
  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    
    var startTime = event.getStartTime();
    var endTime = event.getEndTime();
    
    var duration = (endTime - startTime) / 3600000;
    
    var color = event.getColor();
    var eventTypeIndex = colors.indexOf(color);
    if (eventTypeIndex == -1) {eventTypeIndex = 0};
    
    var dayIndex = startTime.getDay() - 1;
    if (dayIndex == -1) {dayIndex = 6};
    
    values[eventTypeIndex][dayIndex] += duration;
  }
  
  spreadsheet.getRange('C4:I7').setValues(values);
}
