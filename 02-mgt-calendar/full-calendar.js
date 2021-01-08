const provider = window.mgt.Providers.globalProvider;
var mgtGetCalendar = document.querySelector('#mgt-get-calendar');
var globalEvents = [];

mgtGetCalendar.addEventListener('templateRendered', (e) => { 
  var calendar = new FullCalendar.Calendar(e.detail.element, {
    initialView: 'dayGridMonth',
    initialEvents: e.detail.context,
    themeSystem: 'standard',
    height: 800,
    events: function(info, successCallback, failureCallback) { 
      try {
        fetchGraphEvents(info.start, info.end, e.path[0].dataset.recurring === "true").then((events) => {
          successCallback(buildCalendarEvents(events));
        });        
      } catch {
        failureCallback();
      }
    }
  });
  calendar.render();
});

/**
 * Transforming Graph Events into FullCalendar events
 * @param {Array} events 
 */
function buildCalendarEvents(events) {
  events.map((event) => {
    if(!globalEvents.find(e => e.id == event.id)) {
      globalEvents.push({
        id: event.id,
        title: event.subject,
        start: event.start.dateTime,
        end: event.end.dateTime,
        allDay: event.isAllDay,
        location: event.location.displayName,
        body: event.bodyPreview,
        type: event.type
      });
    }
  });

  return globalEvents;
}

/**
 * Getting the events data from Microsoft Graph based on the info passed by the Full Calendar active view
 * @param {Date} start 
 * @param {Date} end 
 */
async function fetchGraphEvents(start, end, includeRecurringEvents) {
  let graphClient = provider.graph.client;
  var apiUrl = `/me/events?$filter=start/dateTime ge '${start.toISOString()}' and end/dateTime le '${end.toISOString()}' and type eq 'singleInstance'`;
  var events = await graphClient.api(apiUrl).get();
  var recurringEvents = [];

  if(includeRecurringEvents) {
    var masterEvents = await getGraphSeriesMasterEvents(start, end);
    var recurringEventsBatches = await Promise.all(masterEvents.map(async (masterEvent) => {
      return await getGraphRecurringEvents(masterEvent, start, end);
    }));


    recurringEventsBatches.map((event) => {
      recurringEvents = recurringEvents.concat(event.value);
    })
  }

  if(recurringEvents.length > 0) {
    return events.value.concat(recurringEvents);
  } else {
    return events.value;
  }  
}

/**
 * Getting the events data from Microsoft Graph based on the info passed by the Full Calendar active view
 * @param {Date} start 
 * @param {Date} end 
 */
async function getGraphSeriesMasterEvents(start, end) {
  let graphClient = provider.graph.client;
  var apiUrl = `/me/events?$filter=type eq 'seriesMaster'`;
  var masterEvents = await graphClient.api(apiUrl).get();
  var inRangeMasterEvents = [];

  masterEvents.value.map((masterEvent) => {
    var recurrenceStartDate = new Date(masterEvent.recurrence.range.startDate);
    var recurrenceEndDate = new Date(masterEvent.recurrence.range.endDate);
      
    if(start > recurrenceStartDate && (recurrenceEndDate > end || recurrenceEndDate.getTime() == -62135596800000)) {
      inRangeMasterEvents.push(masterEvent);
    }
  });

  return inRangeMasterEvents;  
}

/**
 * Getting the events data from Microsoft Graph based on the info passed by the Full Calendar active view
 * @param {Date} start 
 * @param {Date} end 
 */
async function getGraphRecurringEvents(masterEvent, start, end) {
  let graphClient = provider.graph.client;
  var events = [];
  var apiUrl = `/me/events/${masterEvent.id}/instances?startDateTime=${start.toISOString()}&endDateTime=${end.toISOString()}`;
  events = await graphClient.api(apiUrl).get();

  return events;  
}