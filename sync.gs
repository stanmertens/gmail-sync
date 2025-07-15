// Working hours to sync
var syncHourStart = 6;
var syncHourEnd = 18;

// Days of the week to sync (0 = Sunday, 6 = Saturday)
var syncDayStart = 1;
var syncDayEnd = 5;

// How many days in the past and future to sync
var syncDaysPast = 7;
var syncDaysFuture = 56;

// Offset from start and end
var eventOffsetMinutes = 30;

// Title for busy events
var eventTitle = "Busy";

// Calendar IDs
var myCalendarId = "";
var workCalendarId = "";


function syncPersonalToWorkCalendar() {
    var myCalendar = CalendarApp.getCalendarById(myCalendarId);
    var workCalendar = CalendarApp.getCalendarById(workCalendarId);
    if (!myCalendar || !workCalendar) {
        Logger.log("One of the calendars was not found.");
        return;
    }

    var now = new Date();
    var past = new Date();
    past.setDate(now.getDate() - syncDaysPast);
    var future = new Date();
    future.setDate(now.getDate() + syncDaysFuture);

    // Collect personal events
    var myEvents = myCalendar.getEvents(now, future);
    var myRanges = [];
    for (var i = 0; i < myEvents.length; i++) {
        var event = myEvents[i];
        var start = new Date(event.getStartTime().getTime() - eventOffsetMinutes * 60 * 1000);
        var end = new Date(event.getEndTime().getTime() + eventOffsetMinutes * 60 * 1000);
        myRanges.push({ start: start, end: end });
    }

    if (myRanges.length === 0) {
        Logger.log("No personal events found.");
        return;
    }

    var myRangesMerged = mergeRanges(myRanges);
    var myRangesClipped = clipRanges(myRangesMerged);
    var busyEvents = workCalendar.getEvents(past, future, { search: eventTitle });
    var busyEvents, obsoleteEvents = splitEvents(myRangesClipped, busyEvents, now);
    var busyEvents = filterExisting(myRangesClipped, busyEvents);

    // Create and delete events
    for (var i = 0; i < busyEvents.length; i++) {
        var range = busyEvents[i];
        var event = workCalendar.createEvent(eventTitle, range.start, range.end);
        event.setVisibility(CalendarApp.Visibility.PRIVATE);
        event.removeAllReminders();
        event.setColor(CalendarApp.EventColor.GRAY);
    }
    obsoleteEvents.forEach(event => event.deleteEvent());
    Logger.log("Synchronisation completed.");
}

function filterExisting(ranges, existingEvents) {
    return ranges.filter(range => {
        return !existingEvents.some(event => {
            const eventStart = event.getStartTime().getTime();
            const eventEnd = event.getEndTime().getTime();
            const itemStart = range.start.getTime();
            const itemEnd = range.end.getTime();
            return eventStart === itemStart && eventEnd === itemEnd;
        });
    });
}

function mergeRanges(ranges) {
    var merged = [];
    var current = ranges[0];
    for (var i = 1; i < ranges.length; i++) {
        var next = ranges[i];
        if (next.start <= current.end) {
            if (next.end > current.end) {
                current.end = next.end;
            }
        } else {
            merged.push(current);
            current = next;
        }
    }
    merged.push(current);
    return merged;
}

function clipRanges(ranges) {
    var results = [];
    ranges.forEach(function (range) {
        var start = range.start;
        var end = range.end;
        var cursor = new Date(start);
        cursor.setHours(0, 0, 0, 0);
        while (cursor <= end) {
            var dayOfWeek = cursor.getDay();
            if (dayOfWeek >= syncDayStart && dayOfWeek <= syncDayEnd) {
                var dayStart = new Date(cursor);
                dayStart.setHours(syncHourStart, 0, 0, 0);
                var dayEnd = new Date(cursor);
                dayEnd.setHours(syncHourEnd, 0, 0, 0);
                var clippedStart = new Date(Math.max(start, dayStart));
                var clippedEnd = new Date(Math.min(end, dayEnd));
                if (clippedStart < clippedEnd) {
                    results.push({ start: clippedStart, end: clippedEnd });
                }
            }
            cursor.setDate(cursor.getDate() + 1);
        }
    });
    return results;
}

function splitEvents(ranges, workEvents, now) {
    var futureEvents = workEvents.filter((event) => event.getStartTime() >= now);
    var obsoleteEvents = futureEvents.filter(function (event) {
        if (event.getEndTime() < now) {
            return true;
        }
        return !ranges.some(function (range) {
            var startMatch = event.getStartTime().getTime() === range.start.getTime();
            var endMatch = event.getEndTime().getTime() === range.end.getTime();
            return startMatch && endMatch;
        });
    });
    var currentEvents = futureEvents.filter(function (event) {
        return !obsoleteEvents.includes(event);
    });
    return currentEvents, obsoleteEvents;
}