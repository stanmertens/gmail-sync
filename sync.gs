// Working hours to sync
var syncHourStart = 6;
var syncHourEnd = 20;

// Days of the week to sync (0 = Sunday, 6 = Saturday)
var syncDayStart = 1;
var syncDayEnd = 5;

// How many days in the past and future to sync
var syncDaysPast = 7;
var syncDaysFuture = 56;

var eventOffsetMinutes = 30;
var eventTitle = "Busy";

var personalCalendarId = "stnmertens@gmail.com";
var workCalendarId = "stan.mertens@am-flow.com";

function syncPersonalToWorkCalendar() {
    var personalCalendar = CalendarApp.getCalendarById(personalCalendarId);
    var workCalendar = CalendarApp.getCalendarById(workCalendarId);
    if (!personalCalendar || !workCalendar) {
        Logger.log("One of the calendars was not found.");
        return;
    }

    var now = new Date();
    var past = new Date();
    past.setDate(now.getDate() - syncDaysPast);
    var future = new Date();
    future.setDate(now.getDate() + syncDaysFuture);

    // Collect intervals
    var personalEvents = personalCalendar.getEvents(now, future);
    var intervals = [];
    for (var i = 0; i < personalEvents.length; i++) {
        var event = personalEvents[i];
        var start = new Date(event.getStartTime().getTime() - eventOffsetMinutes * 60 * 1000);
        var end = new Date(event.getEndTime().getTime() + eventOffsetMinutes * 60 * 1000);
        intervals.push({ start: start, end: end });
    }
    if (intervals.length === 0) {
        Logger.log("No personal events found.");
        return;
    }

    var merged = mergeIntervals(intervals);
    var clipped = clipIntervals(merged);
    var workEvents = workCalendar.getEvents(now, future, { search: eventTitle });
    var filtered = filterExisting(clipped, workEvents);

    // Create new events
    for (var i = 0; i < filtered.length; i++) {
        var range = filtered[i];
        var newEvent = workCalendar.createEvent(eventTitle, range.start, range.end);
        newEvent.setVisibility(CalendarApp.Visibility.PRIVATE);
        newEvent.removeAllReminders();
        newEvent.setColor(CalendarApp.EventColor.GRAY);
    }

    // Delete past events
    var workEvents = workCalendar.getEvents(past, now, { search: eventTitle });
    workEvents.forEach(function (e) {
        if (e.getEndTime() < now) {
            e.deleteEvent();
        }
    });

    Logger.log("Synchronisation completed.");
}

function filterExisting(toCheckList, existingEvents) {
    return toCheckList.filter(item => {
        return !existingEvents.some(ev => {
            const eventStart = ev.getStartTime().getTime();
            const eventEnd = ev.getEndTime().getTime();
            const itemStart = item.start.getTime();
            const itemEnd = item.end.getTime();
            return eventStart === itemStart && eventEnd === itemEnd;
        });
    });
}

function mergeIntervals(intervals) {
    var merged = [];
    var current = intervals[0];
    for (var i = 1; i < intervals.length; i++) {
        var next = intervals[i];
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

function clipIntervals(intervals) {
    var results = [];
    intervals.forEach(function (interval) {
        var start = interval.start;
        var end = interval.end;
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