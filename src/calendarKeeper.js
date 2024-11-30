// utils.js
function msToHours(milliseconds) {
  return milliseconds / (1000 * 60 * 60);
}

function isTitleExcluded(title, excludedTitles) {
  return excludedTitles.includes(title);
}

function createDate(baseDate, dayOffset, hour) {
  const date = new Date(baseDate);
  date.setDate(baseDate.getDate() + dayOffset);
  date.setHours(hour, 0, 0, 0);
  return date;
}

// config.js
// CONFIGURATION HERE
const WORK_START = 8;
const WORK_END = 18;
const MEETING_THRESHOLD = 5;
const BLOCK_EVENT_TITLE = "focus time";
const EXCLUDED_TITLES = ["OOO", "Lunch"];

// calendarScript.js
function blockFreeSlotsWhenOverloaded() {
  const calendar = CalendarApp.getCalendarById("primary");
  const today = new Date();

  for (let dayOffset = 0; dayOffset <= 5; dayOffset++) {
    const startOfDay = createDate(today, dayOffset, WORK_START);
    const endOfDay = createDate(today, dayOffset, WORK_END);

    const events = calendar.getEvents(startOfDay, endOfDay);

    let totalMeetingHours = 0;
    let existingBlockEvents = [];
    events.forEach(event => {
      const eventTitle = event.getTitle();
      if (eventTitle === BLOCK_EVENT_TITLE) {
        existingBlockEvents.push(event);
      } else if (!isTitleExcluded(eventTitle, EXCLUDED_TITLES)) {
        totalMeetingHours += msToHours(event.getEndTime() - event.getStartTime());
      }
    });

    if (totalMeetingHours <= MEETING_THRESHOLD) {
      existingBlockEvents.forEach(event => event.deleteEvent());
      continue;
    }

    let lastEventEnd = startOfDay;
    events.forEach(event => {
      if (event.getTitle() !== BLOCK_EVENT_TITLE && event.getStartTime() > lastEventEnd) {
        calendar.createEvent(BLOCK_EVENT_TITLE, lastEventEnd, event.getStartTime());
      }
      if (event.getTitle() !== BLOCK_EVENT_TITLE) {
        lastEventEnd = event.getEndTime();
      }
    });

    if (lastEventEnd < endOfDay) {
      calendar.createEvent(BLOCK_EVENT_TITLE, lastEventEnd, endOfDay);
    }
  }
}