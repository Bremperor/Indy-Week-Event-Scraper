// ============================================================
// SETUP: In Apps Script, go to Project Settings > Script Properties
// and add:  ANTHROPIC_API_KEY  =  sk-ant-api03-...
//           GOOGLE_API_KEY     =  AIza...
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Events")
    .addItem("Generate", "generateCalendar")
    .addItem("Clear", "clearSheet")
    .addItem("Export Checked to Doc", "exportCheckedToDoc")
    .addItem("Export Unchecked to Doc", "exportUncheckedToDoc")
    .addSeparator()
    .addItem("Test: CitySpark only", "testCitySpark")
    .addItem("Test: HTML sites only", "testHtmlSites")
    .addItem("Test: Google Calendars only", "testGoogleCalendars")
    .addItem("Check Venue Coverage", "checkVenueCoverage")
    .addToUi();
}

function clearSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sheet.clear();
  sheet.deleteColumn(3);
  sheet.insertColumnAfter(2);
}

function decodeHtml(str) {
  if (!str) return "";
  return str
    .replace(/&#8217;/g, "'")
    .replace(/&#8216;/g, "'")
    .replace(/&#8220;/g, '"')
    .replace(/&#8221;/g, '"')
    .replace(/&#038;/g, "&")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#8211;/g, "–")
    .replace(/&#8212;/g, "—");
}

function formatTime(timeStr) {
  if (!timeStr) return "";
  // Match formats: 7pm, 7PM, 7 pm, 7:30pm, 7:30 pm, 7p.m., 7 p.m., 7:30p.m, 7:30 p.m., etc.
  const match = timeStr.match(/^(\d{1,2})(?::(\d{2}))?\s*([ap])\.?m\.?$/i);
  if (!match) return timeStr;
  const hour    = match[1];
  const minutes = match[2] || "00";
  const period  = match[3].toLowerCase() + ".m";
  if (minutes === "00") return `${hour} ${period}`;
  return `${hour}:${minutes} ${period}`;
}

// ============================================================
// MASTER FUNCTION
// ============================================================

function generateCalendar() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheets()[0];
  const apiKey  = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  const now     = new Date();
  const endDate = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

  if (!apiKey) {
    sheet.getRange(1,1).setValue("ERROR: Missing ANTHROPIC_API_KEY in Script Properties.");
    return;
  }

  sheet.clear();
  sheet.getRange(1, 6).setValue(`Last run: ${new Date().toLocaleString()}`);
  sheet.getRange(2, 6).setValue("Manual entry needed:\nTheatre in the Park\nThe Fruit\nSuccotash\nThe Varsity on Franklin\nPSI Theatre\nCary Theatre\nReynolds Industrial Theatre\nLeggett Theatre\nRegular ComedyWorx shows on Friday");
  sheet.getRange(2, 6).setWrap(true);
  sheet.getRange(3, 6).setValue("Double check:\nSharp 9 Gallery\nCat's Cradle\nRubies on Five Points\nMartin Marietta");
  sheet.getRange(3, 6).setWrap(true);

  sheet.getRange(1, 7).setValue("Copy-paste venues").setFontWeight("bold");
  const venueStrings = [
    ". Theatre in the Park, Raleigh.",
    ". The Fruit, Durham.",
    ". Succotash, Durham.",
    ". Varsity Theatre, Chapel Hill.",
    ". PSI Theatre, Durham.",
    ". The Cary Theater, Cary.",
    ". Reynolds Industries Theater, Durham.",
    ". Leggett Theatre, Raleigh.",
    ". ComedyWorx, Raleigh.",
    ". Sharp 9 Gallery, Durham.",
    ". Cat's Cradle, Carrboro.",
    ". Rubies on Five Points, Durham.",
    ". Martin Marietta Center for the Performing Arts, Raleigh.",
  ];
  venueStrings.forEach((v, i) => sheet.getRange(i + 2, 7).setValue(v));

  // ---- COLLECT ----
  const citySparkEvents = fetchCitySparkEvents(now, endDate);
  const htmlEvents      = scrapeHtmlSites(now, endDate);
  const calendarEvents  = scrapeGoogleCalendars(now, endDate);

  Logger.log(`CitySpark: ${citySparkEvents.length} | HTML: ${htmlEvents.length} | Calendar: ${calendarEvents.length}`);

  // ---- AGGREGATE ----
  const allEvents = [...citySparkEvents, ...htmlEvents, ...calendarEvents];

  // ---- DEDUPLICATE ----
  const seen     = new Set();
  const filtered = [];
  for (const event of allEvents) {
    const key = `${event.title}_${event.venue}_${event.parsed.toDateString()}`;
    if (seen.has(key)) continue;
    seen.add(key);
    filtered.push(event);
  }

  Logger.log(`After dedup: ${filtered.length} events`);

  if (filtered.length === 0) {
    sheet.getRange(1,1).setValue("No events found.");
    return;
  }

  // ---- TAG ----
  const numbered = filtered.map((e, i) => `${i + 1}. ${e.title}`).join("\n");
  let tagText    = "";

  try {
    const tagResponse = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      headers: {
        "x-api-key":         apiKey,
        "anthropic-version": "2023-06-01",
        "content-type":      "application/json",
      },
      payload: JSON.stringify({
        model:      "claude-haiku-4-5-20251001",
        max_tokens: 2048,
        messages: [{
          role:    "user",
          content: `You are a label gun. For each event below, output ONLY its number and exactly one category from this list: Music, Stage, Screen, Page, Community, Sport, Other.

Rules:
- Use Other ONLY if you genuinely cannot infer what the event is from its title.
- Sport covers live sporting events and wrestling (e.g. WWE, NHL, NBA).
- Music covers concerts, live performances, and music-related talks.
- Stage covers theatre, comedy, dance, and live performance arts.
- Screen covers film screenings.
- Page covers book readings, author events, and literary gatherings.
- Community covers fairs, expos, and civic events.

Format: "1. Music" then next line "2. Stage" etc. No other words.

${numbered}`
        }]
      }),
      muteHttpExceptions: true,
    });

    const tagData = JSON.parse(tagResponse.getContentText());
    tagText       = tagData.content[0].text.trim();
  } catch(e) {
    Logger.log(`ERROR calling Anthropic API: ${e}. Defaulting all to Music.`);
  }

  const categories = {};
  for (const line of tagText.split("\n")) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    const parts = trimmed.split(".", 2);
    if (parts.length === 2) {
      const idx = parseInt(parts[0].trim()) - 1;
      if (!isNaN(idx)) categories[idx] = parts[1].trim();
    }
  }

  for (let i = 0; i < filtered.length; i++) {
    filtered[i].category = categories[i] || "Music";
  }

  // ---- SORT ----
  const CATEGORY_ORDER = { Music: 0, Stage: 1, Sport: 2, Screen: 3, Community: 4, Page: 5, Other: 6 };

  filtered.sort((a, b) => {
    const aDay = new Date(a.parsed); aDay.setHours(0,0,0,0);
    const bDay = new Date(b.parsed); bDay.setHours(0,0,0,0);
    if (aDay - bDay !== 0) return aDay - bDay;
    const catDiff = (CATEGORY_ORDER[a.category] ?? 99) - (CATEGORY_ORDER[b.category] ?? 99);
    if (catDiff !== 0) return catDiff;
    const stripArticle = t => t.replace(/^(a|an|the)\s+/i, "");
    return stripArticle(a.title).localeCompare(stripArticle(b.title));
  });

  // ---- WRITE ----
  writeToSheet(filtered, sheet);

  // ---- LOG ----
  let logSheet = ss.getSheetByName("Log");
  if (!logSheet) logSheet = ss.insertSheet("Log");
  logSheet.clearContents();
  logSheet.getRange(1, 1).setValue(`Generated: ${new Date().toLocaleString()} — ${filtered.length} events`);

  SpreadsheetApp.getUi().alert(`Done! Generated ${filtered.length} events.`);
}

// ============================================================
// COLLECTION FUNCTIONS — each returns an array of event objects
// ============================================================

function fetchCitySparkEvents(now, endDate) {
  const tz = "America/New_York";

  function fmtDT(d) {
    return Utilities.formatDate(d, tz, "yyyy-MM-dd'T'HH:mm");
  }

  const TARGET_VENUES = new Set([
    "Lenovo Center",
    "Coastal Credit Union Music Park",
    "Red Hat Amphitheater",
    "DPAC",
    "Eno House Artist's Den",
    "Haw River Ballroom",
    "The Pinhook",
    "Succotash",
    "The Durham Hotel",
    "The Velvet Hippo",
    "PlayMakers Repertory Company",
    "Theatre Raleigh",
    "Durham Arts Council",
    "Burning Coal Theatre Company",
    "The Justice Theater Project",
    "Raleigh Little Theater",
    "Theatre in the Park",
    "Leggett Theatre",
    "Goodnights Pop-Up Club",
    "Goodnights Comedy Club",
    "Meymandi Concert Hall",
    "ComedyWorx",
    "The Cary Theater",
    "Varsity Theatre",
    "Marbles Kids Museum",
    "Quail Ridge Books",
    "Letters Community Bookshop",
    "The Ritz",
  ]);

  const url       = "https://portal.cityspark.com/api/events/GetEvents/INDYWeek";
  const allEvents = [];
  let skip        = 0;

  while (true) {
    const payload = {
      ppid: 10080, start: fmtDT(now), end: fmtDT(endDate),
      labels: [], pick: false, tps: null, sparks: false,
      category: [], defFilter: "all", distance: 50,
      lat: 35.9955684, lng: -78.9002077,
      search: "", skip: skip, sort: "Time"
    };

    let response, raw, parsed;
    try {
      response = UrlFetchApp.fetch(url, {
        method: "post", contentType: "application/json",
        payload: JSON.stringify(payload), muteHttpExceptions: true,
      });
      raw    = response.getContentText();
      parsed = JSON.parse(raw);
    } catch(e) {
      Logger.log(`CitySpark ERROR at skip=${skip}: ${e}`);
      break;
    }

    const batch = parsed.Value;
    if (!batch || batch.length === 0) break;

    allEvents.push(...batch);
    Logger.log(`CitySpark fetched ${allEvents.length} so far...`);
    skip += 25;
    Utilities.sleep(1000);
  }

  const seen    = new Set();
  const results = [];

  for (const event of allEvents) {
    const venue = event.Venue || "";
    if (!TARGET_VENUES.has(venue)) continue;

    const pid       = event.PId;
    const raw       = event.DateStart || event.Date;
    if (!raw) continue;

    const uniqueKey = `${pid}_${raw.slice(0, 10)}`;
    if (seen.has(uniqueKey)) continue;
    seen.add(uniqueKey);

    const datePart = raw.slice(0, 16);
    const parsed   = new Date(datePart + ":00");
    if (isNaN(parsed.getTime()) || parsed > endDate || parsed < now) continue;

    if (event.Name.startsWith("Doors")) continue;

    results.push({
      title:      decodeHtml(event.Name),
      venue:      decodeHtml(venue),
      city:       (event.CityState || "").split(",")[0].trim(),
      parsed:     parsed,
      event_time: Utilities.formatDate(parsed, tz, "h:mm a").toLowerCase(),
      category:   null,
    });
  }

  Logger.log(`CitySpark: ${results.length} filtered events`);
  return results;
}

function scrapeHtmlSites(now, endDate) {
  const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";

  const thisMonth     = Utilities.formatDate(now, "America/New_York", "yyyy/MM");
  const nextMonthDate = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const nextMonth     = Utilities.formatDate(nextMonthDate, "America/New_York", "yyyy/MM");

  const site_configs = [
    {
      source: "Shadowbox Studio", url: "https://shadowboxstudio.org/events/",
      venue: "Shadowbox Studio", city: "Durham", headers: {},
      event_container: ".mec-topsec", title_sel: ".mec-event-title a",
      date_sel: ".mec-start-date-label", date_format: "withYear", date_cleanup: "mec",
      time_sel: ".mec-start-time", time_style: "direct",
    },
    {
      source: "The Rialto", url: "https://therialto.com/events/",
      venue: "The Rialto", city: "Raleigh", headers: { "User-Agent": UA },
      event_container: ".eventWrapper", title_sel: ".rhp-event__title--list",
      date_sel: ".singleEventDate", date_format: "noYear",
      time_sel: ".rhp-event__time-text--list", time_style: "show_time",
    },
    {
      source: "Duke Arts", url: "https://arts.duke.edu/events/?_presenter=duke-chapel",
      venue: "Duke Chapel", city: "Durham", headers: {},
      event_container: "article.post-event", title_sel: "h3.title",
      date_sel: ".event-date-alt", date_format: "noYear",
      time_sel: null, time_style: "from_date",
    },
    {
      source: "Koka Booth Amphitheatre", url: "https://www.boothamphitheatre.com/events",
      venue: "Koka Booth Amphitheatre", city: "Cary", headers: { "User-Agent": UA },
      event_container: ".info.clearfix", title_sel: ".title a",
      date_sel: ".m-date__singleDate", date_format: "withYear",
      time_sel: null, time_style: "direct",
    },
    {
      source: "Carolina Theatre", url: "https://carolinatheatre.org/events/",
      venue: "The Carolina Theatre", city: "Durham", headers: { "User-Agent": UA },
      event_container: ".card.eventCard.event", title_sel: ".card__title",
      date_sel: ".event__dateBox", date_format: "noYear",
      time_sel: ".card__info p:first-child", time_style: "direct",
    },
    {
      source: "ArtsCenter", url: "https://artscenterlive.org/performances-events/",
      venue: "The ArtsCenter", city: "Carrboro", headers: { "User-Agent": UA },
      event_container: "li.upcomingperf", title_sel: "a.work-title",
      date_sel: "a.work-title", date_format: "noYear",
      time_sel: null, time_style: "title_contains_date",
    },
    {
      source: "Local 506", url: "https://local506.com/events/",
      venue: "Local 506", city: "Chapel Hill", headers: { "User-Agent": UA },
      event_container: ".eventWrapper", title_sel: ".rhp-event__title--list",
      date_sel: ".singleEventDate", date_format: "noYear",
      time_sel: ".rhp-event__time-text--list", time_style: "show_time",
    },
    {
      source: "Mettlesome", url: "https://thisismettlesome.com/live",
      venue: "Mettlesome Theater", city: "Durham", headers: { "User-Agent": UA },
      event_container: "article.eventlist-event", title_sel: ".eventlist-title-link",
      date_sel: "time.event-date", date_format: "withYear",
      time_sel: ".event-time-12hr-start", time_style: "direct",
    },
    {
      source: "Rubies on Five Points", url: "https://www.shazam.com/event/venue/I63AF3AAC2E3FAC0F",
      venue: "Rubies on Five Points", city: "Durham", headers: { "User-Agent": UA },
      event_container: "[class*='ConcertEventItem-module_container']",
      title_sel: "[class*='ConcertEventItem-module_title']",
      date_sel: "[class*='ConcertEventItem-module_dateText']", date_format: "withYear",
      time_sel: "[class*='ConcertEventItem-module_event'] > div[class*='size-small']:not([class*='venue']):not(:empty)",
      time_style: "direct",
    },
    {
      source: "Flyleaf Books", url: `https://flyleafbooks.com/events/${thisMonth}`,
      venue: "Flyleaf Books", city: "Chapel Hill", headers: { "User-Agent": UA },
      event_container: ".event-list__second--bot", title_sel: ".event-list__title a",
      date_sel: null, date_format: "withYear", time_sel: null, time_style: "flyleaf_time",
    },
    {
      source: "Flyleaf Books", url: `https://flyleafbooks.com/events/${nextMonth}`,
      venue: "Flyleaf Books", city: "Chapel Hill", headers: { "User-Agent": UA },
      event_container: ".event-list__second--bot", title_sel: ".event-list__title a",
      date_sel: null, date_format: "withYear", time_sel: null, time_style: "flyleaf_time",
    },
    {
      source: "Cat's Cradle", url: "https://catscradle.com/events/",
      venue: "DYNAMIC", city: "Carrboro", headers: { "User-Agent": UA },
      event_container: ".eventWrapper", title_sel: ".rhp-event__title--list",
      date_sel: ".singleEventDate", date_format: "noYear",
      time_sel: ".rhp-event__time-text--list", time_style: "show_time",
      venue_sel: ".venueLink",
    },
    {
      source: "Martin Marietta Center", url: "https://www.martinmariettacenter.com/events",
      venue: "Martin Marietta Center for the Performing Arts", city: "Raleigh",
      headers: { "User-Agent": UA }, event_container: ".eventItem.entry.clearfix",
      title_sel: ".title a", date_sel: ".m-date__singleDate", date_format: "withYear",
      time_sel: ".time", time_style: "direct",
    },
    {
      source: "NCMA", url: "https://ncartmuseum.org/mec_calendars/upcoming-events-list/",
      venue: "North Carolina Museum of Art", city: "Raleigh", headers: { "User-Agent": UA },
      event_container: "article.mec-event-article", title_sel: ".mec-event-title a",
      date_sel: ".mec-start-date-label", date_format: "withYear", date_cleanup: "mec",
      time_sel: ".mec-start-time", time_style: "direct",
    },
    {
      source: "Lincoln Theatre", url: "https://lincolntheatre.com/events/",
      venue: "Lincoln Theatre", city: "Raleigh", headers: { "User-Agent": UA },
      event_container: ".eventWrapper", title_sel: ".rhp-event__title--list",
      date_sel: ".singleEventDate", date_format: "noYear",
      time_sel: ".rhp-event__time-text--list", time_style: "show_time",
    },
    {
      source: "Pure Life Theatre at Leggett", url: "https://www.purelifetheatre.com/buy-tickets",
      venue: "Pure Life Theatre",           // fallback
      city: "Raleigh",
      headers: { "User-Agent": UA },
      event_container: 'li[data-hook="event-list-item"]',
      title_sel: '[data-hook="ev-list-item-title"]',
      date_sel: null,                       // We'll pull everything from the full date line
      date_format: "withYear",
      time_sel: '[data-hook="date"]',       // "Apr 24, 2026, 7:30 PM – 10:00 PM"
      time_style: "direct",
      venue_sel: '[data-hook="ev-list-item-location"]',
    },
  ];

  const results = [];

  for (const config of site_configs) {
    let response;
    try {
      response = UrlFetchApp.fetch(config.url, {
        headers: config.headers, muteHttpExceptions: true,
      });
    } catch(e) {
      Logger.log(`Failed to fetch ${config.source}: ${e}`);
      continue;
    }

    if (response.getResponseCode() !== 200) {
      Logger.log(`${config.source} returned ${response.getResponseCode()}`);
      continue;
    }

    const root   = NodeHtmlParser.parse(response.getContentText());
    const events = root.querySelectorAll(config.event_container);
    Logger.log(`${config.source}: found ${events.length} events`);

    for (const event of events) {
      const titleEl = event.querySelector(config.title_sel);
      if (!titleEl) continue;
      let title = decodeHtml(titleEl.innerText.trim());
      if (!title) continue;

      const rawVenue  = config.venue_sel
        ? (event.querySelector(config.venue_sel)?.innerText.trim() || config.venue)
        : config.venue;
      const venueName = decodeHtml(rawVenue);

      const dateEl  = config.date_sel ? event.querySelector(config.date_sel) : null;
      const rawDate = dateEl ? dateEl.innerText.trim() : "";

      let event_time   = "";
      let rawDateClean = rawDate;

      if (config.time_style === "title_contains_date") {
        const match = rawDate.match(/^([A-Za-z]+ \d+)[a-z]*\s*[–-]\s*(.+)$/);
        if (match) { rawDateClean = match[1].trim(); title = match[2].trim(); }
        else rawDateClean = rawDate;

      } else if (config.time_style === "from_date") {
        const parts  = rawDate.split(" at ");
        rawDateClean = parts[0].trim();
        event_time   = parts.length === 2 ? parts[1].trim() : "";

      } else if (config.time_style === "time_after_date") {
        const match = rawDate.match(/^(.+?\d{4})\s+(\d.*)$/);
        if (match) { rawDateClean = match[1].trim(); event_time = match[2].trim(); }
        else rawDateClean = rawDate;

      } else if (config.time_style === "show_time") {
        const timeEl = config.time_sel ? event.querySelector(config.time_sel) : null;
        if (timeEl) {
          const rawTime = timeEl.innerText.trim();
          const match   = rawTime.match(/Show:\s*(.+)/i);
          event_time    = match ? match[1].trim() : rawTime;
        }

      } else if (config.time_style === "regex_time") {
        const timeEl = config.time_sel ? event.querySelector(config.time_sel) : null;
        if (timeEl) {
          const match = timeEl.innerText.match(/\d{1,2}:\d{2}\s*[ap]m/i);
          event_time  = match ? match[0].trim() : "";
        }

      } else if (config.time_style === "flyleaf_time") {
        const items = event.querySelectorAll(".event-list__details--item");
        if (items.length > 0) rawDateClean = items[0].innerText.replace(/Date:/i, "").replace(/^[A-Za-z]+,\s*/, "").trim();
        if (items.length > 1) event_time = items[1].innerText.replace(/Time:/i, "").trim().split(" - ")[0].trim();

      } else if (config.time_style === "direct") {
        const timeEl = config.time_sel ? event.querySelector(config.time_sel) : null;
        if (timeEl) event_time = timeEl.innerText.trim();
      }

      let parsed;
      try {
        if (config.date_cleanup === "mec") {
          rawDateClean = rawDateClean.replace(/([A-Za-z]+)\s+(\d+)\s+(\d{4})/, "$1 $2, $3");
        }
        if (config.date_format === "withYear") {
          if (/\d{4}/.test(rawDateClean)) {
            parsed = new Date(rawDateClean);
          } else {
            rawDateClean = rawDateClean.replace(/^[A-Za-z]+,\s*/, "");
            parsed = new Date(`${rawDateClean} ${now.getFullYear()}`);
            if (parsed < new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000)) {
              parsed = new Date(`${rawDateClean} ${now.getFullYear() + 1}`);
            }
          }
        } else {
          rawDateClean = rawDateClean.replace(/^[A-Za-z]+,\s*/, "");
          parsed = new Date(`${rawDateClean} ${now.getFullYear()}`);
          if (parsed < new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000)) {
            parsed = new Date(`${rawDateClean} ${now.getFullYear() + 1}`);
          }
        }
        if (isNaN(parsed.getTime())) continue;
      } catch(e) { continue; }

      if (parsed > endDate || parsed < now) continue;

      results.push({ title, venue: venueName, city: config.city, parsed, event_time, category: null });
    }
  }

  Logger.log(`HTML sites: ${results.length} events`);
  return results;
}

function scrapeGoogleCalendars(now, endDate) {
  const googleApiKey = PropertiesService.getScriptProperties().getProperty("GOOGLE_API_KEY");

  const calendar_configs = [
    {
      name: "Durham Jazz Workshop",
      calendar_id: "rocovn10uf1ej49sqfrd5kf7do@group.calendar.google.com",
      venue: "Sharp 9 Gallery", city: "Durham",
    },
  ];

  const results = [];

  for (const cal of calendar_configs) {
    const url = `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(cal.calendar_id)}/events?key=${googleApiKey}&timeMin=${now.toISOString()}&timeMax=${endDate.toISOString()}&singleEvents=true&orderBy=startTime`;

    let response;
    try {
      response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    } catch(e) {
      Logger.log(`Failed to fetch ${cal.name}: ${e}`);
      continue;
    }

    if (response.getResponseCode() !== 200) {
      Logger.log(`${cal.name} returned ${response.getResponseCode()}`);
      continue;
    }

    const events = JSON.parse(response.getContentText()).items || [];
    Logger.log(`${cal.name}: found ${events.length} events`);

    for (const event of events) {
      const title    = event.summary || "";
      if (!title) continue;
      const rawStart = event.start?.dateTime || event.start?.date;
      if (!rawStart) continue;
      const parsed = new Date(rawStart);
      if (isNaN(parsed.getTime()) || parsed > endDate || parsed < now) continue;
      const event_time = event.start?.dateTime
        ? Utilities.formatDate(parsed, "America/New_York", "h:mm a").toLowerCase()
        : "";
      results.push({ title, venue: cal.venue, city: cal.city, parsed, event_time, category: null });
    }
  }

  Logger.log(`Google Calendars: ${results.length} events`);
  return results;
}

// ============================================================
// WRITE FUNCTION
// ============================================================

function writeToSheet(filtered, sheet) {
  const tz        = "America/New_York";
  let currentDate = null;
  let currentCat  = null;
  const sheetData = [];

  for (const event of filtered) {
    const dateHeader = Utilities.formatDate(event.parsed, tz, "EEEE dd MMMM yyyy").toUpperCase();
    const eventTime = formatTime(event.event_time);
    
    if (dateHeader !== currentDate) {
      currentDate = dateHeader;
      currentCat  = null;
      sheetData.push({ type: "DATE", col1: "'" + dateHeader, col2: "", col4: "DATE" });
    }

    if (event.category !== currentCat) {
      currentCat = event.category;
      sheetData.push({ type: "CATEGORY", col1: event.category, col2: "", col4: "CATEGORY" });
    }

    sheetData.push({
      type: "EVENT",
      col1: event.title,
      col2: `${eventTime}. ${event.venue}, ${event.city}.`,
      col4: "EVENT"
    });
    sheetData.push({ type: "SPACER", col1: "", col2: "", col4: "" });
  }

  for (let i = 0; i < sheetData.length; i++) {
    const row  = i + 1;
    const data = sheetData[i];

    if (data.type === "DATE") {
      sheet.getRange(row, 1).setValue(data.col1).setFontWeight("bold").setFontSize(12).setBackground("#d9e1f2");
      sheet.getRange(row, 2).setValue("").setBackground("#d9e1f2");
      sheet.getRange(row, 4).setValue("DATE");
    } else if (data.type === "CATEGORY") {
      sheet.getRange(row, 1).setValue(data.col1).setFontWeight("bold").setFontSize(10).setBackground("#f2f2f2").setFontStyle("italic");
      sheet.getRange(row, 2).setValue("").setBackground("#f2f2f2");
      sheet.getRange(row, 4).setValue("CATEGORY");
    } else if (data.type === "EVENT") {
      sheet.getRange(row, 1).setValue(data.col1).setFontWeight("bold");
      sheet.getRange(row, 2).setValue(data.col2);
      sheet.getRange(row, 3).insertCheckboxes();
      sheet.getRange(row, 4).setValue("EVENT");
    }
  }

  sheet.autoResizeColumns(1, 2);
}

// ============================================================
// TEST FUNCTIONS
// ============================================================

function testCitySpark() {
  const now     = new Date();
  const endDate = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
  const results = fetchCitySparkEvents(now, endDate);
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sheet.clear();
  const rows = results.map(e => [
    Utilities.formatDate(e.parsed, "America/New_York", "yyyy-MM-dd"),
    e.title, e.event_time, e.venue, e.city
  ]);
  if (rows.length > 0) sheet.getRange(1, 1, rows.length, 5).setValues(rows);
  SpreadsheetApp.getUi().alert(`CitySpark: ${results.length} events`);
}

function testHtmlSites() {
  const now     = new Date();
  const endDate = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
  const results = scrapeHtmlSites(now, endDate);
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sheet.clear();
  const rows = results.map(e => [
    Utilities.formatDate(e.parsed, "America/New_York", "yyyy-MM-dd"),
    e.title, e.event_time, e.venue, e.city
  ]);
  if (rows.length > 0) sheet.getRange(1, 1, rows.length, 5).setValues(rows);
  SpreadsheetApp.getUi().alert(`HTML sites: ${results.length} events`);
}

function testGoogleCalendars() {
  const now     = new Date();
  const endDate = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
  const results = scrapeGoogleCalendars(now, endDate);
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sheet.clear();
  const rows = results.map(e => [
    Utilities.formatDate(e.parsed, "America/New_York", "yyyy-MM-dd"),
    e.title, e.event_time, e.venue, e.city
  ]);
  if (rows.length > 0) sheet.getRange(1, 1, rows.length, 5).setValues(rows);
  SpreadsheetApp.getUi().alert(`Google Calendars: ${results.length} events`);
}

// ============================================================
// EXPORT FUNCTION
// ============================================================

function exportCheckedToDoc() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const data  = sheet.getDataRange().getValues();

  const doc  = DocumentApp.create(`INDY Events Export — ${new Date().toDateString()}`);
  const body = doc.getBody();
  body.clear();

  const activeDates      = new Set();
  const activeCategories = new Set();
  let lastDate           = null;
  let lastCategory       = null;

  for (let i = 0; i < data.length; i++) {
    const rowType = String(data[i][3]);
    const checked = data[i][2];
    if (rowType === "DATE")     { lastDate = i; continue; }
    if (rowType === "CATEGORY") { lastCategory = i; continue; }
    if (rowType === "EVENT" && checked === true) {
      if (lastDate !== null)     activeDates.add(lastDate);
      if (lastCategory !== null) activeCategories.add(lastCategory);
    }
  }

  for (let i = 0; i < data.length; i++) {
    const title   = String(data[i][0]);
    const details = data[i][1] ? String(data[i][1]) : "";
    const checked = data[i][2];
    const rowType = String(data[i][3]);

    if (rowType === "DATE") {
      if (activeDates.has(i)) body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      continue;
    }
    if (rowType === "CATEGORY") {
      if (activeCategories.has(i)) body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      continue;
    }
    if (rowType === "EVENT" && checked === true) {
      body.appendParagraph(title).setBold(true);
      if (details) body.appendParagraph(details);
      body.appendParagraph("");
    }
  }

  doc.saveAndClose();
  SpreadsheetApp.getUi().alert(`Doc created!\n\n${doc.getUrl()}`);
}

function exportUncheckedToDoc() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const data  = sheet.getDataRange().getValues();

  const doc  = DocumentApp.create(`INDY Events Export (Unchecked) — ${new Date().toDateString()}`);
  const body = doc.getBody();
  body.clear();

  const activeDates      = new Set();
  const activeCategories = new Set();
  let lastDate           = null;
  let lastCategory       = null;

  for (let i = 0; i < data.length; i++) {
    const rowType = String(data[i][3]);
    const checked = data[i][2];
    if (rowType === "DATE")     { lastDate = i; continue; }
    if (rowType === "CATEGORY") { lastCategory = i; continue; }
    if (rowType === "EVENT" && checked === false) {
      if (lastDate !== null)     activeDates.add(lastDate);
      if (lastCategory !== null) activeCategories.add(lastCategory);
    }
  }

  for (let i = 0; i < data.length; i++) {
    const title   = String(data[i][0]);
    const details = data[i][1] ? String(data[i][1]) : "";
    const checked = data[i][2];
    const rowType = String(data[i][3]);

    if (rowType === "DATE") {
      if (activeDates.has(i)) body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      continue;
    }
    if (rowType === "CATEGORY") {
      if (activeCategories.has(i)) body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      continue;
    }
    if (rowType === "EVENT" && checked === false) {
      body.appendParagraph(title).setBold(true);
      if (details) body.appendParagraph(details);
      body.appendParagraph("");
    }
  }

  doc.saveAndClose();
  SpreadsheetApp.getUi().alert(`Doc created!\n\n${doc.getUrl()}`);
}

function checkVenueCoverage() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data  = sheet.getDataRange().getValues();

  const EXPECTED_VENUES = [
    // CitySpark
    "Lenovo Center", "Coastal Credit Union Music Park", "Red Hat Amphitheater",
    "DPAC", "Eno House Artist's Den", "Haw River Ballroom", "The Pinhook",
    "The Durham Hotel", "The Velvet Hippo", "PlayMakers Repertory Company",
    "Theatre Raleigh", "PSI Theatre", "Durham Arts Council", "Burning Coal Theatre Company",
    "The Justice Theater Project", "Raleigh Little Theater", "Theatre in the Park",
    "Leggett Theatre", "Goodnights Pop-Up Club", "Goodnights Comedy Club",
    "Meymandi Concert Hall", "ComedyWorx", "The Cary Theater", "Varsity Theatre",
    "Marbles Kids Museum", "Quail Ridge Books", "Letters Community Bookshop",
    // HTML
    "Shadowbox Studio", "The Rialto", "Duke Chapel", "Koka Booth Amphitheatre",
    "The Carolina Theatre", "The ArtsCenter", "Local 506", "Mettlesome Theater",
    "Rubies on Five Points", "Flyleaf Books", "Cat's Cradle", "Motorco Music Hall",
    "Martin Marietta Center for the Performing Arts", "North Carolina Museum of Art",
    "Lincoln Theatre",
    // Google Calendar
    "Sharp 9 Gallery",
    // Manual
    "The Fruit", "Succotash", "The Varsity on Franklin",
  ];

  // Collect all venues that appear in the sheet
  const foundVenues = new Set();
  for (let i = 0; i < data.length; i++) {
    const details = String(data[i][1]);
    const rowType = String(data[i][3]);
    if (rowType === "EVENT" && details) {
      // Details format is "7:00 pm. Venue Name, City."
      // Format: "7:30 p.m. Venue Name, City." — split on last ". " before venue
      const lastDotSpace = details.lastIndexOf(". ");
      if (lastDotSpace !== -1) {
        const venueCity = details.slice(lastDotSpace + 2, -1); // strip trailing "."
        const commaIdx  = venueCity.lastIndexOf(", ");
        if (commaIdx !== -1) {
          foundVenues.add(venueCity.slice(0, commaIdx).trim());
        }
      }
    }
  }

  const missing = EXPECTED_VENUES.filter(v => !foundVenues.has(v));

  if (missing.length === 0) {
    SpreadsheetApp.getUi().alert("✅ All venues accounted for.");
  } else {
    SpreadsheetApp.getUi().alert(`⚠️ Missing venues (${missing.length}):\n\n${missing.join("\n")}`);
  }
}
