# INDY Events Scraper

A Google Apps Script that aggregates events listings for the Triangle NC area 
into a Google Sheet for editorial review.

## How it was built
This script was developed with Claude (Anthropic) over multiple sessions. 
The codebase was built iteratively — discussing architecture, writing configs 
for each venue, and debugging output — rather than written in one sitting.

## What it does
- Pulls events from CitySpark (IndyWeek's calendar API)
- Scrapes 15+ venue websites directly using NodeHtmlParser
- Reads a Google Calendar (Durham Jazz Workshop)
- Tags events by category using Claude Haiku
- Writes everything to a Google Sheet with checkboxes for editorial selection
- Exports selected events to a Google Doc

## Setup
- Requires NodeHtmlParser library: `17hRy4vFVAND6qwEjHLQG5CwYpwNGvlloydh7zgRG2bWhqB1_nCIQQo7B`
- Script Properties needed: `ANTHROPIC_API_KEY`, `GOOGLE_API_KEY`
- New users must authorize via Extensions → Apps Script before using the menu

## Venues covered
CitySpark, HTML scraping, and Google Calendar — see code for full list.

## Video tutorial
[![Watch the demo](https://img.youtube.com/vi/v5Clzg81ylY/0.jpg)](https://www.youtube.com/watch?v=v5Clzg81ylY)
