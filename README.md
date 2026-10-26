# Daily Tasks

A simple, dependency-free app for tracking your everyday tasks.

## Features

- **Day-by-day view** – step between days with the arrows; jump back to today anytime.
- **Repeating tasks** – once, every day, weekdays, or weekly. Each day's completion is tracked separately, with a 🔥 streak counter for habits.
- **Priorities** – high / medium / low, color-coded and sorted to the top.
- **Categories** – tag tasks (Work, Home, Health…) and filter the list by category.
- **Reminders** – give a task a time and you get a notification (plus an in-page popup) when it's due.
  Reminders only fire while the app is open in a browser tab.
- **Weekly summary** – switch to the Week view for your completion rate, a per-day chart,
  a per-category breakdown, and your habit streaks. Click a day in the chart to jump to it.
- **Overdue** – unfinished one-off tasks from past days show up on today's view.
- **Progress bar**, filters (All / To do / Done), notes, editing.
- **Export / import** your tasks as JSON (data is stored in your browser's localStorage).
- **Installable app** – add it to your phone's home screen or your computer's app list; works offline.
- Light and dark mode, works on mobile.

## Usage

Open `index.html` in any browser. That's it — no build step or server needed.

### Install it as an app

Installing and offline mode need the app to be served over a web address, not opened as a file.
The easiest free option is GitHub Pages:

1. Merge this branch into `main`.
2. On GitHub, go to **Settings → Pages**, set **Source** to *Deploy from a branch*, pick `main` and `/ (root)`, and save.
3. After a minute your app is live at `https://<your-username>.github.io/<repo-name>/`.

Then open that address and install it:

- **Android (Chrome):** menu ⋮ → *Install app* (or tap *Install app* at the bottom of the page).
- **iPhone/iPad (Safari):** Share button → *Add to Home Screen*.
- **Computer (Chrome/Edge):** the install icon in the address bar, or *Install app* at the bottom of the page.

Each device keeps its own tasks; use Export/Import to copy them between devices.

To try it locally: `python3 -m http.server` and open <http://localhost:8000>.

## Tests

```sh
node --test
```
