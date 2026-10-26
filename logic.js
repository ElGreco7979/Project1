// Pure task logic, shared by the browser app and the Node tests.
(function (root) {
  const REPEATS = ["none", "daily", "weekdays", "weekly"];
  const PRIORITY_ORDER = { high: 0, medium: 1, low: 2 };

  function toKey(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  function fromKey(key) {
    const [y, m, d] = key.split("-").map(Number);
    return new Date(y, m - 1, d);
  }

  function addDays(key, n) {
    const date = fromKey(key);
    date.setDate(date.getDate() + n);
    return toKey(date);
  }

  function createTask({ title, notes = "", priority = "medium", date, repeat = "none", category = "", time = "" }, id) {
    const trimmed = (title || "").trim();
    if (!trimmed) throw new Error("Task title is required");
    if (!REPEATS.includes(repeat)) throw new Error(`Unknown repeat: ${repeat}`);
    if (!(priority in PRIORITY_ORDER)) throw new Error(`Unknown priority: ${priority}`);
    if (time && !/^\d{2}:\d{2}$/.test(time)) throw new Error(`Invalid time: ${time}`);
    return {
      id: id || Date.now().toString(36) + Math.random().toString(36).slice(2, 8),
      title: trimmed,
      notes: notes.trim(),
      priority,
      date,
      repeat,
      category: category.trim(),
      time,
      completedOn: {},
      createdAt: new Date().toISOString(),
    };
  }

  // Does the task appear on the given day?
  function occursOn(task, key) {
    if (key < task.date) return false;
    const day = fromKey(key).getDay();
    switch (task.repeat) {
      case "none":
        return key === task.date;
      case "daily":
        return true;
      case "weekdays":
        return day !== 0 && day !== 6;
      case "weekly":
        return day === fromKey(task.date).getDay();
      default:
        return false;
    }
  }

  function isDone(task, key) {
    return Boolean(task.completedOn[key]);
  }

  function toggle(task, key) {
    const completedOn = { ...task.completedOn };
    if (completedOn[key]) delete completedOn[key];
    else completedOn[key] = true;
    return { ...task, completedOn };
  }

  function sortTasks(tasks, key) {
    return [...tasks].sort(
      (a, b) =>
        isDone(a, key) - isDone(b, key) ||
        PRIORITY_ORDER[a.priority] - PRIORITY_ORDER[b.priority] ||
        // Timed tasks first, earliest first.
        (a.time || "99:99").localeCompare(b.time || "99:99") ||
        a.createdAt.localeCompare(b.createdAt)
    );
  }

  function tasksForDay(tasks, key) {
    return sortTasks(tasks.filter((t) => occursOn(t, key)), key);
  }

  // One-off tasks from earlier days that were never completed.
  function overdue(tasks, todayKey) {
    return sortTasks(
      tasks.filter((t) => t.repeat === "none" && t.date < todayKey && !isDone(t, t.date)),
      todayKey
    );
  }

  // Consecutive days (ending today, or yesterday if today isn't done yet)
  // on which a repeating task was completed.
  function streak(task, todayKey) {
    if (task.repeat === "none") return 0;
    let key = todayKey;
    if (!isDone(task, key)) key = addDays(key, -1);
    let count = 0;
    for (let i = 0; i < 3650 && key >= task.date; i++, key = addDays(key, -1)) {
      if (!occursOn(task, key)) continue;
      if (!isDone(task, key)) break;
      count++;
    }
    return count;
  }

  function categories(tasks) {
    return [...new Set(tasks.map((t) => t.category).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  }

  // Monday of the week containing the given day.
  function weekStart(key) {
    return addDays(key, -((fromKey(key).getDay() + 6) % 7));
  }

  function weekSummary(tasks, key) {
    const start = weekStart(key);
    const days = [];
    const byCategory = {};
    for (let i = 0; i < 7; i++) {
      const day = addDays(start, i);
      const dayTasks = tasks.filter((t) => occursOn(t, day));
      const done = dayTasks.filter((t) => isDone(t, day)).length;
      days.push({ key: day, total: dayTasks.length, done });
      for (const t of dayTasks) {
        const c = (byCategory[t.category || "Uncategorized"] ||= { total: 0, done: 0 });
        c.total++;
        if (isDone(t, day)) c.done++;
      }
    }
    const total = days.reduce((n, d) => n + d.total, 0);
    const done = days.reduce((n, d) => n + d.done, 0);
    return { start, end: addDays(start, 6), days, total, done, byCategory };
  }

  // Timed tasks for today whose time has come, that are unfinished and
  // haven't been announced yet (`notified` holds "id@day" strings).
  function dueReminders(tasks, key, nowTime, notified) {
    return tasks.filter(
      (t) => t.time && t.time <= nowTime && occursOn(t, key) && !isDone(t, key) && !notified.has(`${t.id}@${key}`)
    );
  }

  const api = { REPEATS, toKey, fromKey, addDays, createTask, occursOn, isDone, toggle, tasksForDay, overdue, streak,
    categories, weekStart, weekSummary, dueReminders };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.TaskLogic = api;
})(this);
