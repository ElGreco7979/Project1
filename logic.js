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

  function createTask({ title, notes = "", priority = "medium", date, repeat = "none" }, id) {
    const trimmed = (title || "").trim();
    if (!trimmed) throw new Error("Task title is required");
    if (!REPEATS.includes(repeat)) throw new Error(`Unknown repeat: ${repeat}`);
    if (!(priority in PRIORITY_ORDER)) throw new Error(`Unknown priority: ${priority}`);
    return {
      id: id || Date.now().toString(36) + Math.random().toString(36).slice(2, 8),
      title: trimmed,
      notes: notes.trim(),
      priority,
      date,
      repeat,
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

  const api = { REPEATS, toKey, fromKey, addDays, createTask, occursOn, isDone, toggle, tasksForDay, overdue, streak };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.TaskLogic = api;
})(this);
