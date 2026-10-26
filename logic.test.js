const test = require("node:test");
const assert = require("node:assert");
const L = require("./logic.js");

// 2026-10-26 is a Monday.
const task = (fields) => L.createTask({ title: "t", date: "2026-10-26", ...fields });

test("createTask requires a title", () => {
  assert.throws(() => L.createTask({ title: "   ", date: "2026-10-26" }));
});

test("one-off tasks occur only on their date", () => {
  const t = task();
  assert.ok(L.occursOn(t, "2026-10-26"));
  assert.ok(!L.occursOn(t, "2026-10-27"));
});

test("repeating tasks follow their schedule", () => {
  assert.ok(L.occursOn(task({ repeat: "daily" }), "2026-11-01"));
  assert.ok(!L.occursOn(task({ repeat: "daily" }), "2026-10-25"));
  assert.ok(!L.occursOn(task({ repeat: "weekdays" }), "2026-10-31")); // Saturday
  assert.ok(L.occursOn(task({ repeat: "weekdays" }), "2026-10-30")); // Friday
  assert.ok(L.occursOn(task({ repeat: "weekly" }), "2026-11-02"));
  assert.ok(!L.occursOn(task({ repeat: "weekly" }), "2026-11-03"));
});

test("toggle completes per day without mutating", () => {
  const t = task({ repeat: "daily" });
  const done = L.toggle(t, "2026-10-26");
  assert.ok(L.isDone(done, "2026-10-26"));
  assert.ok(!L.isDone(done, "2026-10-27"));
  assert.ok(!L.isDone(t, "2026-10-26"));
  assert.ok(!L.isDone(L.toggle(done, "2026-10-26"), "2026-10-26"));
});

test("overdue lists unfinished one-off tasks from earlier days", () => {
  const old = task({ date: "2026-10-20" });
  const finished = L.toggle(task({ date: "2026-10-21" }), "2026-10-21");
  const recurring = task({ date: "2026-10-20", repeat: "daily" });
  assert.deepStrictEqual(
    L.overdue([old, finished, recurring], "2026-10-26").map((t) => t.id),
    [old.id]
  );
});

test("tasksForDay sorts unfinished and high priority first", () => {
  const low = task({ title: "low", priority: "low" });
  const high = task({ title: "high", priority: "high" });
  const doneHigh = L.toggle(task({ title: "done", priority: "high" }), "2026-10-26");
  const titles = L.tasksForDay([low, doneHigh, high], "2026-10-26").map((t) => t.title);
  assert.deepStrictEqual(titles, ["high", "low", "done"]);
});

test("streak counts consecutive completed occurrences", () => {
  let t = task({ date: "2026-10-20", repeat: "daily" });
  for (const d of ["2026-10-23", "2026-10-24", "2026-10-25"]) t = L.toggle(t, d);
  assert.strictEqual(L.streak(t, "2026-10-26"), 3); // today not done yet
  assert.strictEqual(L.streak(L.toggle(t, "2026-10-26"), "2026-10-26"), 4);

  let w = task({ date: "2026-10-19", repeat: "weekdays" }); // Monday
  for (const d of ["2026-10-22", "2026-10-23", "2026-10-26"]) w = L.toggle(w, d);
  assert.strictEqual(L.streak(w, "2026-10-26"), 3); // weekend is skipped
});

test("createTask stores category and validates time", () => {
  const t = task({ category: " Work ", time: "09:30" });
  assert.strictEqual(t.category, "Work");
  assert.strictEqual(t.time, "09:30");
  assert.throws(() => task({ time: "9am" }));
});

test("categories lists distinct non-empty names alphabetically", () => {
  const ts = [task({ category: "Work" }), task({ category: "Home" }), task({ category: "Work" }), task()];
  assert.deepStrictEqual(L.categories(ts), ["Home", "Work"]);
});

test("timed tasks sort earliest first within a priority", () => {
  const ts = [task({ title: "none" }), task({ title: "late", time: "18:00" }), task({ title: "early", time: "07:00" })];
  assert.deepStrictEqual(L.tasksForDay(ts, "2026-10-26").map((t) => t.title), ["early", "late", "none"]);
});

test("weekSummary covers Monday to Sunday", () => {
  assert.strictEqual(L.weekStart("2026-11-01"), "2026-10-26"); // Sunday -> Monday
  let gym = task({ repeat: "daily", category: "Health" });
  gym = L.toggle(L.toggle(gym, "2026-10-26"), "2026-10-27");
  const once = task({ date: "2026-10-28" });
  const s = L.weekSummary([gym, once], "2026-10-29");
  assert.strictEqual(s.start, "2026-10-26");
  assert.strictEqual(s.end, "2026-11-01");
  assert.strictEqual(s.total, 8);
  assert.strictEqual(s.done, 2);
  assert.deepStrictEqual(s.days[2], { key: "2026-10-28", total: 2, done: 0 });
  assert.deepStrictEqual(s.byCategory, { Health: { total: 7, done: 2 }, Uncategorized: { total: 1, done: 0 } });
});

test("dueReminders returns unfinished, unannounced tasks whose time has passed", () => {
  const a = task({ time: "08:00" });
  const b = task({ time: "12:00" });
  const c = L.toggle(task({ time: "07:00" }), "2026-10-26");
  const d = task();
  const notified = new Set();
  assert.deepStrictEqual(L.dueReminders([a, b, c, d], "2026-10-26", "09:00", notified), [a]);
  notified.add(`${a.id}@2026-10-26`);
  assert.deepStrictEqual(L.dueReminders([a, b, c, d], "2026-10-26", "12:00", notified), [b]);
});
