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
