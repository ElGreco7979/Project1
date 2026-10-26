(function () {
  const L = window.TaskLogic;
  const STORAGE_KEY = "daily-tasks:v1";
  const NOTIFIED_KEY = "daily-tasks:notified";
  const REPEAT_LABELS = { daily: "Every day", weekdays: "Weekdays", weekly: "Weekly" };

  const $ = (id) => document.getElementById(id);
  const form = $("task-form");
  const template = $("task-template");

  const state = {
    tasks: load(),
    day: L.toKey(new Date()),
    view: "day",
    filter: "all",
    category: "",
    editingId: null,
    notified: loadNotified(),
  };

  function load() {
    try {
      const data = JSON.parse(localStorage.getItem(STORAGE_KEY));
      return Array.isArray(data) ? data : [];
    } catch {
      return [];
    }
  }

  function save() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.tasks));
  }

  // Only today's entries matter, so older ones are dropped on load.
  function loadNotified() {
    try {
      const suffix = `@${today()}`;
      return new Set((JSON.parse(localStorage.getItem(NOTIFIED_KEY)) || []).filter((k) => k.endsWith(suffix)));
    } catch {
      return new Set();
    }
  }

  function today() {
    return L.toKey(new Date());
  }

  function shortDate(key, opts = { month: "short", day: "numeric" }) {
    return L.fromKey(key).toLocaleDateString(undefined, opts);
  }

  function dayLabel(key) {
    const diff = Math.round((L.fromKey(key) - L.fromKey(today())) / 86400000);
    if (diff === 0) return "Today";
    if (diff === -1) return "Yesterday";
    if (diff === 1) return "Tomorrow";
    return shortDate(key, { weekday: "long" });
  }

  function formatTime(time) {
    const [h, m] = time.split(":").map(Number);
    return new Date(2000, 0, 1, h, m).toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" });
  }

  function renderTask(task, key) {
    const li = template.content.firstElementChild.cloneNode(true);
    const done = L.isDone(task, key);
    li.classList.add(`priority-${task.priority}`);
    li.classList.toggle("done", done);

    const checkbox = li.querySelector(".toggle");
    checkbox.checked = done;
    checkbox.addEventListener("change", () => update(task.id, (t) => L.toggle(t, key)));

    li.querySelector(".task-title").textContent = task.title;
    li.querySelector(".task-notes").textContent = task.notes;

    const meta = li.querySelector(".task-meta");
    const addTag = (text) => {
      const span = document.createElement("span");
      span.className = "tag";
      span.textContent = text;
      meta.appendChild(span);
    };
    if (task.time) addTag(`⏰ ${formatTime(task.time)}`);
    if (task.category) addTag(task.category);
    addTag(task.priority);
    if (task.repeat !== "none") {
      addTag(REPEAT_LABELS[task.repeat]);
      const s = L.streak(task, today());
      if (s > 1) addTag(`🔥 ${s} streak`);
    }
    if (key !== state.day) addTag(`from ${L.fromKey(key).toLocaleDateString()}`);

    li.querySelector(".edit").addEventListener("click", () => startEdit(task));
    li.querySelector(".delete").addEventListener("click", () => {
      if (confirm(`Delete "${task.title}"?`)) {
        state.tasks = state.tasks.filter((t) => t.id !== task.id);
        commit();
      }
    });
    return li;
  }

  function matchesFilter(task, key) {
    if (state.category && task.category !== state.category) return false;
    if (state.filter === "active") return !L.isDone(task, key);
    if (state.filter === "done") return L.isDone(task, key);
    return true;
  }

  function renderCategoryOptions() {
    const cats = L.categories(state.tasks);
    if (state.category && !cats.includes(state.category)) state.category = "";
    $("category-options").replaceChildren(...cats.map((c) => new Option(c)));
    $("category-filter").replaceChildren(new Option("All categories", ""), ...cats.map((c) => new Option(c)));
    $("category-filter").value = state.category;
    $("category-filter").hidden = cats.length === 0;
  }

  function renderDay() {
    const key = state.day;
    const isToday = key === today();

    $("day-label").textContent = dayLabel(key);
    $("day-sub").textContent = shortDate(key, { month: "long", day: "numeric", year: "numeric" });
    $("go-today").hidden = isToday;

    const dayTasks = L.tasksForDay(state.tasks, key);
    const doneCount = dayTasks.filter((t) => L.isDone(t, key)).length;
    $("progress-text").textContent = `${doneCount} / ${dayTasks.length} done`;
    $("progress-fill").style.width = dayTasks.length ? `${(doneCount / dayTasks.length) * 100}%` : "0";

    const list = $("task-list");
    list.replaceChildren(...dayTasks.filter((t) => matchesFilter(t, key)).map((t) => renderTask(t, key)));
    $("empty").hidden = list.children.length > 0;
    $("empty").textContent = dayTasks.length ? "Nothing matches this filter." : "No tasks for this day. Add one above!";

    // Overdue one-off tasks show up on today's view so nothing gets lost.
    const late = isToday && state.filter !== "done"
      ? L.overdue(state.tasks, key).filter((t) => matchesFilter(t, t.date))
      : [];
    $("overdue-section").hidden = late.length === 0;
    $("overdue-list").replaceChildren(...late.map((t) => renderTask(t, t.date)));
  }

  function meterRow(name, done, total) {
    const li = document.createElement("li");
    const label = document.createElement("span");
    label.className = "name";
    label.textContent = name;
    const count = document.createElement("span");
    count.className = "count";
    count.textContent = `${done} / ${total}`;
    const meter = document.createElement("div");
    meter.className = "meter";
    const fill = document.createElement("div");
    fill.style.width = total ? `${(done / total) * 100}%` : "0";
    meter.appendChild(fill);
    li.append(label, count, meter);
    return li;
  }

  function renderWeek() {
    const s = L.weekSummary(state.tasks, state.day);
    const thisWeek = s.start === L.weekStart(today());

    $("day-label").textContent = thisWeek ? "This week" : `Week of ${shortDate(s.start)}`;
    $("day-sub").textContent = `${shortDate(s.start)} – ${shortDate(s.end, { month: "short", day: "numeric", year: "numeric" })}`;
    $("go-today").hidden = thisWeek;
    $("go-today").textContent = "Back to this week";

    $("week-rate").textContent = s.total ? `${Math.round((s.done / s.total) * 100)}%` : "–";
    $("week-count").textContent = `${s.done} of ${s.total} tasks done`;

    const chart = $("week-chart");
    chart.setAttribute(
      "aria-label",
      s.days.map((d) => `${shortDate(d.key, { weekday: "long" })}: ${d.done} of ${d.total} done`).join(", ")
    );
    chart.replaceChildren(
      ...s.days.map((d) => {
        const col = document.createElement("button");
        col.className = "bar-col";
        col.classList.toggle("selected", d.key === today());
        col.dataset.tip = `${shortDate(d.key, { weekday: "short", month: "short", day: "numeric" })}: ${d.done} of ${d.total} done`;
        col.setAttribute("aria-label", col.dataset.tip);
        col.addEventListener("click", () => {
          state.day = d.key;
          setView("day");
        });

        const value = document.createElement("span");
        value.className = "bar-value";
        value.textContent = d.total ? `${d.done}/${d.total}` : "";
        const track = document.createElement("div");
        track.className = "bar-track";
        const fill = document.createElement("div");
        fill.className = "bar-fill";
        fill.style.height = d.total ? `${(d.done / d.total) * 100}%` : "0";
        track.appendChild(fill);
        const label = document.createElement("span");
        label.className = "bar-label";
        label.textContent = shortDate(d.key, { weekday: "narrow" });
        col.append(value, track, label);
        return col;
      })
    );

    const cats = Object.entries(s.byCategory).sort((a, b) => b[1].total - a[1].total);
    $("category-breakdown").replaceChildren(...cats.map(([name, c]) => meterRow(name, c.done, c.total)));
    $("category-empty").hidden = cats.length > 0;

    const habits = state.tasks
      .filter((t) => t.repeat !== "none")
      .map((t) => ({ title: t.title, streak: L.streak(t, today()) }))
      .sort((a, b) => b.streak - a.streak);
    $("streak-list").replaceChildren(
      ...habits.map((h) => {
        const li = document.createElement("li");
        const name = document.createElement("span");
        name.className = "name";
        name.textContent = h.title;
        const count = document.createElement("span");
        count.className = "count";
        count.textContent = h.streak ? `🔥 ${h.streak} ${h.streak === 1 ? "day" : "days"}` : "No streak yet";
        li.append(name, count);
        return li;
      })
    );
    $("streak-empty").hidden = habits.length > 0;
  }

  function render() {
    renderCategoryOptions();
    $("day-view").hidden = state.view !== "day";
    $("week-view").hidden = state.view !== "week";
    $("go-today").textContent = "Back to today";
    if (state.view === "day") renderDay();
    else renderWeek();
    $("notify-btn").hidden = !("Notification" in window) || Notification.permission !== "default"
      || !state.tasks.some((t) => t.time);
  }

  function commit() {
    save();
    render();
  }

  function update(id, fn) {
    state.tasks = state.tasks.map((t) => (t.id === id ? fn(t) : t));
    commit();
  }

  function resetForm() {
    state.editingId = null;
    form.reset();
    $("date").value = state.day;
    $("submit-btn").textContent = "Add task";
    $("cancel-edit").hidden = true;
  }

  function startEdit(task) {
    state.editingId = task.id;
    $("title").value = task.title;
    $("notes").value = task.notes;
    $("priority").value = task.priority;
    $("repeat").value = task.repeat;
    $("date").value = task.date;
    $("category").value = task.category || "";
    $("time").value = task.time || "";
    $("submit-btn").textContent = "Save";
    $("cancel-edit").hidden = false;
    $("title").focus();
  }

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const fields = {
      title: $("title").value,
      notes: $("notes").value,
      priority: $("priority").value,
      repeat: $("repeat").value,
      date: $("date").value || state.day,
      category: $("category").value,
      time: $("time").value,
    };
    if (fields.time) requestNotifications();
    if (state.editingId) {
      const id = state.editingId;
      update(id, (t) => ({ ...L.createTask(fields, id), completedOn: t.completedOn, createdAt: t.createdAt }));
    } else {
      state.tasks.push(L.createTask(fields));
      commit();
    }
    resetForm();
  });

  $("cancel-edit").addEventListener("click", resetForm);

  function goTo(key) {
    state.day = key;
    if (!state.editingId) $("date").value = key;
    render();
  }

  function setView(view) {
    state.view = view;
    document.querySelectorAll(".view-tab").forEach((b) => b.classList.toggle("active", b.dataset.view === view));
    goTo(state.day);
  }

  const step = () => (state.view === "week" ? 7 : 1);
  $("prev").addEventListener("click", () => goTo(L.addDays(state.day, -step())));
  $("next").addEventListener("click", () => goTo(L.addDays(state.day, step())));
  $("go-today").addEventListener("click", () => goTo(today()));

  document.querySelectorAll(".view-tab").forEach((btn) => btn.addEventListener("click", () => setView(btn.dataset.view)));

  document.querySelectorAll(".filter").forEach((btn) =>
    btn.addEventListener("click", () => {
      state.filter = btn.dataset.filter;
      document.querySelectorAll(".filter").forEach((b) => b.classList.toggle("active", b === btn));
      render();
    })
  );

  $("category-filter").addEventListener("change", (e) => {
    state.category = e.target.value;
    render();
  });

  // Reminders: checked every 20s while the page is open.
  function requestNotifications() {
    if ("Notification" in window && Notification.permission === "default") {
      Notification.requestPermission().then(render);
    }
  }

  function toast(message) {
    const el = document.createElement("div");
    el.className = "toast";
    el.textContent = message;
    el.addEventListener("click", () => el.remove());
    $("toasts").appendChild(el);
    setTimeout(() => el.remove(), 15000);
  }

  function checkReminders() {
    const now = new Date();
    const key = L.toKey(now);
    const time = `${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;
    const due = L.dueReminders(state.tasks, key, time, state.notified);
    if (!due.length) return;
    for (const task of due) {
      const message = `⏰ ${task.title}`;
      toast(message);
      if ("Notification" in window && Notification.permission === "granted") {
        try {
          new Notification("Daily Tasks", { body: task.title, tag: `${task.id}@${key}` });
        } catch {
          // Some browsers only allow notifications from a service worker; the toast still shows.
        }
      }
      state.notified.add(`${task.id}@${key}`);
    }
    localStorage.setItem(NOTIFIED_KEY, JSON.stringify([...state.notified]));
  }

  $("notify-btn").addEventListener("click", requestNotifications);

  $("export-btn").addEventListener("click", () => {
    const blob = new Blob([JSON.stringify(state.tasks, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `tasks-${today()}.json`;
    a.click();
    URL.revokeObjectURL(a.href);
  });

  $("import-input").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      if (!Array.isArray(data)) throw new Error("Expected a list of tasks");
      if (confirm(`Replace your current tasks with ${data.length} imported task(s)?`)) {
        state.tasks = data;
        commit();
      }
    } catch (err) {
      alert(`Could not import: ${err.message}`);
    }
    e.target.value = "";
  });

  resetForm();
  render();
  checkReminders();
  setInterval(checkReminders, 20000);
})();
