(function () {
  const L = window.TaskLogic;
  const STORAGE_KEY = "daily-tasks:v1";
  const REPEAT_LABELS = { daily: "Every day", weekdays: "Weekdays", weekly: "Weekly" };

  const $ = (id) => document.getElementById(id);
  const form = $("task-form");
  const template = $("task-template");

  const state = {
    tasks: load(),
    day: L.toKey(new Date()),
    filter: "all",
    editingId: null,
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

  function today() {
    return L.toKey(new Date());
  }

  function dayLabel(key) {
    const diff = Math.round((L.fromKey(key) - L.fromKey(today())) / 86400000);
    if (diff === 0) return "Today";
    if (diff === -1) return "Yesterday";
    if (diff === 1) return "Tomorrow";
    return L.fromKey(key).toLocaleDateString(undefined, { weekday: "long" });
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
    if (state.filter === "active") return !L.isDone(task, key);
    if (state.filter === "done") return L.isDone(task, key);
    return true;
  }

  function render() {
    const key = state.day;
    const isToday = key === today();

    $("day-label").textContent = dayLabel(key);
    $("day-sub").textContent = L.fromKey(key).toLocaleDateString(undefined, {
      month: "long", day: "numeric", year: "numeric",
    });
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
    const late = isToday && state.filter !== "done" ? L.overdue(state.tasks, key) : [];
    $("overdue-section").hidden = late.length === 0;
    $("overdue-list").replaceChildren(...late.map((t) => renderTask(t, t.date)));
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
    };
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

  $("prev-day").addEventListener("click", () => goTo(L.addDays(state.day, -1)));
  $("next-day").addEventListener("click", () => goTo(L.addDays(state.day, 1)));
  $("go-today").addEventListener("click", () => goTo(today()));

  document.querySelectorAll(".filter").forEach((btn) =>
    btn.addEventListener("click", () => {
      state.filter = btn.dataset.filter;
      document.querySelectorAll(".filter").forEach((b) => b.classList.toggle("active", b === btn));
      render();
    })
  );

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
})();
