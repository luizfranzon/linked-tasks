let workbookData = [];
let tasks = [];
let currentTask = null;
let sortConfig = { key: null, direction: "asc" };
let techEditor = null;
let planningEditor = null;
let timerInterval = null;
let manualOrder = false;
let lastRendered = [];

const COLOR_OPTIONS = [
  { value: "#a855f7", label: "Roxa – próxima versão" },
  { value: "#ef4444", label: "Vermelha – semana atual" },
  { value: "#f59e0b", label: "Amarela – finalizar nesta semana" },
  { value: "#22c55e", label: "Verde – neutra" },
  { value: "", label: "---" },
  { value: "#64748b", label: "Cinza – sem prioridade" }
];

const COLOR_VALUE_MAP = {
  roxo: "#a855f7",
  roxa: "#a855f7",
  purple: "#a855f7",
  vermelho: "#ef4444",
  vermelha: "#ef4444",
  red: "#ef4444",
  amarelo: "#f59e0b",
  amarela: "#f59e0b",
  yellow: "#f59e0b",
  verde: "#22c55e",
  green: "#22c55e",
  cinza: "#64748b",
  cinzento: "#64748b",
  gray: "#64748b",
  grey: "#64748b"
};

const fileInput = document.getElementById("fileInput");
const devSelect = document.getElementById("devSelect");
const importBtn = document.getElementById("importBtn");
const resetFiltersBtn = document.getElementById("resetFilters");
const exportBackupBtn = document.getElementById("exportBackup");
const importBackupBtn = document.getElementById("importBackup");
const backupFileInput = document.getElementById("backupFile");
const resetSiteBtn = document.getElementById("resetSite");
const manualTitleInput = document.getElementById("manualTitle");
const manualPriorityInput = document.getElementById("manualPriority");
const addManualBtn = document.getElementById("addManualBtn");
const resetModal = document.getElementById("resetModal");
const resetConfirmInput = document.getElementById("resetConfirmInput");
const deleteModal = document.getElementById("deleteModal");
const doneCountEl = document.getElementById("doneCount");
const totalCountEl = document.getElementById("totalCount");
const dayTimerEl = document.getElementById("dayTimer");
const taskList = document.getElementById("taskList");
const modal = document.getElementById("modal");
const techNotesInput = document.getElementById("techNotesInput");
const planningNotesInput = document.getElementById("planningNotesInput");
const todoListEl = document.getElementById("todoList");
const todoNewTextInput = document.getElementById("todoNewText");
const todoAddRootBtn = document.getElementById("todoAddRoot");

let deleteTargetId = null;

fileInput.addEventListener("change", handleFile);
importBtn.addEventListener("click", importTasks);
resetFiltersBtn.addEventListener("click", resetFilters);
exportBackupBtn.addEventListener("click", exportBackup);
importBackupBtn.addEventListener("click", () => backupFileInput.click());
backupFileInput.addEventListener("change", handleBackupFile);
resetSiteBtn.addEventListener("click", openResetModal);
addManualBtn?.addEventListener("click", addManualTask);
todoAddRootBtn?.addEventListener("click", () => addTodoItem(null));
todoNewTextInput?.addEventListener("keydown", evt => {
  if (evt.key === "Enter") {
    evt.preventDefault();
    addTodoItem(null);
  }
});

initEditor();
loadFromStorage();
startDayTimer();

/* ===== Importação ===== */

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(evt.target.result, { type: "binary" });
    const sheetName = getLatestSheet(wb.SheetNames);
    const sheet = wb.Sheets[sheetName];
    workbookData = XLSX.utils.sheet_to_json(sheet);

    populateDevSelect(workbookData);
  };
  reader.readAsBinaryString(file);
}

function getLatestSheet(names) {
  return names
    .filter(n => /^\d{2}-\d{2}$/.test(n))
    .sort((a, b) => {
      const [da, ma] = a.split("-").map(Number);
      const [db, mb] = b.split("-").map(Number);
      return ma !== mb ? mb - ma : db - da;
    })[0];
}

function populateDevSelect(rows) {
  const devs = [...new Set(
    rows.map(r => r["Dev Atual"]).filter(Boolean).map(v => v.trim())
  )];

  devSelect.innerHTML = `<option value="">Selecionar Dev</option>`;
  devs.forEach(dev => {
    const opt = document.createElement("option");
    opt.value = dev;
    opt.textContent = dev;
    devSelect.appendChild(opt);
  });

  devSelect.disabled = false;
  importBtn.disabled = false;
}

function importTasks() {
  const selectedDev = devSelect.value;
  if (!selectedDev) return;

  tasks = workbookData
    .filter(r => r["Dev Atual"]?.trim() === selectedDev)
    .map(r => ({
      id: r["Tarefa"],
      title: r["Título"],
      color: normalizeColor(r["Cor"]),
      priority: r["Prioridade"]?.toString().trim() || "",
      completed: false,
      notes: "",
      techAnalysis: "",
      planning: "",
      todos: []
    }));

  manualOrder = false;

  saveToStorage();
  render();
}

function addManualTask() {
  const title = manualTitleInput?.value.trim();
  const priority = manualPriorityInput?.value.trim();

  const newTask = {
    id: getNextId(),
    title: title || "Tarefa manual",
    color: "",
    priority: priority || "",
    completed: false,
    notes: "",
    techAnalysis: "",
    planning: "",
    todos: []
  };

  tasks = [...tasks, newTask];
  manualOrder = true;
  sortConfig = { key: null, direction: "asc" };
  saveToStorage();
  render();

  if (manualTitleInput) manualTitleInput.value = "";
  if (manualPriorityInput) manualPriorityInput.value = "";
}

/* ===== Render ===== */

function render() {
  taskList.innerHTML = "";

  taskList.appendChild(buildHeader());

  const sorted = sortTasks(tasks);
  lastRendered = sorted;

  updateFooter(sorted);

  sorted.forEach(task => {
    const div = document.createElement("div");
    div.className = "task" + (task.completed ? " completed" : "");
    div.draggable = true;
    div.dataset.taskId = task.id;

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.className = "task-checkbox";
    checkbox.setAttribute("aria-label", "Marcar tarefa como concluída");
    checkbox.checked = task.completed;
    checkbox.onchange = () => {
      task.completed = checkbox.checked;
      saveToStorage();
      render();
      if (task.completed) confetti();
    };

    const colorSelect = buildColorSelect(task);

    const idBtn = document.createElement("button");
    idBtn.className = "cell-btn";
    idBtn.textContent = `#${task.id}`;
    idBtn.onclick = () => editId(task);

    const priorityInput = document.createElement("input");
    priorityInput.className = "inline-input priority-input";
    priorityInput.value = task.priority ?? "";
    priorityInput.placeholder = "Prioridade";
    priorityInput.onblur = () => {
      task.priority = priorityInput.value.trim();
      saveToStorage();
      render();
    };
    priorityInput.onkeydown = evt => {
      if (evt.key === "Enter") {
        evt.preventDefault();
        priorityInput.blur();
      }
    };

    const titleInput = document.createElement("textarea");
    titleInput.className = "inline-input inline-textarea";
    titleInput.value = task.title ?? "";
    titleInput.placeholder = "Título";
    titleInput.rows = 2;
    titleInput.oninput = () => autoSizeTextarea(titleInput);
    titleInput.onblur = () => {
      task.title = titleInput.value.trim();
      saveToStorage();
      render();
    };

    const notesBtn = document.createElement("button");
    notesBtn.className = "icon-btn";
    notesBtn.setAttribute("aria-label", "Abrir anotações");
    notesBtn.title = "Anotações";
    notesBtn.innerHTML = `
      <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 256 256" aria-hidden="true" focusable="false">
        <path fill="currentColor" d="M88,96a8,8,0,0,1,8-8h64a8,8,0,0,1,0,16H96A8,8,0,0,1,88,96Zm8,40h64a8,8,0,0,0,0-16H96a8,8,0,0,0,0,16Zm32,16H96a8,8,0,0,0,0,16h32a8,8,0,0,0,0-16ZM224,48V156.69A15.86,15.86,0,0,1,219.31,168L168,219.31A15.86,15.86,0,0,1,156.69,224H48a16,16,0,0,1-16-16V48A16,16,0,0,1,48,32H208A16,16,0,0,1,224,48ZM48,208H152V160a8,8,0,0,1,8-8h48V48H48Zm120-40v28.7L196.69,168Z"></path>
      </svg>
    `;
    notesBtn.classList.toggle(
      "note-active",
      Boolean(task.notes?.trim() || task.techAnalysis?.trim() || task.planning?.trim())
    );
    notesBtn.onclick = () => openModal(task);

    const deleteBtn = document.createElement("button");
    deleteBtn.className = "icon-btn ghost-danger";
    deleteBtn.setAttribute("aria-label", "Excluir tarefa");
    deleteBtn.title = "Excluir";
    deleteBtn.innerHTML = `
      <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 256 256" aria-hidden="true" focusable="false">
        <path fill="currentColor" d="M216,48H176V40a24,24,0,0,0-24-24H104A24,24,0,0,0,80,40v8H40a8,8,0,0,0,0,16h8V208a16,16,0,0,0,16,16H192a16,16,0,0,0,16-16V64h8a8,8,0,0,0,0-16ZM96,40a8,8,0,0,1,8-8h48a8,8,0,0,1,8,8v8H96Zm96,168H64V64H192ZM112,104v64a8,8,0,0,1-16,0V104a8,8,0,0,1,16,0Zm48,0v64a8,8,0,0,1-16,0V104a8,8,0,0,1,16,0Z"></path>
      </svg>
    `;
    deleteBtn.onclick = () => openDeleteModal(task.id);

    const actionCell = document.createElement("div");
    actionCell.className = "task-actions";
    actionCell.appendChild(notesBtn);
    actionCell.appendChild(deleteBtn);

    div.appendChild(checkbox);
    div.appendChild(colorSelect);
    div.appendChild(idBtn);
    div.appendChild(priorityInput);
    div.appendChild(titleInput);
    div.appendChild(actionCell);

    div.addEventListener("dragstart", e => onDragStart(e, task.id));
    div.addEventListener("dragover", onDragOver);
    div.addEventListener("drop", e => onDrop(e, task.id));
    div.addEventListener("dragend", onDragEnd);

    taskList.appendChild(div);
  });
}

/* ===== Modal ===== */

function openModal(task) {
  currentTask = task;
  ensureTodos(currentTask);
  renderTodoList();
  if (todoNewTextInput) todoNewTextInput.value = "";

  modal.style.display = "flex";

  const legacyNotes = String(task.notes ?? "");
  const tech = String(task.techAnalysis ?? legacyNotes);
  const planning = String(task.planning ?? "");

  // EasyMDE/CodeMirror only paints correctly after the modal is visible.
  const syncEditors = () => {
    if (techEditor) {
      techEditor.value(tech);
    } else if (techNotesInput) {
      techNotesInput.value = tech;
    }

    if (planningEditor) {
      planningEditor.value(planning);
    } else if (planningNotesInput) {
      planningNotesInput.value = planning;
    }

    techEditor?.codemirror.refresh();
    planningEditor?.codemirror.refresh();
  };

  requestAnimationFrame(() => {
    syncEditors();
    requestAnimationFrame(() => {
      techEditor?.codemirror.refresh();
      planningEditor?.codemirror.refresh();
    });
  });
}

function closeModal() {
  modal.style.display = "none";
}

function saveNotes() {
  if (currentTask) {
    const tech = techEditor ? techEditor.value() : (techNotesInput?.value ?? "");
    const planning = planningEditor ? planningEditor.value() : (planningNotesInput?.value ?? "");

    currentTask.techAnalysis = tech;
    currentTask.planning = planning;
    // mantém compatibilidade com recursos existentes (highlight/sort)
    currentTask.notes = [tech, planning].map(v => String(v ?? "").trim()).filter(Boolean).join("\n\n");
    ensureTodos(currentTask);
    saveToStorage();
    render();
  }
  closeModal();
}

/* ===== Storage ===== */

function saveToStorage() {
  localStorage.setItem("tasks", JSON.stringify(tasks));
  localStorage.setItem("manualOrder", manualOrder ? "1" : "0");
}

function loadFromStorage() {
  const saved = localStorage.getItem("tasks");
  if (!saved) return render();

  try {
    const parsed = JSON.parse(saved);
    tasks = Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    console.warn("Falha ao ler tasks do storage, limpando.", err);
    tasks = [];
    localStorage.removeItem("tasks");
  }

  manualOrder = localStorage.getItem("manualOrder") === "1";

  render();
}

/* ===== Sorting ===== */

function buildHeader() {
  const header = document.createElement("div");
  header.className = "task task-header";
  header.draggable = false;
  header.innerHTML = `
    <button type="button" class="header-btn" data-key="completed">Feita</button>
    <button type="button" class="header-btn" data-key="color">Cor</button>
    <button type="button" class="header-btn" data-key="id">ID</button>
    <button type="button" class="header-btn" data-key="priority">Prioridade</button>
    <button type="button" class="header-btn" data-key="title">Título</button>
    <button type="button" class="header-btn" data-key="notes">Notas/Ações</button>
  `;

  header.querySelectorAll(".header-btn").forEach(btn => {
    btn.onclick = () => setSort(btn.dataset.key);
    if (sortConfig.key === btn.dataset.key) {
      btn.setAttribute("data-direction", sortConfig.direction);
    } else {
      btn.removeAttribute("data-direction");
    }
  });

  return header;
}

function setSort(key) {
  if (sortConfig.key === key) {
    sortConfig.direction = sortConfig.direction === "asc" ? "desc" : "asc";
  } else {
    sortConfig = { key, direction: "asc" };
  }
  render();
}

function sortTasks(list) {
  if (!sortConfig.key) {
    if (manualOrder) return [...list];

    return [...list].sort((a, b) => {
      const aHasColor = Boolean(a.color);
      const bHasColor = Boolean(b.color);
      if (aHasColor !== bHasColor) return aHasColor ? -1 : 1;

      const aPrio = normalizePriority(a.priority);
      const bPrio = normalizePriority(b.priority);
      if (aPrio !== bPrio) return aPrio - bPrio;

      return (a.id ?? 0) - (b.id ?? 0);
    });
  }

  const dir = sortConfig.direction === "asc" ? 1 : -1;

  return [...list].sort((a, b) => {
    const va = a[sortConfig.key];
    const vb = b[sortConfig.key];

    if (typeof va === "number" && typeof vb === "number") {
      return (va - vb) * dir;
    }

    if (typeof va === "boolean" && typeof vb === "boolean") {
      return (Number(va) - Number(vb)) * dir;
    }

    return String(va ?? "").localeCompare(String(vb ?? ""), "pt", {
      sensitivity: "base"
    }) * dir;
  });
}

function buildColorSelect(task) {
  const select = document.createElement("select");
  select.className = "color-select";

  const applyBg = value => {
    select.style.backgroundColor = value;
    select.style.color = getTextColor(value);
  };

  COLOR_OPTIONS.forEach(opt => {
    const option = document.createElement("option");
    option.value = opt.value;
    option.textContent = opt.label;
    option.style.backgroundColor = opt.value || "#020617";
    option.style.color = opt.value ? getTextColor(opt.value) : "var(--muted)";
    select.appendChild(option);
  });

  const existing = COLOR_OPTIONS.find(o => o.value === task.color);
  select.value = existing ? existing.value : "";
  applyBg(select.value || "transparent");

  select.onchange = () => {
    task.color = select.value;
    applyBg(task.color || "transparent");
    saveToStorage();
    render();
  };

  return select;
}

function getTextColor(hex) {
  const h = hex.replace("#", "");
  const [r, g, b] = h.length === 3
    ? h.split("").map(c => parseInt(c + c, 16))
    : [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
  const luma = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return luma > 140 ? "#0f172a" : "#f8fafc";
}

function normalizeColor(raw) {
  if (!raw) return "";
  const str = String(raw).trim();
  if (!str) return "";

  if (str.startsWith("#")) {
    const hex = str.length === 4 || str.length === 7 ? str : "";
    return hex;
  }

  const key = str.toLowerCase();
  return COLOR_VALUE_MAP[key] ?? "";
}

function editId(task) {
  const next = prompt("Editar ID da tarefa", task.id ?? "");
  if (next === null) return;
  const parsed = Number(next);
  if (!Number.isNaN(parsed)) {
    task.id = parsed;
    saveToStorage();
    render();
  }
}

function resetFilters() {
  sortConfig = { key: null, direction: "asc" };
  manualOrder = false;
  render();
}

function openDeleteModal(taskId) {
  deleteTargetId = taskId;
  deleteModal.style.display = "flex";
}

function closeDeleteModal() {
  deleteModal.style.display = "none";
  deleteTargetId = null;
}

function confirmDelete() {
  if (deleteTargetId === null) return closeDeleteModal();
  tasks = tasks.filter(t => String(t.id) !== String(deleteTargetId));
  if (!sortConfig.key) {
    manualOrder = true;
  }
  saveToStorage();
  render();
  closeDeleteModal();
}

function getNextId() {
  const numericIds = tasks
    .map(t => Number(t.id))
    .filter(n => !Number.isNaN(n));
  if (!numericIds.length) return 1;
  return Math.max(...numericIds) + 1;
}

function normalizePriority(value) {
  const num = Number(value);
  if (!Number.isNaN(num)) return num;
  return Number.MAX_SAFE_INTEGER;
}

function autoSizeTextarea(el) {
  el.style.height = "auto";
  el.style.height = `${el.scrollHeight}px`;
}

/* ===== Nested To-do List (per task) ===== */

let todoFocusId = null;

function ensureTodos(task) {
  if (!task) return;
  if (!Array.isArray(task.todos)) task.todos = [];
}

function uid() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function findTodoRef(list, todoId) {
  for (let i = 0; i < list.length; i++) {
    const item = list[i];
    if (String(item.id) === String(todoId)) return { list, index: i, item };
    if (Array.isArray(item.children) && item.children.length) {
      const hit = findTodoRef(item.children, todoId);
      if (hit) return hit;
    }
  }
  return null;
}

function addTodoItem(parentTodoId) {
  if (!currentTask) return;
  ensureTodos(currentTask);

  const textFromInput = todoNewTextInput?.value.trim() ?? "";
  const text = parentTodoId ? "Novo subitem" : (textFromInput || "Novo item");

  const item = {
    id: uid(),
    text,
    completed: false,
    children: []
  };

  if (!parentTodoId) {
    currentTask.todos.push(item);
    if (todoNewTextInput) todoNewTextInput.value = "";
  } else {
    const parentRef = findTodoRef(currentTask.todos, parentTodoId);
    if (!parentRef) return;
    if (!Array.isArray(parentRef.item.children)) parentRef.item.children = [];
    parentRef.item.children.push(item);
  }

  todoFocusId = item.id;
  saveToStorage();
  renderTodoList();
}

function updateTodoText(todoId, text) {
  if (!currentTask) return;
  ensureTodos(currentTask);
  const ref = findTodoRef(currentTask.todos, todoId);
  if (!ref) return;
  ref.item.text = text;
  saveToStorage();
}

function toggleTodo(todoId, checked) {
  if (!currentTask) return;
  ensureTodos(currentTask);
  const ref = findTodoRef(currentTask.todos, todoId);
  if (!ref) return;
  ref.item.completed = Boolean(checked);
  saveToStorage();
  renderTodoList();
}

function deleteTodo(todoId) {
  if (!currentTask) return;
  ensureTodos(currentTask);
  const ref = findTodoRef(currentTask.todos, todoId);
  if (!ref) return;
  ref.list.splice(ref.index, 1);
  saveToStorage();
  renderTodoList();
}

function renderTodoList() {
  if (!todoListEl) return;
  if (!currentTask) return;
  ensureTodos(currentTask);

  todoListEl.innerHTML = "";
  const frag = document.createDocumentFragment();

  const renderItems = (items, depth) => {
    items.forEach(item => {
      const row = document.createElement("div");
      row.className = "todo-item" + (item.completed ? " is-done" : "");
      row.style.setProperty("--depth", String(depth));

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.className = "task-checkbox";
      checkbox.checked = Boolean(item.completed);
      checkbox.setAttribute("aria-label", "Marcar item como concluído");
      checkbox.onchange = () => toggleTodo(item.id, checkbox.checked);

      const input = document.createElement("input");
      input.className = "inline-input todo-text";
      input.value = item.text ?? "";
      input.placeholder = "Item do to-do";
      input.dataset.todoId = item.id;
      input.onblur = () => updateTodoText(item.id, input.value.trim());
      input.onkeydown = evt => {
        if (evt.key === "Enter") {
          evt.preventDefault();
          input.blur();
        }
      };

      const addChildBtn = document.createElement("button");
      addChildBtn.className = "icon-btn";
      addChildBtn.type = "button";
      addChildBtn.title = "Adicionar subitem";
      addChildBtn.setAttribute("aria-label", "Adicionar subitem");
      addChildBtn.textContent = "+";
      addChildBtn.onclick = () => addTodoItem(item.id);

      const deleteBtn = document.createElement("button");
      deleteBtn.className = "icon-btn ghost-danger";
      deleteBtn.type = "button";
      deleteBtn.title = "Excluir item";
      deleteBtn.setAttribute("aria-label", "Excluir item");
      deleteBtn.innerHTML = `
        <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 256 256" aria-hidden="true" focusable="false">
          <path fill="currentColor" d="M216,48H176V40a24,24,0,0,0-24-24H104A24,24,0,0,0,80,40v8H40a8,8,0,0,0,0,16h8V208a16,16,0,0,0,16,16H192a16,16,0,0,0,16-16V64h8a8,8,0,0,0,0-16ZM96,40a8,8,0,0,1,8-8h48a8,8,0,0,1,8,8v8H96Zm96,168H64V64H192ZM112,104v64a8,8,0,0,1-16,0V104a8,8,0,0,1,16,0Zm48,0v64a8,8,0,0,1-16,0V104a8,8,0,0,1,16,0Z"></path>
        </svg>
      `;
      deleteBtn.onclick = () => deleteTodo(item.id);

      row.appendChild(checkbox);
      row.appendChild(input);
      row.appendChild(addChildBtn);
      row.appendChild(deleteBtn);
      frag.appendChild(row);

      if (Array.isArray(item.children) && item.children.length) {
        renderItems(item.children, depth + 1);
      }
    });
  };

  renderItems(currentTask.todos, 0);
  todoListEl.appendChild(frag);

  if (todoFocusId) {
    const toFocus = todoListEl.querySelector(`[data-todo-id="${CSS.escape(String(todoFocusId))}"]`);
    if (toFocus) {
      toFocus.focus();
      toFocus.select?.();
    }
    todoFocusId = null;
  }
}

/* ===== Drag & Drop ===== */

function onDragStart(e, taskId) {
  e.dataTransfer.effectAllowed = "move";
  e.dataTransfer.setData("text/plain", String(taskId));
  e.currentTarget.classList.add("dragging");
}

function onDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = "move";
}

function onDrop(e, targetId) {
  e.preventDefault();
  const sourceId = e.dataTransfer.getData("text/plain");
  if (!sourceId || String(sourceId) === String(targetId)) return;
  if (!lastRendered || !lastRendered.length) return;

  const order = [...lastRendered];
  const from = order.findIndex(t => String(t.id) === sourceId);
  const to = order.findIndex(t => String(t.id) === String(targetId));
  if (from === -1 || to === -1) return;

  const [moved] = order.splice(from, 1);
  order.splice(to, 0, moved);

  tasks = order;
  manualOrder = true;
  sortConfig = { key: null, direction: "asc" };
  saveToStorage();
  render();
}

function onDragEnd(e) {
  e.currentTarget.classList.remove("dragging");
}

/* ===== Backup ===== */

function exportBackup() {
  const payload = { tasks, manualOrder };
  const data = JSON.stringify(payload, null, 2);
  const blob = new Blob([data], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const stamp = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = `tasks-backup-${stamp}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function handleBackupFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = evt => {
    try {
      const parsed = JSON.parse(evt.target.result);
      if (Array.isArray(parsed)) {
        tasks = parsed;
        manualOrder = false;
      } else if (parsed && Array.isArray(parsed.tasks)) {
        tasks = parsed.tasks;
        manualOrder = Boolean(parsed.manualOrder);
      } else {
        throw new Error("Formato inválido");
      }
      saveToStorage();
      render();
    } catch (err) {
      console.error("Falha ao importar backup", err);
      alert("Não foi possível importar o backup. Verifique o arquivo.");
    } finally {
      backupFileInput.value = "";
    }
  };
  reader.readAsText(file);
}

/* ===== Reset Site ===== */

function openResetModal() {
  resetConfirmInput.value = "";
  resetModal.style.display = "flex";
  resetConfirmInput.focus();
}

function closeResetModal() {
  resetModal.style.display = "none";
}

function confirmReset() {
  const value = resetConfirmInput.value.trim().toLowerCase();
  if (value !== "apagar") {
    alert("Digite 'apagar' para confirmar a deleção.");
    return;
  }

  tasks = [];
  workbookData = [];
  manualOrder = false;
  localStorage.removeItem("tasks");
  localStorage.removeItem("manualOrder");
  saveToStorage();
  render();
  updateFooter(tasks);
  closeResetModal();
}

function updateFooter(list) {
  if (!doneCountEl || !totalCountEl) return;
  const total = list.length;
  const done = list.filter(t => t.completed).length;
  doneCountEl.textContent = done;
  totalCountEl.textContent = total;
}

function startDayTimer() {
  if (!dayTimerEl) return;
  if (timerInterval) clearInterval(timerInterval);

  const targetMinutes = 17 * 60 + 30; // 17:30

  const update = () => {
    const now = new Date();
    const currentMinutes = now.getHours() * 60 + now.getMinutes();
    let diff = targetMinutes - currentMinutes;

    if (diff < 0) {
      // already past 17:30, show 00:00:00
      dayTimerEl.textContent = "00:00:00";
      return;
    }

    const hours = Math.floor(diff / 60);
    const minutes = diff % 60;
    const seconds = 59 - now.getSeconds();

    const pad = n => String(n).padStart(2, "0");
    dayTimerEl.textContent = `${pad(hours)}:${pad(minutes)}:${pad(seconds)}`;
  };

  update();
  timerInterval = setInterval(update, 1000);
}

function initEditor() {
  if (typeof EasyMDE === "undefined") {
    console.warn("EasyMDE não carregou; usando textarea simples.");
    return;
  }

  techEditor = new EasyMDE({
    element: techNotesInput,
    autofocus: false,
    spellChecker: false,
    status: ["lines", "words", "cursor"],
    autosave: { enabled: false },
    toolbar: [
      "bold",
      "italic",
      "heading",
      "quote",
      "unordered-list",
      "ordered-list",
      "link",
      "preview",
      "guide"
    ],
    shortcuts: {
      drawBold: "Ctrl-B",
      drawItalic: "Ctrl-I",
      drawLink: "Ctrl-K",
      togglePreview: "Ctrl-P"
    }
  });

  planningEditor = new EasyMDE({
    element: planningNotesInput,
    autofocus: false,
    spellChecker: false,
    status: ["lines", "words", "cursor"],
    autosave: { enabled: false },
    toolbar: [
      "bold",
      "italic",
      "heading",
      "quote",
      "unordered-list",
      "ordered-list",
      "link",
      "preview",
      "guide"
    ],
    shortcuts: {
      drawBold: "Ctrl-B",
      drawItalic: "Ctrl-I",
      drawLink: "Ctrl-K",
      togglePreview: "Ctrl-P"
    }
  });
}
