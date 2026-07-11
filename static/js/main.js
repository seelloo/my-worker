/**
 * Task Engine Server v3.0 — 数据员工 Dashboard
 * Modular JS: TaskRunner, QueueManager, ScheduleManager, LogViewer
 */

// ─── Utilities ────────────────────────────────────────────────────────────
const $ = (sel, ctx) => (ctx || document).querySelector(sel);
const $$ = (sel, ctx) => [...(ctx || document).querySelectorAll(sel)];

function esc(str) {
  const d = document.createElement('div');
  d.textContent = str;
  return d.innerHTML;
}

function fmtDate(iso) {
  if (!iso) return '-';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleString('zh-CN', { hour12: false });
}

function statusBadge(s) {
  const map = {
    pending: ['#f59e0b', '⏳ 等待中'],
    running: ['#3b82f6', '▶️ 运行中'],
    completed: ['#10b981', '✅ 已完成'],
    failed: ['#ef4444', '❌ 失败'],
    success: ['#10b981', '✅ 成功'],
    error: ['#ef4444', '❌ 错误'],
    active: ['#10b981', '✅ 活跃'],
    paused: ['#f59e0b', '⏸️ 暂停'],
  };
  const [color, label] = map[s] || ['#6b7280', s];
  return `<span style="display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:12px;font-size:0.8rem;font-weight:500;background:${color}20;color:${color};border:1px solid ${color}40;">${label}</span>`;
}

async function fetchJSON(url, opts) {
  const res = await fetch(url, opts);
  const data = await res.json();
  if (!res.ok) throw new Error(data.message || `HTTP ${res.status}`);
  return data;
}

// ─── Toast ─────────────────────────────────────────────────────────────────
function showToast(msg, type) {
  const icons = { success: '✅', error: '❌', info: 'ℹ️', warn: '⚠️' };
  const t = document.createElement('div');
  t.className = `toast toast-${type || 'info'}`;
  t.innerHTML = `<span>${icons[type] || 'ℹ️'}</span> ${esc(msg)}`;
  const ct = $('#toast-container') || (() => { const c = document.createElement('div'); c.id = 'toast-container'; document.body.appendChild(c); return c; })();
  ct.appendChild(t);
  setTimeout(() => { t.style.opacity = '0'; t.style.transform = 'translateX(40px)'; setTimeout(() => t.remove(), 300); }, 3500);
}

// ─── Console ───────────────────────────────────────────────────────────────
const consoleOut = $('#console-output');
function appendLog(text, isErr) {
  if (!consoleOut) return;
  const div = document.createElement('div');
  div.className = isErr ? 'log-error' : 'log-info';
  const now = new Date();
  const ts = `[${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}:${String(now.getSeconds()).padStart(2,'0')}] `;
  div.textContent = ts + text;
  if (consoleOut.childNodes.length > 500) consoleOut.removeChild(consoleOut.firstChild);
  consoleOut.appendChild(div);
  consoleOut.scrollTop = consoleOut.scrollHeight;
}

// ─── API Endpoints ─────────────────────────────────────────────────────────
const API = {
  taskList: '/api/task_list',
  runTask: (id) => `/api/run_task/${id}`,
  queue: '/api/queue',
  queueBatch: '/api/queue/batch',
  queueItem: (id) => `/api/queue/${id}`,
  queueRetry: (id) => `/api/queue/${id}/retry`,
  queueClear: '/api/queue/clear',
  queueStart: '/api/queue/start',
  queueStop: '/api/queue/stop',
  queueStatus: '/api/queue/status',
  schedules: '/api/schedules',
  scheduleItem: (id) => `/api/schedules/${id}`,
  scheduleToggle: (id) => `/api/schedules/${id}/toggle`,
  runNow: '/api/schedules/run_now',
  logs: '/api/logs',
  logItem: (id) => `/api/logs/${id}`,
  logStats: '/api/logs/stats',
  logCleanup: '/api/logs/cleanup',
  previewExcel: '/api/preview/excel',
  previewFile: '/api/preview/file',
  excelSheets: '/api/preview/excel/sheets',
  githubTrending: '/api/github/trending',
  dbConnections: '/api/connections',
};

// ─── Theme / Removed ──────────────────────────────────────────────────────────

// ─── Navigation ─────────────────────────────────────────────────────────────
let allTasks = [];
let currentPanel = 'dashboard';

function switchPanel(name, titleMain, titleSub) {
  currentPanel = name;
  $$('.nav-item').forEach(el => el.classList.toggle('active', el.dataset.panel === name));
  $$('.dashboard-panel').forEach(el => el.classList.toggle('active', el.id === `panel-${name}`));
  
  const topH2 = document.querySelector('.header-left h2');
  if (name === 'dashboard') {
    if (topH2) topH2.style.display = 'block';
  } else {
    if (topH2) topH2.style.display = 'none';
  }

  if (titleMain) {
    $('#current-breadcrumb').textContent = titleMain;
    $('#header-title-main').textContent = titleMain;
    $('#header-title-sub').textContent = titleSub || '';
  }

  if (name === 'dashboard') DashboardManager.refresh();
  if (name === 'queue') QueueManager.refresh();
  if (name === 'schedules') ScheduleManager.refresh();
  if (name === 'logs') LogViewer.refresh();
  if (name === 'db-connections') DBConnectionManager.refresh();
}

// ─── DashboardManager ───────────────────────────────────────────────────────
const DashboardManager = {
  clockInterval: null,
  sysStatusInterval: null,

  init() {
    this.startClock();
    this.startSysStatus();
    this.refresh();
    $('#btn-refresh-dashboard')?.addEventListener('click', () => this.refresh());
  },

  startClock() {
    if (this.clockInterval) clearInterval(this.clockInterval);
    const update = () => {
      const now = new Date();
      const days = ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'];
      $('#clock-date').textContent = `${now.getFullYear()}/${now.getMonth()+1}/${now.getDate()} ${days[now.getDay()]}`;
      $('#clock-time').textContent = `${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}:${String(now.getSeconds()).padStart(2,'0')}`;
    };
    update();
    this.clockInterval = setInterval(update, 1000);
  },

  startSysStatus() {
    if (this.sysStatusInterval) clearInterval(this.sysStatusInterval);
    const update = async () => {
      try {
        const res = await fetchJSON('/api/system/status');
        if (res.status === 'success') {
          const { cpu, memory } = res.data;
          const sysStatusText = $('#sys-status-text');
          const cpuBar = $('#cpu-bar-fill');
          const memBar = $('#mem-bar-fill');
          if (sysStatusText) sysStatusText.textContent = `CPU: ${cpu.toFixed(1)}% | 内存: ${memory.toFixed(1)}%`;
          if (cpuBar) {
            cpuBar.style.width = `${cpu}%`;
            cpuBar.style.background = cpu > 85 ? '#ef4444' : cpu > 60 ? '#f59e0b' : '#4f46e5';
          }
          if (memBar) {
            memBar.style.width = `${memory}%`;
            memBar.style.background = memory > 85 ? '#ef4444' : memory > 60 ? '#f59e0b' : '#059669';
          }
        }
      } catch (e) {
        // silent
      }
    };
    update();
    this.sysStatusInterval = setInterval(update, 3000);
  },

  async refresh() {
    try {
      // Fetch stats
      const [tasksRes, schedRes] = await Promise.all([
        fetchJSON(API.taskList),
        fetchJSON(API.schedules)
      ]);
      
      const tasks = tasksRes.tasks || [];
      const schedules = schedRes.schedules || [];
      
      $('#stat-total-tasks').textContent = tasks.length;
      $('#stat-total-schedules').textContent = schedules.length;

      // Render Activity List (Mix of tasks and schedules)
      const list = $('#dashboard-activity-list');
      list.innerHTML = '';
      
      let activities = [];
      
      // Top 3 common tasks
      tasks.slice(0, 3).forEach(t => {
        activities.push({
          type: 'task', text: `常用功能：<strong>${t.name}</strong> 待命运行中`, time: '刚刚',
        });
      });

      // Scheduled workflows
      schedules.forEach(s => {
        const tName = tasks.find(t => t.id === s.task_id)?.name || s.task_id;
        const nextTime = s.next_run_time ? new Date(s.next_run_time).toLocaleTimeString() : '未定';
        activities.push({
          type: 'sched', text: `定时任务：<strong>${tName}</strong> 下次运行于 <span style="color:var(--primary)">${nextTime}</span>`, time: '计划中',
        });
      });

      if (activities.length === 0) {
        list.innerHTML = '<li class="activity-item"><div class="activity-content" style="color:var(--text-muted)">暂无动态数据</div></li>';
      } else {
        activities.slice(0, 5).forEach(a => {
          const li = document.createElement('li');
          li.className = 'activity-item';
          const isOrange = a.type === 'sched' ? 'orange' : '';
          li.innerHTML = `
            <div class="activity-dot ${isOrange}"></div>
            <div class="activity-content">${a.text}</div>
            <div class="activity-time">${a.time}</div>
          `;
          list.appendChild(li);
        });
      }

    } catch (e) {
      console.error("Dashboard refresh failed", e);
    }
  }
};

// ─── Sidebar Builder ───────────────────────────────────────────────────────
async function buildSidebar() {
  try {
    const data = await fetchJSON(API.taskList + '?t=' + Date.now());
    allTasks = data.tasks || [];
  } catch (e) {
    console.error('Failed to load tasks', e);
    allTasks = [];
  }

  const catMap = {};
  allTasks.forEach(t => {
    const cat = t.category || 'other';
    if (!catMap[cat]) catMap[cat] = [];
    catMap[cat].push(t);
  });

  const sidebarNav = $('#sidebar-nav');
  sidebarNav.innerHTML = '';

  // Dashboard top link
  const dashboardItem = document.createElement('div');
  dashboardItem.className = 'nav-item nav-main active';
  dashboardItem.dataset.panel = 'dashboard';
  dashboardItem.innerHTML = `<span class="icon">📊</span><span>首页大盘</span>`;
  dashboardItem.addEventListener('click', () => switchPanel('dashboard', '首页大盘', 'Dashboard'));
  sidebarNav.appendChild(dashboardItem);

  const catNames = { archive: '文件归档与管理', processing: '内容清洗与加工', extraction: '报表信息抽取', workflow: '工作流引擎', other: '其他工具', pdf: 'PDF 管理' };
  const catIcons = { archive: '📦', processing: '⚙️', extraction: '📊', workflow: '🔀', other: '🛠️', pdf: '📄' };

  Object.keys(catMap).forEach(cat => {
    const section = document.createElement('div');
    section.className = 'menu-section';
    section.innerHTML = `
      <h3 class="menu-category"><span>${catIcons[cat] || '📋'} ${catNames[cat] || cat}</span><span class="chevron"></span></h3>
      <ul class="menu" id="menu-${cat}"></ul>
    `;
    const ul = section.querySelector('ul');
    catMap[cat].forEach(task => {
      const li = document.createElement('li');
      li.className = 'nav-item';
      li.dataset.panel = 'tasks';
      li.dataset.taskId = task.id;
      li.innerHTML = `<span>${esc(task.name)}</span>`;
      li.addEventListener('click', () => {
        switchPanel('tasks', '任务执行', 'Task Runner');
        TaskRunner.renderForm(task);
        $$('.nav-item').forEach(el => el.classList.remove('active'));
        li.classList.add('active');
      });
      ul.appendChild(li);
    });

    section.querySelector('.menu-category').addEventListener('click', () => {
      section.classList.toggle('expanded');
    });

    sidebarNav.appendChild(section);
  });

  // Fixed nav: Queue, Schedules, Logs
  const fixedSection = document.createElement('div');
  fixedSection.className = 'menu-section';
  fixedSection.innerHTML = '<h3 class="menu-category"><span>📋 系统管理</span><span class="chevron"></span></h3><ul class="menu" id="menu-system"></ul>';
  const sysUl = fixedSection.querySelector('ul');

  [
    { panel: 'queue', icon: '📬', label: '任务队列' },
    { panel: 'schedules', icon: '⏰', label: '定时任务' },
    { panel: 'logs', icon: '📜', label: '操作日志' },
    { panel: 'db-connections', icon: '🔌', label: '数据库连接' },
    { panel: 'github', icon: '🔥', label: 'GitHub Trending' },
  ].forEach(item => {
    const li = document.createElement('li');
    li.className = 'nav-item';
    li.dataset.panel = item.panel;
    li.innerHTML = `<span class="icon">${item.icon}</span><span>${item.label}</span>`;
    li.addEventListener('click', () => {
      switchPanel(item.panel, item.label, 'System Tools');
      if (item.panel === 'queue') QueueManager.refresh();
      if (item.panel === 'schedules') ScheduleManager.refresh();
      if (item.panel === 'logs') LogViewer.refresh();
      if (item.panel === 'db-connections') DBConnectionManager.init();
      if (item.panel === 'github') GithubTrending.init();
    });
    sysUl.appendChild(li);
  });

  fixedSection.querySelector('.menu-category').addEventListener('click', () => {
    fixedSection.classList.toggle('expanded');
  });

  sidebarNav.appendChild(fixedSection);
  DashboardManager.init();
}

// ─── TaskRunner ─────────────────────────────────────────────────────────────
const TaskRunner = {
  currentTask: null,

  renderForm(task) {
    this.currentTask = task;
    const container = $('#task-form-container');
    const title = $('#task-panel-title');
    const desc = $('#task-panel-desc');

    title.textContent = task.name;
    desc.textContent = task.description || '';

    container.innerHTML = '';
    const form = document.createElement('form');
    form.className = 'modern-form';
    form.id = 'form-task-run';

    const params = task.params || [];
    for (let i = 0; i < params.length; i++) {
      // ── 全宽信息面板（info 类型）─────────────────────────────────────────
      const pi = params[i];
      if (pi.type === 'info') {
        const card = document.createElement('div');
        card.className = 'form-group';
        const opsRows = (pi.operators || []).map(([op, lbl, note]) =>
          `<tr>
            <td style="font-family:monospace;color:#3b82f6;padding:4px 14px 4px 0;white-space:nowrap;font-size:0.82rem;">${esc(op)}</td>
            <td style="padding:4px 14px 4px 0;font-weight:500;font-size:0.82rem;">${esc(lbl)}</td>
            <td style="color:var(--text-muted);font-size:0.78rem;">${esc(note)}</td>
          </tr>`
        ).join('');
        const colHint = (pi.col_hint || []).map(([type, names]) =>
          `<tr>
            <td style="font-weight:500;padding:3px 14px 3px 0;font-size:0.82rem;">${esc(type)}</td>
            <td style="font-family:monospace;color:var(--text-muted);font-size:0.78rem;">${esc(names)}</td>
          </tr>`
        ).join('');
        card.innerHTML = `
          <details style="background:rgba(59,130,246,0.07);border:1px solid rgba(59,130,246,0.22);border-radius:10px;">
            <summary style="cursor:pointer;padding:10px 16px;font-weight:600;color:#3b82f6;font-size:0.88rem;
                           list-style:none;display:flex;align-items:center;gap:8px;user-select:none;">
              ${esc(pi.label || '格式说明')}
              <span style="font-size:0.72rem;opacity:0.65;font-weight:400;">点击展开查看</span>
            </summary>
            <div style="padding:2px 16px 14px;border-top:1px solid rgba(59,130,246,0.15);">
              ${ colHint ? `
              <p style="margin:10px 0 6px;font-size:0.82rem;font-weight:500;">📌 表头要求（首行，列名支持中英文）：</p>
              <table style="border-collapse:collapse;"><tbody>${colHint}</tbody></table>` : '' }
              ${ opsRows ? `
              <p style="margin:12px 0 6px;font-size:0.82rem;font-weight:500;">🔧 支持的操作符：</p>
              <table style="width:100%;border-collapse:collapse;">
                <thead><tr style="border-bottom:1px solid rgba(59,130,246,0.2);">
                  <th style="text-align:left;padding:4px 14px 4px 0;color:#3b82f6;font-size:0.8rem;">操作符</th>
                  <th style="text-align:left;padding:4px 14px 4px 0;font-size:0.8rem;">含义</th>
                  <th style="text-align:left;padding:4px 0;font-size:0.8rem;">备注</th>
                </tr></thead>
                <tbody>${opsRows}</tbody>
              </table>` : '' }
              ${ pi.example ? `
              <p style="margin:10px 0 0;font-size:0.78rem;color:var(--text-muted);">
                💡 示例：<span style="font-family:monospace;background:rgba(0,0,0,0.15);padding:1px 8px;border-radius:4px;">${esc(pi.example)}</span>
              </p>` : '' }
            </div>
          </details>
        `;
        form.appendChild(card);
        continue;
      }

      const row = document.createElement('div');
      row.className = 'form-group row';

      let _skipAdvance = false;
      for (let j = 0; j < 2 && i + j < params.length; j++) {
        const p = params[i + j];
        if (p.type === 'condition_builder') continue;
        if (p.type === 'info') { _skipAdvance = true; continue; }  // info 由外层渲染，不消耗
        const col = document.createElement('div');
        col.className = 'col';
        col.dataset.field = p.id;

        const label = document.createElement('label');
        label.textContent = (p.label || p.id) + (p.required !== false ? ' *' : '');

        const wrapper = document.createElement('div');
        wrapper.className = 'input-with-clear';

        let input;
        if (p.type === 'db_connection_select') {
          // 动态数据库连接下拉：渲染时自动拉取已保存的连接列表
          input = document.createElement('select');
          input.name = p.id;
          input.id = `db-conn-select-${p.id}`;
          input.style.cssText = 'flex:1;';
          input.innerHTML = '<option value="">— 选择已保存的数据库连接 —</option>';
          setTimeout(async () => {
            try {
              const data = await fetchJSON('/api/connections');
              input.innerHTML = '<option value="">— 选择已保存的数据库连接 —</option>';
              (data.connections || []).forEach(c => {
                const o = document.createElement('option');
                o.value = c.name;
                o.textContent = `${c.name} (${c.host}:${c.port}/${c.database_name})`;
                input.appendChild(o);
              });
            } catch (e) {
              input.innerHTML = '<option value="">⚠️ 加载失败</option>';
            }
          }, 0);

        } else if (p.type === 'sheet_select') {
          input = document.createElement('select');
          input.name = p.id;
          input.id = `sheet-select-${p.id}`;
          input.style.cssText = 'flex:1;';
          const placeholderOpt = document.createElement('option');
          placeholderOpt.value = '';
          placeholderOpt.textContent = '— 请先填写文件路径 —';
          input.appendChild(placeholderOpt);

          let lastFetchedPath = null;
          // 拉取 Sheet 列表并刷新下拉
          const fetchSheets = async (filePath) => {
            if (!filePath || filePath === lastFetchedPath) return;
            lastFetchedPath = filePath;
            const currentVal = input.value;
            
            input.disabled = true;
            input.innerHTML = '<option value="">⏳ 读取中...</option>';
            try {
              const data = await fetchJSON(`${API.excelSheets}?path=${encodeURIComponent(filePath)}`);
              input.innerHTML = '';
              const emptyOpt = document.createElement('option');
              emptyOpt.value = '';
              emptyOpt.textContent = '— 默认（第一个 Sheet）—';
              input.appendChild(emptyOpt);
              (data.sheets || []).forEach(name => {
                const o = document.createElement('option');
                o.value = name;
                o.textContent = name;
                input.appendChild(o);
              });
              
              // 尝试恢复之前已选的值
              if (currentVal) input.value = currentVal;
              
              input.disabled = false;
            } catch (e) {
              lastFetchedPath = null;
              input.innerHTML = '<option value="">⚠️ 读取失败，请检查路径</option>';
              input.disabled = false;
            }
          };

          // 延迟绑定：等表单渲染完成后找到联动字段
          const srcRef = p.src_ref;
          if (srcRef) {
            setTimeout(() => {
              const srcInput = form.querySelector(`[name="${srcRef}"]`);
              if (srcInput) {
                srcInput.addEventListener('blur', () => fetchSheets(srcInput.value.trim()));
                srcInput.addEventListener('change', () => fetchSheets(srcInput.value.trim()));
                // 如果已有内容则立即拉取
                if (srcInput.value.trim()) fetchSheets(srcInput.value.trim());
              }
            }, 0);
          }

        } else if (p.type === 'select' && p.options) {
          input = document.createElement('select');
          input.name = p.id;
          p.options.forEach(opt => {
            const o = document.createElement('option');
            o.value = opt;
            o.textContent = opt;
            input.appendChild(o);
          });
        } else {
          input = document.createElement('input');
          input.type = p.type || 'text';
          input.name = p.id;
          if (p.placeholder) input.placeholder = p.placeholder;
          if (p.required !== false) input.required = true;
          if (p.default) input.value = p.default;
        }

        const clearBtn = document.createElement('button');
        clearBtn.type = 'button';
        clearBtn.className = 'btn-clear';
        clearBtn.tabIndex = -1;
        clearBtn.textContent = '×';

        wrapper.appendChild(input);
        if (p.type !== 'select' && p.type !== 'sheet_select') wrapper.appendChild(clearBtn);
        col.appendChild(label);
        col.appendChild(wrapper);
        row.appendChild(col);
      }
      if (!_skipAdvance) i += 1;  // info 类型占位时不跳过它，让外层循环自己处理
      form.appendChild(row);
    }

    const actions = document.createElement('div');
    actions.className = 'form-actions';
    actions.style.cssText = 'display:flex;gap:12px;margin-top:20px;';
    actions.innerHTML = `
      <button type="submit" class="btn-primary" style="flex:2;">⚡ 立即执行</button>
      <button type="button" id="btn-add-to-queue" class="btn-secondary" style="flex:1;">📬 加入队列</button>
    `;
    form.appendChild(actions);

    // ── 数据库连接选择关联：选连接则隐藏手动字段 ──
    const connSelect = form.querySelector('select[name="connection_name"]');
    if (connSelect) {
      const manualFields = ['host', 'port', 'user', 'password', 'database'];
      const toggleManualFields = () => {
        const hasConn = !!connSelect.value;
        manualFields.forEach(fid => {
          const col = form.querySelector(`[data-field="${fid}"]`);
          if (col) {
            col.style.display = hasConn ? 'none' : '';
            const input = col.querySelector('input');
            if (input) input.required = !hasConn;
          }
        });
      };
      connSelect.addEventListener('change', toggleManualFields);
      setTimeout(toggleManualFields, 50);
    }

    // ── 表配置方式切换：手动输入 JSON / 导入配置文件 ──
    const tablesModeSelect = form.querySelector('select[name="tables_mode"]');
    if (tablesModeSelect) {
      const cfgTextarea = form.querySelector('[data-field="tables_cfg"]');
      const fileField = form.querySelector('[data-field="tables_file"]');
      const toggleTablesMode = () => {
        const isFile = tablesModeSelect.value === "导入配置文件";
        if (cfgTextarea) cfgTextarea.style.display = isFile ? 'none' : '';
        if (fileField) fileField.style.display = isFile ? '' : 'none';
      };
      tablesModeSelect.addEventListener('change', toggleTablesMode);
      setTimeout(toggleTablesMode, 50);
    }

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(form);
      const payload = Object.fromEntries(fd.entries());
      await this.executeTask(task.id, payload, form.querySelector('button[type="submit"]'));
    });

    form.querySelector('#btn-add-to-queue').addEventListener('click', async () => {
      const fd = new FormData(form);
      const payload = Object.fromEntries(fd.entries());
      try {
        await fetchJSON(API.queue, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ task_id: task.id, params: payload }),
        });
        showToast('已加入队列', 'success');
      } catch (err) {
        showToast(err.message, 'error');
      }
    });

    container.appendChild(form);
  },

  async executeTask(taskId, payload, btn) {
    const orig = btn.innerHTML;
    btn.classList.add('loading');
    btn.innerHTML = '⏳ 处理中...';
    consoleOut.textContent = `🚀 [${new Date().toLocaleTimeString()}] >> 启动: ${taskId}\n`;

    try {
      const resp = await fetch(API.runTask(taskId), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      if (!resp.ok) {
        const err = await resp.json().catch(() => ({}));
        throw new Error(err.message || `HTTP ${resp.status}`);
      }
      const reader = resp.body.getReader();
      const decoder = new TextDecoder();
      while (true) {
        const { done, value } = await reader.read();
        if (done) { consoleOut.textContent += '\n✨ 任务执行完毕。'; break; }
        consoleOut.textContent += decoder.decode(value, { stream: true });
        consoleOut.scrollTop = consoleOut.scrollHeight;
      }
    } catch (err) {
      appendLog(`❌ 执行失败: ${err.message}`, true);
    } finally {
      btn.classList.remove('loading');
      btn.innerHTML = orig;
    }
  },
};

// ─── QueueManager ───────────────────────────────────────────────────────────
const QueueManager = {
  pollingTimer: null,
  items: [],
  stats: {},
  processing: false,

  async refresh() {
    try {
      const [qData, sData] = await Promise.all([
        fetchJSON(API.queue),
        fetchJSON(API.queueStatus),
      ]);
      this.items = qData.queue || [];
      this.stats = sData.stats || {};
      this.processing = sData.processing || false;
      this.render();
    } catch (err) {
      console.error('Queue refresh failed', err);
    }
  },

  render() {
    // 捕获当前的勾选状态，防止 polling 刷新时丢失
    const checkedIds = new Set($$('.queue-select:checked').map(cb => cb.value));
    const allChecked = $('#queue-select-all')?.checked || false;

    const container = $('#queue-content');
    if (!container) return;

    const { pending = 0, running = 0, completed = 0, failed = 0 } = this.stats;

    container.innerHTML = `
      <div class="stats-row">
        <div class="stat-card-mini"><div class="stat-val" style="color:#f59e0b">${pending}</div><div class="stat-lbl">等待中</div></div>
        <div class="stat-card-mini"><div class="stat-val" style="color:#3b82f6">${running}</div><div class="stat-lbl">运行中</div></div>
        <div class="stat-card-mini"><div class="stat-val" style="color:#10b981">${completed}</div><div class="stat-lbl">已完成</div></div>
        <div class="stat-card-mini"><div class="stat-val" style="color:#ef4444">${failed}</div><div class="stat-lbl">失败</div></div>
      </div>
      <div class="queue-toolbar">
        <div class="toolbar-left">
          <span class="processing-indicator ${this.processing ? 'on' : 'off'}">${this.processing ? '🟢 处理中' : '🔴 已停止'}</span>
          <button class="btn-sm ${this.processing ? 'btn-warn' : 'btn-success'}" id="btn-toggle-worker">${this.processing ? '⏹️ 停止' : '▶️ 启动'}</button>
        </div>
        <div class="toolbar-right">
          <button class="btn-sm btn-secondary" id="btn-export-queue">📤 导出选中</button>
          <button class="btn-sm btn-danger" id="btn-delete-selected" style="margin-right: 8px;">❌ 删除选中</button>
          <button class="btn-sm btn-secondary" id="btn-import-queue">📥 导入</button>
          <input type="file" id="input-import-queue" accept=".json" style="display:none">
          <button class="btn-sm btn-secondary" id="btn-clear-completed">🧹 清除已完成</button>
          <button class="btn-sm btn-danger" id="btn-clear-all">🗑️ 清空队列</button>
          <button class="btn-sm btn-secondary" id="btn-refresh-queue">🔄 刷新</button>
        </div>
      </div>
      <div class="queue-table-wrap">
        <table class="modern-table" id="queue-table">
          <thead><tr>
            <th style="width:36px;text-align:center;"><input type="checkbox" id="queue-select-all" title="全选/取消全选" style="cursor:pointer;"></th>
            <th>任务</th><th>参数</th><th>状态</th><th>失败原因</th><th>优先级</th><th>创建时间</th><th>操作</th>
          </tr></thead>
          <tbody id="queue-tbody">
            ${this.items.length === 0 ? '<tr><td colspan="8" style="text-align:center;padding:30px;">📭 队列为空</td></tr>' : ''}
          </tbody>
        </table>
      </div>
    `;

    $('#btn-toggle-worker')?.addEventListener('click', () => this.toggleWorker());
    $('#btn-clear-completed')?.addEventListener('click', () => this.clearCompleted());
    $('#btn-clear-all')?.addEventListener('click', () => this.clearAll());
    $('#btn-refresh-queue')?.addEventListener('click', () => this.refresh());
    $('#btn-export-queue')?.addEventListener('click', () => this.exportSelected());
    $('#btn-delete-selected')?.addEventListener('click', () => this.deleteSelected());
    const _importInput = $('#input-import-queue');
    $('#btn-import-queue')?.addEventListener('click', () => _importInput?.click());
    _importInput?.addEventListener('change', (e) => {
      if (e.target.files[0]) { this.importFromFile(e.target.files[0]); e.target.value = ''; }
    });
    $('#queue-select-all')?.addEventListener('change', (e) => {
      $$('.queue-select').forEach(cb => cb.checked = e.target.checked);
    });

    const tbody = $('#queue-tbody');
    this.items.forEach(item => {
      const tr = document.createElement('tr');
      let paramsStr = '-';
      try { paramsStr = typeof item.params === 'string' ? item.params : JSON.stringify(item.params || {}); } catch (e) {}
      if (paramsStr.length > 60) paramsStr = paramsStr.slice(0, 60) + '...';

      // 失败原因列
      const errorFull = (item.status === 'failed' && item.error) ? item.error : '';
      const errorDisplay = errorFull.length > 80 ? errorFull.slice(0, 80) + '…' : errorFull;
      const errorCell = errorFull
        ? `<td style="font-family:monospace;font-size:0.78rem;color:#ef4444;max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${esc(errorFull)}">${esc(errorDisplay)}</td>`
        : `<td style="color:var(--text-muted);font-size:0.85rem;">-</td>`;

      tr.innerHTML = `
        <td style="text-align:center;"><input type="checkbox" class="queue-select" value="${item.id}" style="cursor:pointer;" ${checkedIds.has(item.id) ? 'checked' : ''}></td>
        <td><strong>${esc(item.task_name || item.task_id)}</strong></td>
        <td style="font-family:monospace;font-size:0.8rem;color:var(--text-muted);max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${esc(paramsStr)}">${esc(paramsStr)}</td>
        <td>${statusBadge(item.status)}</td>
        ${errorCell}
        <td>${item.priority || 0}</td>
        <td style="font-size:0.85rem;">${fmtDate(item.created_at)}</td>
        <td>
          ${item.status === 'pending' ? `<button class="btn-sm btn-danger" data-remove="${item.id}">移除</button>` : ''}
          ${item.status === 'failed'  ? `<button class="btn-sm btn-primary" data-retry="${item.id}" style="padding:2px 8px;font-size:0.78rem;line-height:1.4;">🔁 重试</button>` : ''}
        </td>
      `;
      tr.querySelector('[data-remove]')?.addEventListener('click', () => this.removeItem(item.id));
      tr.querySelector('[data-retry]')?.addEventListener('click',  () => this.retryItem(item.id));
      tbody.appendChild(tr);
    });

    const selectAllCb = $('#queue-select-all');
    if (selectAllCb) selectAllCb.checked = allChecked;

    if (!this.pollingTimer) this.startPolling();
  },

  startPolling() {
    if (this.pollingTimer) return;
    this.pollingTimer = setInterval(() => this.refreshSilent(), 3000);
  },

  stopPolling() {
    if (this.pollingTimer) { clearInterval(this.pollingTimer); this.pollingTimer = null; }
  },

  async refreshSilent() {
    try {
      const [qData, sData] = await Promise.all([
        fetchJSON(API.queue),
        fetchJSON(API.queueStatus),
      ]);
      this.items = qData.queue || [];
      this.stats = sData.stats || {};
      this.processing = sData.processing || false;
      this.render();
    } catch (e) { /* silent */ }
  },

  async toggleWorker() {
    try {
      if (this.processing) {
        await fetchJSON(API.queueStop, { method: 'POST' });
        showToast('Worker 已停止', 'info');
      } else {
        await fetchJSON(API.queueStart, { method: 'POST' });
        showToast('Worker 已启动', 'success');
      }
      await this.refresh();
    } catch (err) { showToast(err.message, 'error'); }
  },

  async clearCompleted() {
    try {
      const data = await fetchJSON(API.queueClear, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ scope: 'completed' }),
      });
      showToast(`已清除 ${data.removed || 0} 个已完成项`, 'success');
      await this.refresh();
    } catch (err) { showToast(err.message, 'error'); }
  },

  async clearAll() {
    if (!confirm('确定要清空所有队列项（包括等待中的）吗？')) return;
    try {
      const data = await fetchJSON(API.queueClear, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ scope: 'all' }),
      });
      showToast(`已清空 ${data.removed || 0} 个队列项`, 'success');
      await this.refresh();
    } catch (err) { showToast(err.message, 'error'); }
  },

  async removeItem(id) {
    try {
      await fetchJSON(API.queueItem(id), { method: 'DELETE' });
      showToast('已移除', 'success');
      await this.refresh();
    } catch (err) { showToast(err.message, 'error'); }
  },

  async retryItem(id) {
    try {
      await fetchJSON(API.queueRetry(id), { method: 'POST' });
      showToast('已重新加入队列，等待 Worker 处理', 'success');
      // 若 Worker 未启动则自动启动
      if (!this.processing) {
        await fetchJSON(API.queueStart, { method: 'POST' });
      }
      await this.refresh();
    } catch (err) { showToast(err.message, 'error'); }
  },

  async deleteSelected() {
    const checked = $$('.queue-select:checked');
    if (checked.length === 0) {
      showToast('请先勾选要删除的任务', 'warn');
      return;
    }
    if (!confirm(`确定要删除选中的 ${checked.length} 个任务吗？\n（注：运行中的任务无法删除）`)) return;

    let successCount = 0;
    let failCount = 0;
    
    // 禁用按钮防连点
    const btn = $('#btn-delete-selected');
    const origText = btn.innerHTML;
    btn.innerHTML = '⏳ 删除中...';
    btn.disabled = true;

    for (const cb of checked) {
      try {
        await fetchJSON(API.queueItem(cb.value), { method: 'DELETE' });
        successCount++;
      } catch (e) {
        failCount++;
        console.error('Failed to delete', cb.value, e);
      }
    }
    
    btn.innerHTML = origText;
    btn.disabled = false;

    if (failCount > 0) {
      showToast(`已删除 ${successCount} 个，${failCount} 个失败`, 'warn');
    } else {
      showToast(`已成功删除 ${successCount} 个任务`, 'success');
    }
    await this.refresh();
  },

  // ── 导出选中项为 JSON 文件 ──────────────────────────────────────────────
  exportSelected() {
    const checked = $$('.queue-select:checked');
    const selectedIds = new Set(checked.map(c => c.value));
    const toExport = selectedIds.size > 0
      ? this.items.filter(item => selectedIds.has(item.id))
      : this.items;

    if (toExport.length === 0) {
      showToast('没有可导出的队列项', 'warn');
      return;
    }

    const exportData = {
      exported_at: new Date().toISOString(),
      version: '1.0',
      note: '可直接修改 params / priority 后重新导入执行',
      tasks: toExport.map(item => ({
        task_id:   item.task_id,
        task_name: item.task_name,
        params:    typeof item.params === 'string' ? JSON.parse(item.params) : (item.params || {}),
        priority:  item.priority || 0,
      }))
    };

    const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `queue_export_${new Date().toISOString().slice(0,10)}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast(`已导出 ${toExport.length} 个任务`, 'success');
  },

  // ── 解析上传的 JSON 文件 ────────────────────────────────────────────────
  importFromFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data  = JSON.parse(e.target.result);
        const tasks = data.tasks || (Array.isArray(data) ? data : []);
        if (!tasks.length) { showToast('文件中没有找到任务', 'warn'); return; }
        this.showImportPreview(tasks);
      } catch (err) {
        showToast('文件解析失败: ' + err.message, 'error');
      }
    };
    reader.readAsText(file, 'utf-8');
  },

  // ── 导入预览弹窗 ────────────────────────────────────────────────────────
  showImportPreview(tasks) {
    const rows = tasks.map((t, i) => {
      const ps = JSON.stringify(t.params || {});
      const psShort = ps.length > 60 ? ps.slice(0, 60) + '…' : ps;
      return `
        <tr>
          <td style="color:var(--text-muted);">${i + 1}</td>
          <td><strong>${esc(t.task_name || t.task_id)}</strong><br>
            <span style="font-size:0.75rem;color:var(--text-muted);">${esc(t.task_id)}</span></td>
          <td style="font-family:monospace;font-size:0.78rem;color:var(--text-muted);max-width:280px;
                     overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"
              title="${esc(ps)}">${esc(psShort)}</td>
          <td style="text-align:center;">${t.priority || 0}</td>
        </tr>`;
    }).join('');

    const modal = document.createElement('div');
    modal.className = 'modal-backdrop';
    modal.innerHTML = `
      <div class="modal-content wide" style="max-width:780px;">
        <h3>📥 导入预览</h3>
        <p style="color:var(--text-muted);margin:8px 0 16px;">
          共 <strong style="color:var(--primary);font-size:1.1em;">${tasks.length}</strong> 个任务将以 <strong>pending</strong> 状态加入队列末尾
        </p>
        <div style="max-height:340px;overflow-y:auto;border-radius:8px;border:1px solid var(--panel-border);">
          <table class="modern-table" style="margin:0;">
            <thead><tr>
              <th style="width:40px;">#</th>
              <th>任务</th>
              <th>参数</th>
              <th style="width:60px;text-align:center;">优先级</th>
            </tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
        <div style="display:flex;gap:12px;justify-content:flex-end;margin-top:20px;">
          <button class="btn-secondary" id="btn-cancel-import">取消</button>
          <button class="btn-primary" id="btn-confirm-import" style="margin-top:0;">✅ 确认导入 ${tasks.length} 个任务</button>
        </div>
      </div>
    `;
    document.body.appendChild(modal);

    modal.querySelector('#btn-cancel-import').addEventListener('click', () => modal.remove());
    modal.addEventListener('click', (e) => { if (e.target === modal) modal.remove(); });

    modal.querySelector('#btn-confirm-import').addEventListener('click', async () => {
      const btn = modal.querySelector('#btn-confirm-import');
      btn.textContent = '⏳ 导入中...';
      btn.disabled = true;
      try {
        const batchTasks = tasks.map(t => ({
          task_id:  t.task_id,
          params:   t.params || {},
          priority: t.priority || 0,
        }));
        const result = await fetchJSON(API.queueBatch, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tasks: batchTasks }),
        });
        const ok   = result.enqueued ?? result.total ?? tasks.length;
        const fail = (result.total ?? tasks.length) - ok;
        showToast(
          fail > 0
            ? `导入完成：${ok} 成功，${fail} 失败（任务ID不存在或参数不合法）`
            : `成功导入 ${ok} 个任务`,
          fail > 0 ? 'warn' : 'success'
        );
        modal.remove();
        await this.refresh();
      } catch (err) {
        showToast('导入失败: ' + err.message, 'error');
        btn.textContent = '✅ 确认导入';
        btn.disabled = false;
      }
    });
  },
};

// ─── ScheduleManager ────────────────────────────────────────────────────────
const ScheduleManager = {
  schedules: [],

  CRON_PRESETS: [
    { label: '每分钟', value: '* * * * *' },
    { label: '每5分钟', value: '*/5 * * * *' },
    { label: '每30分钟', value: '*/30 * * * *' },
    { label: '每小时', value: '0 * * * *' },
    { label: '每天 9:00', value: '0 9 * * *' },
    { label: '每天 18:00', value: '0 18 * * *' },
    { label: '工作日 9:00', value: '0 9 * * 1-5' },
    { label: '每周一 8:00', value: '0 8 * * 1' },
    { label: '每月1日 0:00', value: '0 0 1 * *' },
  ],

  async refresh() {
    try {
      const data = await fetchJSON(API.schedules);
      this.schedules = data.schedules || [];
      this.render();
    } catch (err) {
      console.error('Schedule refresh failed', err);
    }
  },

  render() {
    const container = $('#schedules-content');
    if (!container) return;

    const hasItems = this.schedules.length > 0;

    container.innerHTML = `
      <div class="schedules-toolbar">
        <button class="btn-primary btn-sm" id="btn-create-schedule">➕ 新建定时任务</button>
        <button class="btn-sm btn-secondary" id="btn-refresh-schedules">🔄 刷新</button>
      </div>
      <div class="schedules-grid" id="schedules-grid">
        ${!hasItems ? '<div class="empty-state">📭 暂无定时任务，点击上方按钮创建</div>' : ''}
      </div>
    `;

    $('#btn-create-schedule')?.addEventListener('click', () => this.showCreateModal());
    $('#btn-refresh-schedules')?.addEventListener('click', () => this.refresh());

    if (hasItems) {
      const grid = $('#schedules-grid');
      this.schedules.forEach(s => {
        const card = document.createElement('div');
        card.className = 'schedule-card glass-panel';
        const taskName = allTasks.find(t => t.id === s.task_id)?.name || s.task_id;
        const nextRun = s.next_run_time ? fmtDate(s.next_run_time) : '等待调度';
        card.innerHTML = `
          <div class="schedule-card-header">
            <span class="schedule-icon">⏰</span>
            <div class="schedule-info">
              <strong>${esc(taskName)}</strong>
              <span class="schedule-cron">${esc(s.trigger || '')} ${esc(s.cron || '')}</span>
            </div>
            <label class="toggle-switch">
              <input type="checkbox" ${s.enabled ? 'checked' : ''} data-toggle="${s.id}">
              <span class="toggle-slider"></span>
            </label>
          </div>
          <div class="schedule-card-body">
            <div class="schedule-meta">
              <span>📋 ID: ${esc(s.id).slice(0, 12)}...</span>
              <span>🕐 下次: ${nextRun}</span>
            </div>
            <div class="schedule-params" style="font-family:monospace;font-size:0.78rem;color:var(--text-muted);max-height:40px;overflow:hidden;">
              ${esc(JSON.stringify(s.params || {}))}
            </div>
          </div>
          <div class="schedule-card-actions">
            <button class="btn-sm btn-secondary" data-run="${s.id}">▶️ 立即执行</button>
            <button class="btn-sm btn-danger" data-delete="${s.id}">🗑️ 删除</button>
          </div>
        `;

        card.querySelector('[data-toggle]')?.addEventListener('change', (e) => {
          this.toggleSchedule(s.id, e.target.checked);
        });
        card.querySelector('[data-run]')?.addEventListener('click', () => this.runNow(s.task_id, s.params));
        card.querySelector('[data-delete]')?.addEventListener('click', () => this.deleteSchedule(s.id));

        grid.appendChild(card);
      });
    }
  },

  showCreateModal() {
    const taskOptions = allTasks.map(t =>
      `<option value="${esc(t.id)}">${esc(t.name)}</option>`
    ).join('');

    const cronPresetOptions = this.CRON_PRESETS.map(p =>
      `<option value="${p.value}">${p.label} (${p.value})</option>`
    ).join('');

    const modal = document.createElement('div');
    modal.className = 'modal-backdrop';
    modal.innerHTML = `
      <div class="modal-content wide">
        <h3>新建定时任务</h3>
        <form id="form-schedule" class="modern-form" style="margin-top:16px;">
          <div class="form-group">
            <label>任务 *</label>
            <select name="task_id" required>${taskOptions}</select>
          </div>
          <div class="form-group">
            <label>触发类型</label>
            <select name="trigger_type" id="trigger-type">
              <option value="cron">Cron 表达式</option>
              <option value="interval">间隔执行</option>
            </select>
          </div>
          <div id="cron-config" class="form-group">
            <label>Cron 表达式 *</label>
            <input type="text" name="cron" placeholder="*/5 * * * *" required list="cron-presets">
            <datalist id="cron-presets">${cronPresetOptions}</datalist>
            <div id="cron-preset-chips" style="display:flex;flex-wrap:wrap;gap:6px;margin-top:8px;"></div>
          </div>
          <div id="interval-config" class="form-group" style="display:none;">
            <label>间隔（秒）*</label>
            <input type="number" name="seconds" placeholder="3600" min="1">
          </div>
          <div class="form-group">
            <label>任务参数 (JSON)</label>
            <textarea name="params_json" rows="3" placeholder='{"key": "value"}' style="width:100%;background:rgba(0,0,0,0.15);border:1px solid var(--panel-border);color:var(--text-main);border-radius:8px;padding:8px;"></textarea>
          </div>
          <div class="form-actions" style="display:flex;gap:12px;justify-content:flex-end;">
            <button type="button" class="btn-secondary" id="btn-cancel-schedule">取消</button>
            <button type="submit" class="btn-primary" style="margin-top:0;">创建</button>
          </div>
        </form>
      </div>
    `;

    document.body.appendChild(modal);

    const chipsDiv = modal.querySelector('#cron-preset-chips');
    this.CRON_PRESETS.forEach(p => {
      const chip = document.createElement('span');
      chip.className = 'cron-chip';
      chip.textContent = p.label;
      chip.title = p.value;
      chip.addEventListener('click', () => {
        modal.querySelector('input[name="cron"]').value = p.value;
      });
      chipsDiv.appendChild(chip);
    });

    modal.querySelector('#trigger-type').addEventListener('change', (e) => {
      const isCron = e.target.value === 'cron';
      modal.querySelector('#cron-config').style.display = isCron ? '' : 'none';
      modal.querySelector('#interval-config').style.display = isCron ? 'none' : '';
    });

    modal.querySelector('#btn-cancel-schedule').addEventListener('click', () => modal.remove());
    modal.addEventListener('click', (e) => { if (e.target === modal) modal.remove(); });

    modal.querySelector('#form-schedule').addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const triggerType = fd.get('trigger_type');
      let triggerConfig = {};
      let params = {};

      if (triggerType === 'cron') {
        triggerConfig = { cron: fd.get('cron') };
      } else {
        triggerConfig = { seconds: parseInt(fd.get('seconds')) || 3600 };
      }

      try { params = JSON.parse(fd.get('params_json') || '{}'); } catch (e) { params = {}; }

      try {
        await fetchJSON(API.schedules, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            task_id: fd.get('task_id'),
            trigger_type: triggerType,
            trigger_config: triggerConfig,
            params,
          }),
        });
        showToast('定时任务已创建', 'success');
        modal.remove();
        this.refresh();
      } catch (err) {
        showToast(err.message, 'error');
      }
    });
  },

  async toggleSchedule(id, enabled) {
    try {
      await fetchJSON(API.scheduleToggle(id), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ enabled }),
      });
      showToast(enabled ? '已启用' : '已暂停', 'success');
    } catch (err) { showToast(err.message, 'error'); }
  },

  async deleteSchedule(id) {
    if (!confirm('确定要删除此定时任务吗？')) return;
    try {
      await fetchJSON(API.scheduleItem(id), { method: 'DELETE' });
      showToast('已删除', 'success');
      this.refresh();
    } catch (err) { showToast(err.message, 'error'); }
  },

  async runNow(taskId, params) {
    try {
      await fetchJSON(API.runNow, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ task_id: taskId, params }),
      });
      showToast('任务已触发执行', 'success');
    } catch (err) { showToast(err.message, 'error'); }
  },
};

// ─── LogViewer ──────────────────────────────────────────────────────────────
const LogViewer = {
  logs: [],
  stats: {},
  filter: { task_id: '', status: '' },

  async refresh() {
    try {
      const params = new URLSearchParams();
      if (this.filter.task_id) params.set('task_id', this.filter.task_id);
      if (this.filter.status) params.set('status', this.filter.status);
      params.set('limit', '100');

      const [logData, statsData] = await Promise.all([
        fetchJSON(API.logs + '?' + params.toString()),
        fetchJSON(API.logStats),
      ]);
      this.logs = logData.logs || [];
      this.stats = statsData.stats || {};
      this.render();
    } catch (err) {
      console.error('Log refresh failed', err);
    }
  },

  render() {
    const container = $('#logs-content');
    if (!container) return;

    const s = this.stats;
    const taskOptions = allTasks.map(t => `<option value="${esc(t.id)}">${esc(t.name)}</option>`).join('');

    container.innerHTML = `
      <div class="stats-row">
        <div class="stat-card-mini"><div class="stat-val" style="color:#10b981">${s.completed || 0}</div><div class="stat-lbl">成功</div></div>
        <div class="stat-card-mini"><div class="stat-val" style="color:#ef4444">${s.failed || 0}</div><div class="stat-lbl">失败</div></div>
        <div class="stat-card-mini"><div class="stat-val" style="color:#f59e0b">${s.total || 0}</div><div class="stat-lbl">总计</div></div>
      </div>
      <div class="logs-toolbar">
        <div class="toolbar-left">
          <select id="log-filter-task" style="width:180px;">
            <option value="">全部任务</option>
            ${taskOptions}
          </select>
          <select id="log-filter-status" style="width:120px;">
            <option value="">全部状态</option>
            <option value="success">成功</option>
            <option value="failed">失败</option>
            <option value="running">运行中</option>
          </select>
        </div>
        <div class="toolbar-right">
          <button class="btn-sm btn-secondary" id="btn-cleanup-logs">🧹 清理旧日志</button>
          <button class="btn-sm btn-secondary" id="btn-refresh-logs">🔄 刷新</button>
        </div>
      </div>
      <div class="queue-table-wrap">
        <table class="modern-table" id="logs-table">
          <thead><tr>
            <th>任务</th><th>状态</th><th>开始时间</th><th>耗时</th><th>操作</th>
          </tr></thead>
          <tbody id="logs-tbody">
            ${this.logs.length === 0 ? '<tr><td colspan="5" style="text-align:center;padding:30px;">📭 暂无日志记录</td></tr>' : ''}
          </tbody>
        </table>
      </div>
      <div id="log-detail-modal" class="modal-backdrop" style="display:none;"></div>
    `;

    $('#log-filter-task').value = this.filter.task_id;
    $('#log-filter-status').value = this.filter.status;
    $('#log-filter-task')?.addEventListener('change', (e) => {
      this.filter.task_id = e.target.value;
      this.refresh();
    });
    $('#log-filter-status')?.addEventListener('change', (e) => {
      this.filter.status = e.target.value;
      this.refresh();
    });

    $('#btn-cleanup-logs')?.addEventListener('click', () => this.cleanup());
    $('#btn-refresh-logs')?.addEventListener('click', () => this.refresh());

    const tbody = $('#logs-tbody');
    this.logs.forEach(log => {
      const tr = document.createElement('tr');
      const taskName = allTasks.find(t => t.id === log.task_id)?.name || log.task_id || '-';
      const duration = log.finished_at && log.started_at
        ? ((new Date(log.finished_at) - new Date(log.started_at)) / 1000).toFixed(1) + 's'
        : '-';
      tr.innerHTML = `
        <td><strong>${esc(taskName)}</strong></td>
        <td>${statusBadge(log.status)}</td>
        <td style="font-size:0.85rem;">${fmtDate(log.started_at)}</td>
        <td>${duration}</td>
        <td><button class="btn-sm btn-secondary" data-detail="${log.id}">📋 详情</button></td>
      `;
      tr.querySelector('[data-detail]')?.addEventListener('click', () => this.showDetail(log.id));
      tbody.appendChild(tr);
    });
  },

  async showDetail(logId) {
    const modal = $('#log-detail-modal');
    modal.style.display = 'flex';
    modal.innerHTML = `<div class="modal-content wide" style="max-width:800px;"><div style="text-align:center;padding:40px;">⏳ 加载中...</div></div>`;

    try {
      const data = await fetchJSON(API.logItem(logId));
      const log = data.log;

      let filePaths = [];
      const stdout = log.stdout || '';
      const pathRx = /([A-Za-z]:\\[^\s\n"'<>|*?]+(?:\.xlsx|\.csv|\.txt|\.pdf|\.docx|\.png|\.jpg|\.json))/gi;
      filePaths = [...new Set([...stdout.matchAll(pathRx)].map(m => m[1]))];

      const taskName = allTasks.find(t => t.id === log.task_id)?.name || log.task_id || '-';

      modal.innerHTML = `
        <div class="modal-content wide" style="max-width:800px;">
          <div class="modal-header">
            <h3>📋 执行详情</h3>
            <button class="btn-icon" id="btn-close-detail" style="font-size:1.5rem;">&times;</button>
          </div>
          <div class="log-detail-body">
            <div class="detail-row"><strong>任务:</strong> ${esc(taskName)} (${esc(log.task_id)})</div>
            <div class="detail-row"><strong>状态:</strong> ${statusBadge(log.status)}</div>
            <div class="detail-row"><strong>开始:</strong> ${fmtDate(log.started_at)}</div>
            <div class="detail-row"><strong>结束:</strong> ${fmtDate(log.finished_at)}</div>
            ${filePaths.length > 0 ? `
            <div class="detail-row">
              <strong>输出文件:</strong>
              <div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:6px;">
                ${filePaths.map(p => `
                  <span class="file-path-tag">
                    ${esc(p.split('\\').pop())}
                    <button class="btn-sm btn-secondary" data-preview="${esc(p)}" style="margin-left:6px;">👁️ 预览</button>
                  </span>
                `).join('')}
              </div>
            </div>` : ''}
            <div class="detail-tabs">
              <button class="detail-tab active" data-tab="stdout">📤 stdout</button>
              <button class="detail-tab" data-tab="stderr">📥 stderr</button>
              <button class="detail-tab" data-tab="params">⚙️ 参数</button>
            </div>
            <div class="detail-tab-content active" id="tab-stdout"><pre>${esc(stdout || '(无输出)')}</pre></div>
            <div class="detail-tab-content" id="tab-stderr"><pre style="color:var(--danger);">${esc(log.stderr || '(无错误输出)')}</pre></div>
            <div class="detail-tab-content" id="tab-params"><pre>${esc(JSON.stringify(log.params || {}, null, 2))}</pre></div>
          </div>
          <div id="file-preview-area" style="display:none;margin-top:16px;border-top:1px solid var(--panel-border);padding-top:16px;"></div>
        </div>
      `;

      modal.querySelectorAll('.detail-tab').forEach(tab => {
        tab.addEventListener('click', () => {
          modal.querySelectorAll('.detail-tab').forEach(t => t.classList.remove('active'));
          modal.querySelectorAll('.detail-tab-content').forEach(t => t.classList.remove('active'));
          tab.classList.add('active');
          modal.querySelector(`#tab-${tab.dataset.tab}`).classList.add('active');
        });
      });

      modal.querySelectorAll('[data-preview]').forEach(btn => {
        btn.addEventListener('click', async () => {
          const path = btn.dataset.preview;
          const area = modal.querySelector('#file-preview-area');
          area.style.display = 'block';
          area.innerHTML = '<div style="text-align:center;padding:20px;">⏳ 加载预览...</div>';

          const ext = path.split('.').pop().toLowerCase();
          if (['xlsx', 'xls', 'csv'].includes(ext)) {
            try {
              const previewData = await fetchJSON(`${API.previewExcel}?path=${encodeURIComponent(path)}&rows=20`);
              area.innerHTML = `
                <h4>📊 文件预览: ${esc(path.split('\\').pop())} (Sheet: ${previewData.sheet_name}, 共 ${previewData.total_rows} 行)</h4>
                <div class="queue-table-wrap" style="max-height:400px;"><table class="modern-table">
                  <thead><tr>${(previewData.headers || []).map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead>
                  <tbody>${(previewData.rows || []).map(row => `<tr>${row.map(c => `<td>${esc(String(c))}</td>`).join('')}</tr>`).join('')}</tbody>
                </table></div>
              `;
            } catch (err) {
              area.innerHTML = `<div class="log-error">❌ 预览失败: ${esc(err.message)}</div>`;
            }
          } else {
            try {
              const previewData = await fetchJSON(`${API.previewFile}?path=${encodeURIComponent(path)}&lines=50`);
              area.innerHTML = `
                <h4>📄 文件预览: ${esc(path.split('\\').pop())} (${previewData.lines} 行)</h4>
                <pre style="background:var(--terminal-bg);color:#a9dc76;padding:16px;border-radius:8px;max-height:400px;overflow:auto;font-size:0.85rem;">${esc(previewData.content)}</pre>
              `;
            } catch (err) {
              area.innerHTML = `<div class="log-error">❌ 预览失败: ${esc(err.message)}</div>`;
            }
          }
        });
      });

      modal.querySelector('#btn-close-detail').addEventListener('click', () => { modal.style.display = 'none'; });
      modal.addEventListener('click', (e) => { if (e.target === modal) modal.style.display = 'none'; });
    } catch (err) {
      modal.innerHTML = `<div class="modal-content"><div class="log-error">❌ 加载失败: ${esc(err.message)}</div></div>`;
    }
  },

  async cleanup() {
    const days = prompt('删除多少天前的日志？（默认30天）', '30');
    if (days === null) return;
    try {
      const data = await fetchJSON(API.logCleanup, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ days: parseInt(days) || 30 }),
      });
      showToast(`已清理 ${data.deleted || 0} 条旧日志`, 'success');
      this.refresh();
    } catch (err) { showToast(err.message, 'error'); }
  },
};

// ─── GithubTrending ──────────────────────────────────────────────────────────
const GithubTrending = {
  lang: 'all',
  since: 'daily',
  repos: [],
  inited: false,
  loading: false,

  LANGS: [
    { key: 'all', label: '全部语言' }, { key: 'python', label: 'Python' },
    { key: 'javascript', label: 'JavaScript' }, { key: 'typescript', label: 'TypeScript' },
    { key: 'go', label: 'Go' }, { key: 'rust', label: 'Rust' },
    { key: 'java', label: 'Java' }, { key: 'cpp', label: 'C++' },
    { key: 'csharp', label: 'C#' }, { key: 'shell', label: 'Shell' },
    { key: 'vue', label: 'Vue' },
  ],
  SINCE: [
    { key: 'daily', label: '今日新增' },
    { key: 'weekly', label: '本周新增' },
    { key: 'monthly', label: '本月新增' },
  ],

  init() {
    this.render();
    this.fetchRepos();
  },

  async fetchRepos() {
    if (this.loading) return;
    this.loading = true;
    this._setLoadingState(true);
    try {
      const data = await fetchJSON(`${API.githubTrending}?lang=${this.lang}&since=${this.since}`);
      this.repos = data.repos || [];
      this._renderCards(data);
    } catch (err) {
      const grid = $('#gh-card-grid');
      if (grid) grid.innerHTML = `<div class="empty-state" style="color:var(--danger);">❌ 加载失败: ${esc(err.message)}</div>`;
    } finally {
      this.loading = false;
      this._setLoadingState(false);
    }
  },

  _setLoadingState(on) {
    const btn = $('#gh-btn-refresh');
    if (btn) { btn.disabled = on; btn.textContent = on ? '↻ 加载中...' : '↻ 刷新'; }
  },

  render() {
    const container = $('#github-content');
    if (!container) return;
    const langTabs = this.LANGS.map(l =>
      `<button class="gh-tab ${this.lang === l.key ? 'active' : ''}" data-lang="${l.key}">${l.label}</button>`
    ).join('');
    const sinceTabs = this.SINCE.map(s =>
      `<button class="gh-since-tab ${this.since === s.key ? 'active' : ''}" data-since="${s.key}">${s.label}</button>`
    ).join('');
    container.innerHTML = `
      <div class="gh-toolbar">
        <div class="gh-lang-bar">${langTabs}</div>
        <div class="gh-right-bar">
          <div class="gh-since-bar">${sinceTabs}</div>
          <button class="btn-sm btn-secondary" id="gh-btn-refresh">↻ 刷新</button>
        </div>
      </div>
      <div id="gh-rate-info" class="gh-rate-info"></div>
      <div id="gh-card-grid" class="gh-card-grid"><div class="empty-state">⏳ 加载中...</div></div>
    `;
    container.querySelectorAll('.gh-tab').forEach(btn => {
      btn.addEventListener('click', () => { this.lang = btn.dataset.lang; this.render(); this.fetchRepos(); });
    });
    container.querySelectorAll('.gh-since-tab').forEach(btn => {
      btn.addEventListener('click', () => { this.since = btn.dataset.since; this.render(); this.fetchRepos(); });
    });
    $('#gh-btn-refresh')?.addEventListener('click', () => this.fetchRepos());
  },

  _renderCards(data) {
    const grid = $('#gh-card-grid');
    if (!grid) return;
    const rateInfo = $('#gh-rate-info');
    if (rateInfo && data.rate_remaining !== undefined && data.rate_remaining !== -1) {
      const pct = Math.round((data.rate_remaining / data.rate_limit) * 100);
      const color = pct > 30 ? '#10b981' : pct > 10 ? '#f59e0b' : '#ef4444';
      rateInfo.innerHTML = `<span style="color:${color};font-size:0.8rem;">⚡ GitHub API 配额: ${data.rate_remaining} / ${data.rate_limit}${data.cached ? ' &nbsp;📦 缓存中' : ''}</span>`;
    }
    if (!this.repos.length) { grid.innerHTML = '<div class="empty-state">💭 暂无数据，请稍后重试</div>'; return; }

    const LANG_COLORS = {
      Python:'#3572A5', JavaScript:'#f1e05a', TypeScript:'#3178c6', Go:'#00ADD8',
      Rust:'#dea584', Java:'#b07219', Vue:'#41b883', 'C++':'#f34b7d',
      'C#':'#178600', Shell:'#89e051', Unknown:'#8b949e',
    };
    grid.innerHTML = this.repos.map(repo => {
      const langColor = LANG_COLORS[repo.language] || '#8b949e';
      const starsStr = repo.stars >= 1000 ? (repo.stars/1000).toFixed(1)+'k' : repo.stars;
      const forksStr = repo.forks >= 1000 ? (repo.forks/1000).toFixed(1)+'k' : repo.forks;
      const topics = (repo.topics || []).map(t => `<span class="gh-topic">${esc(t)}</span>`).join('');
      return `
        <a class="gh-card" href="${esc(repo.url)}" target="_blank" rel="noopener">
          <div class="gh-card-header">
            <img class="gh-avatar" src="${esc(repo.avatar)}" alt="${esc(repo.owner)}" loading="lazy">
            <div class="gh-card-title">
              <div class="gh-owner">${esc(repo.owner)}</div>
              <div class="gh-repo">${esc(repo.name)}</div>
            </div>
          </div>
          <p class="gh-desc">${esc(repo.description || '暂无描述')}</p>
          ${topics ? `<div class="gh-topics">${topics}</div>` : ''}
          <div class="gh-meta">
            <span class="gh-lang" style="--lang-color:${langColor}"><span class="gh-lang-dot"></span>${esc(repo.language)}</span>
            <span class="gh-stat">⭐ ${starsStr}</span>
            <span class="gh-stat">🍴 ${forksStr}</span>
          </div>
        </a>`;
    }).join('');
  },
};

// ─── DBConnectionManager ──────────────────────────────────────────────────────
const DBConnectionManager = {
  connections: [],

  init() {
    this.refresh();
  },

  async refresh() {
    try {
      const data = await fetchJSON(API.dbConnections);
      this.connections = data.connections || [];
      this.render();
    } catch (err) {
      console.error('DB connections refresh failed', err);
      const container = $('#db-connections-content');
      if (container) container.innerHTML = `<div class="empty-state" style="color:var(--danger);">❌ 加载失败: ${esc(err.message)}</div>`;
    }
  },

  render() {
    const container = $('#db-connections-content');
    if (!container) return;

    const hasItems = this.connections.length > 0;
    container.innerHTML = `
      <div class="schedules-toolbar">
        <button class="btn-primary btn-sm" id="btn-create-db-conn">➕ 新建连接</button>
        <button class="btn-sm btn-secondary" id="btn-refresh-db-conn">🔄 刷新</button>
      </div>
      <div class="schedules-grid" id="db-conn-grid">
        ${!hasItems ? '<div class="empty-state">📭 暂无数据库连接，点击上方按钮创建</div>' : ''}
      </div>
    `;

    $('#btn-create-db-conn')?.addEventListener('click', () => this.showCreateModal());
    $('#btn-refresh-db-conn')?.addEventListener('click', () => this.refresh());

    if (hasItems) {
      const grid = $('#db-conn-grid');
      this.connections.forEach(c => {
        const card = document.createElement('div');
        card.className = 'schedule-card glass-panel';
        card.innerHTML = `
          <div class="schedule-card-header">
            <span class="schedule-icon">🔌</span>
            <div class="schedule-info">
              <strong>${esc(c.name)}</strong>
              <span class="schedule-cron">${esc(c.host)}:${c.port}/${c.database_name}</span>
            </div>
          </div>
          <div class="schedule-card-body">
            <div class="schedule-meta">
              <span>👤 ${esc(c.user)}</span>
              <span>💾 ${esc(c.db_type || 'mysql')}</span>
              <span>⏱️ ${c.timeout}s</span>
            </div>
          </div>
          <div class="schedule-card-actions">
            <button class="btn-sm btn-success" data-test="${c.id}">🔗 测试</button>
            <button class="btn-sm btn-secondary" data-edit="${c.id}">✏️ 编辑</button>
            <button class="btn-sm btn-danger" data-delete="${c.id}">🗑️ 删除</button>
          </div>
        `;

        card.querySelector('[data-test]')?.addEventListener('click', () => this.testConnection(c.id, card));
        card.querySelector('[data-edit]')?.addEventListener('click', () => this.showEditModal(c));
        card.querySelector('[data-delete]')?.addEventListener('click', () => this.deleteConnection(c.id));

        grid.appendChild(card);
      });
    }
  },

  showCreateModal() {
    this._showForm(null);
  },

  showEditModal(conn) {
    this._showForm(conn);
  },

  _showForm(conn) {
    const isEdit = !!conn;
    const modal = document.createElement('div');
    modal.className = 'modal-backdrop';
    modal.innerHTML = `
      <div class="modal-content wide">
        <h3>${isEdit ? '编辑' : '新建'}数据库连接</h3>
        <form id="form-db-conn" class="modern-form" style="margin-top:16px;">
          <input type="hidden" name="id" value="${isEdit ? esc(conn.id) : ''}">
          <div class="form-group">
            <label>连接名称 *</label>
            <input type="text" name="name" required placeholder="如：生产库、测试库" value="${isEdit ? esc(conn.name) : ''}">
          </div>
          <div class="form-group row" style="display:flex;gap:16px;">
            <div class="col" style="flex:2;">
              <label>主机 *</label>
              <input type="text" name="host" required value="${isEdit ? esc(conn.host) : '127.0.0.1'}" style="width:100%;background:rgba(0,0,0,0.15);border:1px solid var(--panel-border);color:var(--text-main);border-radius:8px;padding:8px;">
            </div>
            <div class="col" style="flex:1;">
              <label>端口 *</label>
              <input type="number" name="port" required value="${isEdit ? conn.port : 3306}" style="width:100%;background:rgba(0,0,0,0.15);border:1px solid var(--panel-border);color:var(--text-main);border-radius:8px;padding:8px;">
            </div>
          </div>
          <div class="form-group">
            <label>用户名 *</label>
            <input type="text" name="user" required value="${isEdit ? esc(conn.user) : ''}" style="width:100%;background:rgba(0,0,0,0.15);border:1px solid var(--panel-border);color:var(--text-main);border-radius:8px;padding:8px;">
          </div>
          <div class="form-group">
            <label>密码 *</label>
            <input type="password" name="password" required placeholder="${isEdit ? '留空则保留原密码' : '输入密码'}" value="" style="width:100%;background:rgba(0,0,0,0.15);border:1px solid var(--panel-border);color:var(--text-main);border-radius:8px;padding:8px;">
          </div>
          <div class="form-group">
            <label>数据库名 *</label>
            <input type="text" name="database_name" required value="${isEdit ? esc(conn.database_name) : ''}" style="width:100%;background:rgba(0,0,0,0.15);border:1px solid var(--panel-border);color:var(--text-main);border-radius:8px;padding:8px;">
          </div>
          <div class="form-group row" style="display:flex;gap:16px;">
            <div class="col" style="flex:2;">
              <label>字符集</label>
              <input type="text" name="charset" value="${isEdit ? esc(conn.charset || 'utf8mb4') : 'utf8mb4'}" style="width:100%;background:rgba(0,0,0,0.15);border:1px solid var(--panel-border);color:var(--text-main);border-radius:8px;padding:8px;">
            </div>
            <div class="col" style="flex:1;">
              <label>超时（秒）</label>
              <input type="number" name="timeout" value="${isEdit ? conn.timeout : 30}" min="1" style="width:100%;background:rgba(0,0,0,0.15);border:1px solid var(--panel-border);color:var(--text-main);border-radius:8px;padding:8px;">
            </div>
          </div>
          <div class="form-actions" style="display:flex;gap:12px;justify-content:flex-end;">
            <button type="button" class="btn-secondary" id="btn-cancel-db-conn">取消</button>
            <button type="submit" class="btn-primary" style="margin-top:0;">${isEdit ? '保存修改' : '创建'}</button>
          </div>
        </form>
      </div>
    `;

    document.body.appendChild(modal);
    modal.querySelector('#btn-cancel-db-conn').addEventListener('click', () => modal.remove());
    modal.addEventListener('click', (e) => { if (e.target === modal) modal.remove(); });

    modal.querySelector('#form-db-conn').addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const body = Object.fromEntries(fd.entries());

      // 编辑时如果密码为空，移除该字段让后端保留原值
      if (isEdit && !body.password) {
        delete body.password;
      }

      const btn = modal.querySelector('.btn-primary');
      btn.textContent = '⏳ 保存中...';
      btn.disabled = true;
      try {
        const method = isEdit ? 'POST' : 'POST';
        const url = isEdit ? API.dbConnections : API.dbConnections;
        await fetchJSON(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(body),
        });
        showToast(isEdit ? '连接已更新' : '连接已创建', 'success');
        modal.remove();
        this.refresh();
      } catch (err) {
        showToast(err.message, 'error');
        btn.textContent = isEdit ? '保存修改' : '创建';
        btn.disabled = false;
      }
    });
  },

  async testConnection(connId, card) {
    const btn = card.querySelector('[data-test]');
    const origText = btn.innerHTML;
    btn.innerHTML = '⏳ 测试中...';
    btn.disabled = true;
    try {
      const data = await fetchJSON(`${API.dbConnections}/${connId}/test`, { method: 'POST' });
      showToast(data.message || '连接成功', 'success');
    } catch (err) {
      showToast(err.message || '连接失败', 'error');
    } finally {
      btn.innerHTML = origText;
      btn.disabled = false;
    }
  },

  async deleteConnection(connId) {
    if (!confirm('确定要删除此连接吗？')) return;
    try {
      await fetchJSON(`${API.dbConnections}/${connId}`, { method: 'DELETE' });
      showToast('已删除', 'success');
      this.refresh();
    } catch (err) {
      showToast(err.message, 'error');
    }
  },
};

// ─── Initialize ─────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
  if (!$('#toast-container')) {
    const tc = document.createElement('div');
    tc.id = 'toast-container';
    document.body.appendChild(tc);
  }

  await buildSidebar();

  $('#clear-log')?.addEventListener('click', () => {
    consoleOut.innerHTML = '';
    appendLog('日志已清空。');
  });

  // Global input clear button delegation
  document.addEventListener('click', (e) => {
    const btn = e.target.closest('.btn-clear');
    if (btn) {
      const container = btn.closest('.input-with-clear');
      if (container) {
        const input = container.querySelector('input');
        if (input) { input.value = ''; input.focus(); input.dispatchEvent(new Event('input', { bubbles: true })); }
      }
    }
  });

  switchPanel('tasks');
  if (allTasks.length > 0) {
    TaskRunner.renderForm(allTasks[0]);
  }

  appendLog('✅ 数据员工 Dashboard 已就绪。');
});
