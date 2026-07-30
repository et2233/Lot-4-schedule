"""
项目管理 Web 应用 v2 - Flask 后端
- 日期全部清空，用户输入开始日期 → 自动计算预计完成日期 + 级联更新
- 实际完成日期 → 驱动状态自动变更
- 支持手动切换 IN PROGRESS / PLANNING
"""

import datetime
import os
import tempfile

from flask import Flask, jsonify, render_template_string, request, send_file

from engine import ProjectManager

app = Flask(__name__)

pm: ProjectManager | None = None
EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Construction_Schedule_with_GanttSheet.xlsx')


def get_pm() -> ProjectManager:
    global pm
    if pm is None:
        pm = ProjectManager(EXCEL_PATH).load()
    return pm


# ═══════════════════════════════════════════════════════════════
# API 路由
# ═══════════════════════════════════════════════════════════════

@app.route('/api/tasks')
def api_tasks():
    """获取所有任务列表"""
    try:
        m = get_pm()
        return jsonify({'tasks': m.get_task_list(), 'summary': m.get_summary(), 'error': None})
    except Exception as e:
        return jsonify({'tasks': [], 'summary': {}, 'error': str(e)})


@app.route('/api/update-start', methods=['POST'])
def api_update_start():
    """
    用户输入开始日期 → 自动计算预计完成并级联更新
    Body: {"task_id": 1, "new_start": "2026-04-20"}
    """
    try:
        data = request.get_json()
        task_id = data['task_id']
        new_start = data['new_start']
        m = get_pm()
        result = m.update_start(task_id, new_start)
        return jsonify(result)
    except Exception as e:
        return jsonify({'updated': {}, 'error': str(e)})


@app.route('/api/set-actual-finish', methods=['POST'])
def api_set_actual_finish():
    """
    设置实际完成日期
    Body: {"task_id": 1, "actual_finish": "2026-04-25"}  或  {"task_id": 1, "actual_finish": null}
    """
    try:
        data = request.get_json()
        task_id = data['task_id']
        actual_finish = data.get('actual_finish')  # None 表示清除
        m = get_pm()
        result = m.set_actual_finish(task_id, actual_finish)
        return jsonify(result)
    except Exception as e:
        return jsonify({'updated': {}, 'error': str(e)})


@app.route('/api/set-status', methods=['POST'])
def api_set_status():
    """
    手动修改任务状态
    Body: {"task_id": 1, "status": "IN PROGRESS"}
    """
    try:
        data = request.get_json()
        task_id = data['task_id']
        new_status = data['status']
        m = get_pm()
        result = m.set_status(task_id, new_status)
        return jsonify(result)
    except Exception as e:
        return jsonify({'updated': {}, 'error': str(e)})


@app.route('/api/update-deps', methods=['POST'])
def api_update_deps():
    """
    编辑前置任务和依赖类型
    Body: {"task_id": 1, "pred_ids": "2,5", "dep_specs": "FS+3,SS-1"}
    """
    try:
        data = request.get_json()
        task_id = data['task_id']
        pred_ids = data.get('pred_ids', '')
        dep_specs = data.get('dep_specs', '')
        m = get_pm()
        result = m.update_deps(task_id, pred_ids, dep_specs)
        return jsonify(result)
    except Exception as e:
        return jsonify({'updated': {}, 'error': str(e)})


@app.route('/api/delete-task', methods=['POST'])
def api_delete_task():
    """
    删除任务
    Body: {"task_id": 5}
    """
    try:
        data = request.get_json()
        task_id = data['task_id']
        m = get_pm()
        result = m.delete_task(task_id)
        return jsonify(result)
    except Exception as e:
        return jsonify({'updated': {}, 'error': str(e)})


@app.route('/api/export')
def api_export():
    """导出更新后的 Excel 文件"""
    try:
        m = get_pm()
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            output_path = tmp.name
        m.save(output_path)
        return send_file(
            output_path,
            as_attachment=True,
            download_name='Construction_Schedule_Updated.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/reload')
def api_reload():
    """重新加载（清空所有日期和修改，恢复初始状态）"""
    global pm
    try:
        pm = ProjectManager(EXCEL_PATH).reset()
        return jsonify({'status': 'ok', 'error': None})
    except Exception as e:
        return jsonify({'status': 'error', 'error': str(e)})


# ═══════════════════════════════════════════════════════════════
# Web 界面 v2
# ═══════════════════════════════════════════════════════════════

INDEX_HTML = r'''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>项目管理工具 - Construction Schedule</title>
<style>
  :root {
    --bg: #f0f2f5; --card-bg: #ffffff; --text: #1a1a2e; --text-secondary: #6b7280;
    --border: #e5e7eb; --primary: #2563eb; --primary-light: #eff6ff;
    --success: #059669; --warning: #d97706; --danger: #dc2626; --progress: #7c3aed;
    --shadow: 0 1px 3px rgba(0,0,0,0.06); --shadow-lg: 0 10px 25px rgba(0,0,0,0.1);
    --radius: 8px;
    --font: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", sans-serif;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: var(--font); background: var(--bg); color: var(--text); line-height: 1.5; }
  .header { background: var(--card-bg); border-bottom: 1px solid var(--border); padding: 14px 24px; display: flex; align-items: center; justify-content: space-between; box-shadow: var(--shadow); position: sticky; top: 0; z-index: 100; }
  .header-left h1 { font-size: 18px; font-weight: 700; }
  .header-left .subtitle { font-size: 12px; color: var(--text-secondary); margin-top: 2px; }
  .header-actions { display: flex; gap: 8px; }
  .btn { padding: 7px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 600; transition: all 0.15s; display: inline-flex; align-items: center; gap: 5px; white-space: nowrap; }
  .btn-primary { background: var(--primary); color: #fff; } .btn-primary:hover { background: #1d4ed8; }
  .btn-outline { background: #fff; color: var(--primary); border: 1.5px solid #d1d5db; } .btn-outline:hover { background: var(--primary-light); border-color: var(--primary); }
  .btn-danger { background: #fff; color: var(--danger); border: 1.5px solid #fca5a5; } .btn-danger:hover { background: #fef2f2; }
  .summary-bar { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; padding: 16px 24px; }
  .summary-card { background: var(--card-bg); border-radius: var(--radius); padding: 14px 18px; box-shadow: var(--shadow); border-left: 3px solid transparent; }
  .summary-card.card-tasks { border-left-color: var(--primary); } .summary-card.card-start { border-left-color: #0891b2; }
  .summary-card.card-finish { border-left-color: var(--success); } .summary-card.card-completed { border-left-color: var(--success); }
  .summary-card.card-progress { border-left-color: var(--progress); }
  .summary-card .label { font-size: 11px; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600; }
  .summary-card .value { font-size: 20px; font-weight: 700; margin-top: 2px; } .summary-card .value.small { font-size: 15px; }
  .main-content { padding: 0 24px 30px; }
  .table-container { background: var(--card-bg); border-radius: var(--radius); box-shadow: var(--shadow); overflow: hidden; }
  .table-toolbar { padding: 12px 16px; border-bottom: 1px solid var(--border); display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; }
  .table-toolbar h2 { font-size: 15px; font-weight: 600; } .toolbar-right { display: flex; gap: 8px; align-items: center; }
  .search-box { padding: 6px 12px; border: 1px solid var(--border); border-radius: 6px; font-size: 13px; width: 200px; outline: none; background: #f9fafb; }
  .search-box:focus { border-color: var(--primary); background: #fff; }
  .filter-status { padding: 6px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 12px; outline: none; background: #f9fafb; cursor: pointer; }
  .table-wrap { overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  thead { background: #f8fafc; }
  th { padding: 9px 10px; text-align: left; font-weight: 600; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; color: var(--text-secondary); border-bottom: 2px solid var(--border); white-space: nowrap; }
  td { padding: 6px 10px; border-bottom: 1px solid #f3f4f6; vertical-align: middle; }
  tr:hover td { background: #f8faff; } tr.cascaded td { background: #fffbeb; } tr.cascaded:hover td { background: #fef3c7; }
  .task-name { font-weight: 500; max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .task-name.placeholder { color: #999; font-style: italic; }
  .status-badge { display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: 11px; font-weight: 600; text-transform: uppercase; cursor: pointer; transition: all 0.15s; border: 1px solid transparent; }
  .status-badge:hover { transform: scale(1.05); }
  .status-COMPLETED { background: #d1fae5; color: #065f46; border-color: #a7f3d0; }
  .status-IN_PROGRESS, .status-IN\ PROGRESS { background: #ede9fe; color: #5b21b6; border-color: #c4b5fd; }
  .status-PLANNING { background: #dbeafe; color: #1e40af; border-color: #93c5fd; }
  .inline-input { border: 1px solid transparent; background: transparent; font-family: "SF Mono","Fira Code",monospace; font-size: 12px; padding: 3px 6px; border-radius: 4px; outline: none; transition: all 0.15s; }
  .inline-input:hover { border-color: #d1d5db; background: #fff; }
  .inline-input:focus { border-color: var(--primary); background: #fff; box-shadow: 0 0 0 2px rgba(37,99,235,0.1); }
  .inline-input.pred-ids { width: 70px; text-align: center; }
  .inline-input.dep-specs { width: 90px; text-align: center; }
  .del-btn { background: none; border: 1px solid #fecaca; color: #dc2626; border-radius: 4px; padding: 2px 7px; cursor: pointer; font-size: 11px; transition: all 0.15s; }
  .del-btn:hover { background: #fef2f2; border-color: #dc2626; }
  .inline-input.date-input { width: 115px; }
  .date-cell { font-family: "SF Mono","Fira Code",monospace; font-size: 12px; white-space: nowrap; }
  .date-cell.empty { color: #cbd5e1; font-style: italic; }
  .text-muted { color: #9ca3af; font-size: 12px; }
  .loading { text-align: center; padding: 40px; color: var(--text-secondary); }
  .cascade-alert { display: none; margin: 12px 0; padding: 12px 16px; border-radius: 8px; background: #fffbeb; border: 1px solid #fcd34d; font-size: 12px; }
  .cascade-alert.active { display: block; } .cascade-alert h4 { font-size: 13px; margin-bottom: 6px; color: #92400e; }
  .cascade-alert ul { padding-left: 16px; color: #78350f; max-height: 200px; overflow-y: auto; }
  .toast { position: fixed; top: 16px; right: 16px; z-index: 400; padding: 12px 20px; border-radius: 8px; color: #fff; font-weight: 600; font-size: 13px; transform: translateX(120%); transition: transform 0.25s ease; box-shadow: 0 8px 20px rgba(0,0,0,0.12); }
  .toast.show { transform: translateX(0); } .toast-success { background: #059669; } .toast-error { background: #dc2626; } .toast-info { background: #2563eb; }
  @media (max-width: 768px) { .header { flex-direction: column; gap: 8px; } .summary-bar { grid-template-columns: 1fr 1fr; padding: 10px 12px; } .main-content { padding: 0 8px 16px; } th, td { padding: 4px 6px; font-size: 11px; } }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>🏗️ 建筑施工进度管理</h1>
    <div class="subtitle" id="project-info">加载中...</div>
  </div>
  <div class="header-actions">
    <button class="btn btn-outline" onclick="exportExcel()">📥 导出 Excel</button>
    <button class="btn btn-danger" onclick="reloadData()">🔄 重置全部</button>
  </div>
</div>

<div class="summary-bar" id="summary-bar">
  <div class="summary-card card-tasks"><div class="label">任务总数</div><div class="value" id="sum-tasks">-</div></div>
  <div class="summary-card card-start"><div class="label">开工日期</div><div class="value small" id="sum-start">-</div></div>
  <div class="summary-card card-completed"><div class="label">已完成</div><div class="value" id="sum-completed" style="color:#059669">0</div></div>
  <div class="summary-card card-progress"><div class="label">进行中</div><div class="value" id="sum-progress" style="color:#7c3aed">0</div></div>
  <div class="summary-card card-finish"><div class="label">最晚预计完成</div><div class="value small" id="sum-late">-</div></div>
</div>

<div class="main-content">
  <div class="cascade-alert" id="cascade-alert">
    <h4>🔄 级联更新 — 以下任务已自动调整</h4>
    <ul id="cascade-list"></ul>
  </div>

  <div class="table-container">
    <div class="table-toolbar">
      <h2>📋 任务列表 <span class="text-muted">（点击单元格直接编辑）</span></h2>
      <div class="toolbar-right">
        <select class="filter-status" id="filter-status" onchange="filterTasks()">
          <option value="">全部状态</option>
          <option value="PLANNING">📘 PLANNING</option>
          <option value="IN PROGRESS">🟣 IN PROGRESS</option>
          <option value="COMPLETED">🟢 COMPLETED</option>
        </select>
        <input type="text" class="search-box" placeholder="🔍 搜索任务..." oninput="filterTasks()" id="search-input">
      </div>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>ID</th>
            <th>任务名称</th>
            <th>工期</th>
            <th>前置任务 ID</th>
            <th>依赖类型</th>
            <th>计划开始</th>
            <th>预计完成</th>
            <th>实际完成</th>
            <th>状态</th>
          </tr>
        </thead>
        <tbody id="task-table-body">
          <tr><td colspan="9" class="loading">加载中...</td></tr>
        </tbody>
      </table>
    </div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
let allTasks = [];
let cascadedTasks = new Set();

async function loadTasks() {
  try {
    const resp = await fetch('/api/tasks');
    const data = await resp.json();
    if (data.error) { showToast('加载失败: ' + data.error, 'error'); return; }
    allTasks = data.tasks;
    renderSummary(data.summary);
    renderTasks(allTasks);
    document.getElementById('project-info').textContent =
      (data.summary.project_address || 'Construction Schedule') +
      ' | 开工日期: ' + (data.summary.project_start || '—');
  } catch (e) { showToast('网络错误: ' + e.message, 'error'); }
}

function renderSummary(s) {
  document.getElementById('sum-tasks').textContent = s.total_tasks || '-';
  document.getElementById('sum-start').textContent = s.project_start || '-';
  document.getElementById('sum-completed').textContent = s.completed_count ?? '0';
  document.getElementById('sum-progress').textContent = s.in_progress_count ?? '0';
  document.getElementById('sum-late').textContent = s.latest_finish || '—';
}

function renderTasks(tasks) {
  const tbody = document.getElementById('task-table-body');
  if (!tasks.length) { tbody.innerHTML = '<tr><td colspan="9" class="loading">无匹配任务</td></tr>'; return; }
  tbody.innerHTML = tasks.map(t => {
    const isCascaded = cascadedTasks.has(t.task_id);
    const rowClass = isCascaded ? 'cascaded' : '';
    const statusDisplay = (t.status || 'PLANNING').replace(' ', '_');
    const statusClass = 'status-' + statusDisplay;
    const predIdsVal = t.pred_ids_str || '';
    const depSpecsVal = t.dep_specs_str || '';
    return `
    <tr class="${rowClass}" id="task-row-${t.task_id}">
      <td><strong>${t.task_id}</strong></td>
      <td><span class="task-name${!t.task_name?' placeholder':''}">${escHtml(t.task_name) || '(辅助节点)'}</span></td>
      <td>${t.duration}d</td>
      <td><input class="inline-input pred-ids" value="${predIdsVal}"
          onchange="onDepsChange(${t.task_id}, this.value, document.getElementById('depspec-${t.task_id}').value)"
          placeholder="如 1,2" title="输入前置任务ID，逗号分隔，如 1,5,6"></td>
      <td><input class="inline-input dep-specs" id="depspec-${t.task_id}" value="${depSpecsVal}"
          onchange="onDepsChange(${t.task_id}, document.getElementById('predid-${t.task_id}').value, this.value)"
          placeholder="FS" title="如 FS+3 或 SS-1，逗号分隔多个"></td>
      <td class="date-cell">
        <input class="inline-input date-input" type="date" value="${t.start||''}" onchange="onStartChange(${t.task_id}, this.value)" title="设置计划开始日期">
      </td>
      <td class="date-cell${!t.planned_finish?' empty':''}">${formatDateDMY(t.planned_finish)||'—'}</td>
      <td class="date-cell">
        <input class="inline-input date-input" type="date" value="${t.actual_finish||''}" onchange="onActualFinishChange(${t.task_id}, this.value)" title="设置实际完成日期">
      </td>
      <td>
        <span class="status-badge ${statusClass}" onclick="cycleStatus(${t.task_id}, '${t.status||'PLANNING'}')" title="点击切换 PLANNING ↔ IN PROGRESS ↔ COMPLETED">
          ${t.status||'PLANNING'}
        </span>
      </td>
    </tr>`;
  }).join('');
  // 给 pred-ids 的 input 加上 id（动态生成的，需要在渲染后重新设置）
  tasks.forEach(t => {
    const row = document.getElementById('task-row-' + t.task_id);
    if (row) {
      const predInput = row.querySelector('.pred-ids');
      if (predInput) predInput.id = 'predid-' + t.task_id;
    }
  });
}

function escHtml(s) { return (s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

function formatDateDMY(dateStr) {
  if (!dateStr) return '';
  // dateStr is YYYY-MM-DD
  const parts = dateStr.split('-');
  if (parts.length !== 3) return dateStr;
  return parts[2] + '/' + parts[1] + '/' + parts[0];
}

function filterTasks() {
  const q = document.getElementById('search-input').value.toLowerCase();
  const sf = document.getElementById('filter-status').value;
  let f = allTasks;
  if (q) f = f.filter(t => String(t.task_id).includes(q) || (t.task_name||'').toLowerCase().includes(q));
  if (sf) f = f.filter(t => (t.status||'PLANNING') === sf);
  renderTasks(f);
}

// ★ 编辑前置任务 + 依赖类型
async function onDepsChange(taskId, predIds, depSpecs) {
  try {
    const resp = await fetch('/api/update-deps', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({task_id: taskId, pred_ids: predIds, dep_specs: depSpecs})
    });
    const result = await resp.json();
    if (result.error) { showToast(result.error, 'error'); loadTasks(); return; }

    cascadedTasks.clear();
    for (const [tidStr, changes] of Object.entries(result.updated)) {
      const tid = parseInt(tidStr); cascadedTasks.add(tid);
      const task = allTasks.find(t => t.task_id === tid);
      if (task) {
        if (changes.new_pred_ids !== undefined) { task.pred_ids = changes.new_pred_ids; task.pred_ids_str = changes.new_pred_ids.join(','); }
        if (changes.new_pred_types !== undefined) task.pred_types = changes.new_pred_types;
        if (changes.new_lag_days !== undefined) { task.lag_days = changes.new_lag_days; task.dep_specs_str = changes.new_pred_types.map((pt,i) => pt + (changes.new_lag_days[i] ? (changes.new_lag_days[i]>0?'+':'')+changes.new_lag_days[i] : '')).join(','); }
        if (changes.new_start !== undefined) task.start = changes.new_start;
        if (changes.new_planned_finish !== undefined) task.planned_finish = changes.new_planned_finish;
      }
    }
    showCascadeSummary(result.updated);
    renderSummaryFromTasks();
    renderTasks(allTasks);
    showToast(`✅ 依赖已更新，${Object.keys(result.updated).length} 个任务已重算`, 'success');
    setTimeout(() => cascadedTasks.clear(), 5000);
  } catch (e) { showToast('请求失败: ' + e.message, 'error'); }
}

// 计划开始日期
async function onStartChange(taskId, newStart) {
  if (!newStart) return;
  try {
    const resp = await fetch('/api/update-start', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({task_id: taskId, new_start: newStart})
    });
    const result = await resp.json();
    if (result.error) { showToast(result.error, 'error'); loadTasks(); return; }
    cascadedTasks.clear();
    for (const [tidStr, changes] of Object.entries(result.updated)) {
      const tid = parseInt(tidStr); cascadedTasks.add(tid);
      const task = allTasks.find(t => t.task_id === tid);
      if (task) {
        if (changes.new_start !== undefined) task.start = changes.new_start;
        if (changes.new_planned_finish !== undefined) task.planned_finish = changes.new_planned_finish;
        if (changes.new_status !== undefined) task.status = changes.new_status;
      }
    }
    showCascadeSummary(result.updated); renderSummaryFromTasks(); renderTasks(allTasks);
    showToast(`✅ 已更新 ${Object.keys(result.updated).length} 个任务`, 'success');
    setTimeout(() => cascadedTasks.clear(), 5000);
  } catch (e) { showToast('请求失败: ' + e.message, 'error'); }
}

// 实际完成日期
async function onActualFinishChange(taskId, newActual) {
  try {
    const resp = await fetch('/api/set-actual-finish', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({task_id: taskId, actual_finish: newActual || null})
    });
    const result = await resp.json();
    if (result.error) { showToast(result.error, 'error'); loadTasks(); return; }
    for (const [tidStr, changes] of Object.entries(result.updated)) {
      const tid = parseInt(tidStr); const task = allTasks.find(t => t.task_id === tid);
      if (task) {
        if (changes.new_actual_finish !== undefined) task.actual_finish = changes.new_actual_finish;
        if (changes.new_status !== undefined) task.status = changes.new_status;
      }
    }
    renderSummaryFromTasks(); renderTasks(allTasks);
    showToast(newActual ? '✅ 实际完成日期已设置' : '🔄 实际完成日期已清除', 'success');
  } catch (e) { showToast('请求失败: ' + e.message, 'error'); }
}

// 状态切换 (三向: PLANNING → IN PROGRESS → COMPLETED)
async function cycleStatus(taskId, currentStatus) {
  const task = allTasks.find(t => t.task_id === taskId);
  if (!task) return;
  const statuses = ['PLANNING', 'IN PROGRESS', 'COMPLETED'];
  const idx = statuses.indexOf(currentStatus);
  const nextStatus = statuses[(idx + 1) % statuses.length];
  try {
    const resp = await fetch('/api/set-status', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({task_id: taskId, status: nextStatus})
    });
    const result = await resp.json();
    if (result.error) { showToast(result.error, 'error'); return; }
    for (const [tidStr, changes] of Object.entries(result.updated)) {
      const tid = parseInt(tidStr); const t = allTasks.find(tt => tt.task_id === tid);
      if (t && changes.new_status !== undefined) t.status = changes.new_status;
    }
    renderSummaryFromTasks(); renderTasks(allTasks);
    showToast(`状态已切换为 ${nextStatus}`, 'info');
  } catch (e) { showToast('请求失败: ' + e.message, 'error'); }
}

function renderSummaryFromTasks() {
  document.getElementById('sum-completed').textContent = allTasks.filter(t => t.status === 'COMPLETED').length;
  document.getElementById('sum-progress').textContent = allTasks.filter(t => t.status === 'IN PROGRESS').length;
}

function showCascadeSummary(updated) {
  const container = document.getElementById('cascade-alert');
  const list = document.getElementById('cascade-list'); list.innerHTML = '';
  for (const [tidStr, changes] of Object.entries(updated)) {
    const tid = parseInt(tidStr); const task = allTasks.find(t => t.task_id === tid);
    const name = task ? task.task_name : `Task ${tid}`; const parts = [];
    if (changes.new_pred_ids !== undefined) parts.push(`前置: [${changes.old_pred_ids}] → [${changes.new_pred_ids}]`);
    if (changes.new_start !== undefined) parts.push(`开始: ${changes.old_start||'—'} → ${changes.new_start}`);
    if (changes.new_planned_finish !== undefined) parts.push(`预计完成: ${changes.old_planned_finish||'—'} → ${changes.new_planned_finish}`);
    list.innerHTML += `<li><strong>ID ${tid} ${escHtml(name)}</strong>: ${parts.join(', ')}</li>`;
  }
  container.classList.add('active'); setTimeout(() => container.classList.remove('active'), 8000);
}

function exportExcel() { window.open('/api/export', '_blank'); }

async function deleteTask(taskId) {
  const task = allTasks.find(t => t.task_id === taskId);
  const name = task ? task.task_name : `ID ${taskId}`;
  if (!confirm(`确定要删除任务 "${name}" (ID: ${taskId}) 吗？\n\n其后继任务将自动重新连接到该任务的前置任务。`)) return;
  try {
    const resp = await fetch('/api/delete-task', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({task_id: taskId})
    });
    const result = await resp.json();
    if (result.error) { showToast(result.error, 'error'); return; }
    // 重新加载
    await loadTasks();
    showToast(`✅ 已删除任务 "${name}"`, 'success');
  } catch (e) { showToast('请求失败: ' + e.message, 'error'); }
}

async function reloadData() {
  if (!confirm('⚠️ 这将清空所有日期、状态和依赖修改，恢复为初始状态。确定继续？')) return;
  try {
    await fetch('/api/reload'); cascadedTasks.clear();
    document.getElementById('cascade-alert').classList.remove('active');
    await loadTasks(); showToast('🔄 已重置', 'info');
  } catch (e) { showToast('重载失败: ' + e.message, 'error'); }
}

function showToast(msg, type) {
  const toast = document.getElementById('toast');
  toast.textContent = msg; toast.className = 'toast toast-' + type + ' show';
  setTimeout(() => toast.classList.remove('show'), 3000);
}

loadTasks();
</script>
</body>
</html>
'''


@app.route('/')
def index():
    return render_template_string(INDEX_HTML)


if __name__ == '__main__':
    print("🚀 项目管理工具 v2 启动中...")
    print(f"📂 Excel 文件: {EXCEL_PATH}")
    print(f"📁 Excel 存在: {os.path.exists(EXCEL_PATH)}")
    print(f"🌐 访问地址: http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)
