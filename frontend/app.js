const API_BASE = (window.BIOTECH_PM_API_BASE || "").replace(/\/$/, "");
const STAGES = ["样本入库", "样本核对", "建库", "测序", "数据存放", "生信分析", "合同资料", "回款", "项目结束"];

const state = {
  token: localStorage.getItem("biotech_pm_token") || "",
  user: null,
  authMode: "login",
  view: "projects",
  projects: [],
  project: null,
  samples: [],
  tasks: [],
  files: [],
  payment: null,
  activeStage: "样本入库",
  editingSampleId: null,
  showProjectForm: false,
  search: "",
};

const app = document.querySelector("#app");
const toast = document.querySelector("#toast");

document.addEventListener("click", handleClick);
document.addEventListener("submit", handleSubmit);
document.addEventListener("input", handleInput);

boot();

async function boot() {
  if (!state.token) {
    renderAuth();
    return;
  }
  try {
    const data = await api("/api/me");
    state.user = data.user;
    await loadProjects();
    renderProjects();
  } catch (_error) {
    logout();
  }
}

async function api(path, options = {}) {
  const headers = { ...(options.headers || {}) };
  if (state.token) headers.Authorization = `Bearer ${state.token}`;

  let body = options.body;
  if (body && !(body instanceof FormData)) {
    headers["Content-Type"] = "application/json";
    body = JSON.stringify(body);
  }

  const response = await fetch(`${API_BASE}${path}`, { ...options, headers, body });
  const contentType = response.headers.get("Content-Type") || "";
  const data = contentType.includes("application/json") ? await response.json() : {};
  if (!response.ok || data.ok === false) {
    throw new Error(data.error || data.message || `请求失败：${response.status}`);
  }
  return data;
}

async function downloadFile(path, filename) {
  const headers = {};
  if (state.token) headers.Authorization = `Bearer ${state.token}`;
  const response = await fetch(`${API_BASE}${path}`, { headers });
  if (!response.ok) {
    throw new Error("下载失败");
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function renderAuth() {
  const isLogin = state.authMode === "login";
  const isRegister = state.authMode === "register";
  app.innerHTML = `
    <main class="auth-shell">
      <section class="auth-panel">
        <div class="brand">
          <div class="brand-mark">PM</div>
          <h2>${isLogin ? "账号登录" : isRegister ? "申请账号" : "忘记密码"}</h2>
          <p>实验项目管理系统</p>
        </div>
        ${isLogin ? renderLoginForm() : isRegister ? renderRegisterForm() : renderForgotForm()}
      </section>
      <section class="auth-side">
        <h1>项目、样本、实验、生信、合同和回款集中管理</h1>
        <p>围绕样本主数据追踪项目进度，保留 Excel 导入导出习惯，同时沉淀可检索、可追责的数据记录。</p>
      </section>
    </main>
  `;
}

function renderLoginForm() {
  return `
    <form class="form" data-form="login">
      <label class="field"><span>账号</span><input name="username" autocomplete="username" required /></label>
      <label class="field"><span>密码</span><input name="password" type="password" autocomplete="current-password" required /></label>
      <button class="primary" type="submit">登录</button>
      <div class="auth-links">
        <button class="link-button" type="button" data-action="auth-mode" data-mode="register">申请账号</button>
        <button class="link-button" type="button" data-action="auth-mode" data-mode="forgot">忘记密码</button>
      </div>
    </form>
  `;
}

function renderRegisterForm() {
  return `
    <form class="form" data-form="register">
      <label class="field"><span>账号</span><input name="username" autocomplete="username" required /></label>
      <label class="field"><span>姓名</span><input name="display_name" required /></label>
      <label class="field"><span>密码</span><input name="password" type="password" minlength="6" required /></label>
      <button class="primary" type="submit">提交申请</button>
      <button class="link-button" type="button" data-action="auth-mode" data-mode="login">返回登录</button>
    </form>
  `;
}

function renderForgotForm() {
  return `
    <form class="form" data-form="forgot">
      <label class="field"><span>账号</span><input name="username" autocomplete="username" required /></label>
      <button class="primary" type="submit">提交重置申请</button>
      <button class="link-button" type="button" data-action="auth-mode" data-mode="login">返回登录</button>
    </form>
  `;
}

function renderTopbar(title, actions = "") {
  return `
    <header class="topbar">
      <div class="topbar-title">
        <div class="brand-mark" style="width:30px;height:30px;margin:0">PM</div>
        <h1>${h(title)}</h1>
      </div>
      <div class="user-box">
        <span>${h(state.user?.display_name || state.user?.username || "")}</span>
        ${actions}
        <button class="ghost" type="button" data-action="logout">退出</button>
      </div>
    </header>
  `;
}

function renderProjects() {
  const rows = filteredProjects();
  app.innerHTML = `
    <div class="app-shell">
      ${renderTopbar("项目管理系统", `<button class="secondary" type="button" data-action="export-summary">导出汇总</button>`)}
      <main class="content">
        <div class="section-head">
          <div>
            <h2>项目列表</h2>
            <p>${state.projects.length} 个项目</p>
          </div>
          <div class="toolbar">
            <input class="compact-input" data-action="search-projects" value="${attr(state.search)}" placeholder="搜索项目" />
            <button class="primary" type="button" data-action="toggle-project-form">${state.showProjectForm ? "收起创建" : "创建项目"}</button>
          </div>
        </div>
        ${state.showProjectForm ? renderProjectForm() : ""}
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>项目名称</th>
                <th>优先级</th>
                <th>进度</th>
                <th>科研经理</th>
                <th>样本量</th>
                <th>回款</th>
                <th>更新时间</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              ${rows.map(renderProjectRow).join("") || `<tr><td colspan="8"><div class="empty">暂无项目</div></td></tr>`}
            </tbody>
          </table>
        </div>
      </main>
    </div>
  `;
}

function renderProjectForm() {
  return `
    <section class="panel" style="margin-bottom:16px">
      <div class="panel-body">
        <form class="grid-form" data-form="project">
          <label class="field wide"><span>项目名称</span><input name="name" /></label>
          <label class="field"><span>OA立项编号</span><input name="oa_project_code" /></label>
          <label class="field"><span>订单编号</span><input name="order_code" /></label>
          <label class="field"><span>医院名称</span><input name="hospital_name" /></label>
          <label class="field"><span>客户名称</span><input name="customer_name" /></label>
          <label class="field"><span>科研经理</span><input name="research_manager" /></label>
          <label class="field"><span>优先级</span>
            <select name="priority">
              <option>常规</option>
              <option>重要</option>
              <option>紧急</option>
            </select>
          </label>
          <label class="field"><span>样本类型</span><input name="sample_type" /></label>
          <label class="field"><span>到样时间</span><input name="arrival_date" type="date" /></label>
          <label class="field wide"><span>具体实验室流程</span><input name="lab_flow" /></label>
          <label class="field wide"><span>具体操作补充</span><input name="operation_note" /></label>
          <div class="toolbar full">
            <button class="primary" type="submit">保存项目</button>
            <button class="ghost" type="button" data-action="toggle-project-form">取消</button>
          </div>
        </form>
      </div>
    </section>
  `;
}

function renderProjectRow(project) {
  return `
    <tr>
      <td>
        <strong>${h(project.name)}</strong><br />
        <span class="file-name">${h(project.oa_project_code || project.order_code || "")}</span>
      </td>
      <td><span class="priority ${priorityClass(project.priority)}">${h(project.priority || "常规")}</span></td>
      <td>${renderProgress(project.progress_percent, project.current_progress)}</td>
      <td>${h(project.research_manager || project.owner_name || "")}</td>
      <td>${h(project.sample_count || project.received_sample_count || 0)}</td>
      <td>${h(project.received_amount || 0)} / ${h(project.total_amount || 0)}</td>
      <td>${h(project.updated_at || "")}</td>
      <td class="row-action"><button class="secondary" type="button" data-action="open-project" data-id="${project.id}">进入</button></td>
    </tr>
  `;
}

function renderDetail() {
  const project = state.project;
  if (!project) return renderProjects();
  app.innerHTML = `
    <div class="app-shell">
      ${renderTopbar(project.name, `
        <button class="ghost" type="button" data-action="back-projects">项目列表</button>
        <button class="secondary" type="button" data-action="export-summary">导出汇总</button>
      `)}
      <div class="detail-layout">
        <aside class="sidebar">
          <div class="project-card">
            <h2>${h(project.name)}</h2>
            ${renderProgress(project.progress_percent, project.current_progress)}
            <div class="meta-list" style="margin-top:12px">
              <span>${h(project.hospital_name || "未填医院")}</span>
              <span>${h(project.customer_name || "未填客户")}</span>
              <span>${h(project.oa_project_code || project.order_code || "未填编号")}</span>
            </div>
          </div>
          <nav class="stage-nav">
            ${STAGES.map(stage => `
              <button class="stage-button ${stage === state.activeStage ? "active" : ""}" type="button" data-action="stage" data-stage="${stage}">
                ${stage}
              </button>
            `).join("")}
          </nav>
        </aside>
        <main class="detail-main">
          ${renderModule()}
        </main>
      </div>
    </div>
  `;
}

function renderModule() {
  const stage = state.activeStage;
  if (stage === "样本入库") return renderSampleImport();
  if (stage === "样本核对") return renderSampleCheck();
  if (stage === "建库") return renderLibrary();
  if (stage === "测序") return renderSequencing();
  if (stage === "数据存放") return renderStorage();
  if (stage === "生信分析") return renderBioinfo();
  if (stage === "合同资料") return renderContracts();
  if (stage === "回款") return renderPayment();
  return renderFinish();
}

function renderModuleHead(title, actions = "") {
  return `
    <div class="module-head">
      <div>
        <h2>${h(title)}</h2>
        <p>${state.samples.length} 个样本</p>
      </div>
      <div class="toolbar">${actions}</div>
    </div>
  `;
}

function renderSampleImport() {
  return `
    ${renderModuleHead("样本入库")}
    <div class="stack">
      ${renderMetrics()}
      <section class="panel">
        <div class="panel-body">
          <form class="form" data-form="sample-import">
            <label class="field"><span>1.样本信息表.xlsx</span><input name="file" type="file" accept=".xlsx" required /></label>
            <button class="primary" type="submit">导入样本</button>
          </form>
        </div>
      </section>
      ${renderSampleMiniTable(["sample_code", "patient_name", "sample_category", "sample_subcategory", "group_name", "sample_status"])}
    </div>
  `;
}

function renderSampleCheck() {
  return `
    ${renderModuleHead("样本核对", `<button class="ghost" type="button" data-action="new-sample">新增样本</button>`)}
    <div class="split">
      <section class="panel"><div class="panel-body">${renderSampleForm()}</div></section>
      <section>${renderSampleCheckTable()}</section>
    </div>
  `;
}

function renderSampleForm() {
  const sample = state.samples.find(item => item.id === state.editingSampleId) || {};
  return `
    <form class="form" data-form="sample">
      <input type="hidden" name="id" value="${attr(sample.id || "")}" />
      <h3 class="panel-title">${sample.id ? "编辑样本" : "新增样本"}</h3>
      <label class="field"><span>样本管编号</span><input name="sample_code" value="${attr(sample.sample_code)}" required /></label>
      <label class="field"><span>患者姓名</span><input name="patient_name" value="${attr(sample.patient_name)}" /></label>
      <label class="field"><span>样本大类</span><input name="sample_category" value="${attr(sample.sample_category)}" /></label>
      <label class="field"><span>样本小类</span><input name="sample_subcategory" value="${attr(sample.sample_subcategory)}" /></label>
      <label class="field"><span>分组名</span><input name="group_name" value="${attr(sample.group_name)}" /></label>
      <label class="field"><span>体积</span><input name="volume" value="${attr(sample.volume)}" /></label>
      <label class="field"><span>检测项目</span><input name="test_project" value="${attr(sample.test_project)}" /></label>
      <label class="field"><span>状态</span>
        <select name="sample_status">
          ${["待核对", "正常", "信息缺失", "样本异常", "已建库", "已下机", "已完成"].map(value => `<option ${value === sample.sample_status ? "selected" : ""}>${value}</option>`).join("")}
        </select>
      </label>
      <label class="field"><span>备注</span><textarea name="sample_note">${h(sample.sample_note)}</textarea></label>
      <div class="toolbar">
        <button class="primary" type="submit">${sample.id ? "保存修改" : "新增样本"}</button>
        ${sample.id ? `<button class="ghost" type="button" data-action="cancel-sample-edit">取消编辑</button>` : ""}
      </div>
    </form>
  `;
}

function renderSampleCheckTable() {
  if (!state.samples.length) return `<div class="panel"><div class="empty">暂无样本</div></div>`;
  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>样本管编号</th>
            <th>患者姓名</th>
            <th>样本类型</th>
            <th>分组</th>
            <th>检测项目</th>
            <th>状态</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          ${state.samples.map(sample => `
            <tr>
              <td>${h(sample.sample_code)}</td>
              <td>${h(sample.patient_name)}</td>
              <td>${h([sample.sample_category, sample.sample_subcategory].filter(Boolean).join(" / "))}</td>
              <td>${h(sample.group_name)}</td>
              <td>${h(sample.test_project)}</td>
              <td><span class="status">${h(sample.sample_status || "待核对")}</span></td>
              <td class="row-action">
                <button class="secondary" type="button" data-action="edit-sample" data-id="${sample.id}">编辑</button>
                <button class="danger" type="button" data-action="delete-sample" data-id="${sample.id}">删除</button>
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderLibrary() {
  return `
    ${renderModuleHead("建库", "")}
    <div class="stack">
      <section class="panel">
        <div class="panel-body">
          <form class="form" data-form="library-import">
            <label class="field"><span>2.实验室上机.xlsx</span><input name="file" type="file" accept=".xlsx" required /></label>
            <button class="primary" type="submit">导入建库信息</button>
          </form>
        </div>
      </section>
      ${renderSampleMiniTable(["sample_code", "lab_code", "sub_library_code", "adapter_code", "extraction_concentration", "library_concentration", "operator_name"])}
    </div>
  `;
}

function renderSequencing() {
  return `
    ${renderModuleHead("测序")}
    ${renderEditableTable(
      ["sample_code", "sub_library_code", "total_reads", "host_rate", "q20", "q30", "report_result", "detected_pathogens"],
      {
        sample_code: "样本编号",
        sub_library_code: "子文库编号",
        total_reads: "total reads",
        host_rate: "Host_rate",
        q20: "Q20",
        q30: "Q30",
        report_result: "报告结果",
        detected_pathogens: "检出病原菌",
      }
    )}
  `;
}

function renderStorage() {
  return `
    ${renderModuleHead("数据存放")}
    ${renderEditableTable(
      ["sample_code", "lab_code", "sub_library_code", "data_storage_path", "storage_notes"],
      {
        sample_code: "样本编号",
        lab_code: "实验室编号",
        sub_library_code: "子文库编号",
        data_storage_path: "网盘存放路径",
        storage_notes: "备注",
      }
    )}
  `;
}

function renderBioinfo() {
  return `
    ${renderModuleHead("生信分析", `<button class="secondary" type="button" data-action="export-bioinfo">导出生信信息表</button>`)}
    <div class="split">
      <section class="panel">
        <div class="panel-body">
          <form class="form" data-form="bioinfo-task">
            <h3 class="panel-title">生信任务</h3>
            <label class="field"><span>任务名称</span><input name="title" required /></label>
            <label class="field"><span>负责人</span><input name="owner" /></label>
            <label class="field"><span>截止日期</span><input name="due_date" type="date" /></label>
            <label class="field"><span>状态</span>
              <select name="status"><option>待开始</option><option>进行中</option><option>已完成</option></select>
            </label>
            <label class="field"><span>任务内容</span><textarea name="content"></textarea></label>
            <button class="primary" type="submit">新增任务</button>
          </form>
        </div>
      </section>
      <section>${renderTaskTable()}</section>
    </div>
  `;
}

function renderTaskTable() {
  if (!state.tasks.length) return `<div class="panel"><div class="empty">暂无生信任务</div></div>`;
  return `
    <div class="table-wrap">
      <table>
        <thead><tr><th>任务</th><th>负责人</th><th>状态</th><th>截止日期</th><th>内容</th></tr></thead>
        <tbody>
          ${state.tasks.map(task => `
            <tr>
              <td>${h(task.title)}</td>
              <td>${h(task.owner)}</td>
              <td><span class="status">${h(task.status)}</span></td>
              <td>${h(task.due_date)}</td>
              <td>${h(task.content)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderContracts() {
  const contractFiles = state.files.filter(file => file.module === "contract");
  return `
    ${renderModuleHead("合同资料")}
    <div class="split">
      <section class="panel">
        <div class="panel-body">
          <form class="form" data-form="contract-upload">
            <label class="field"><span>合同文件</span><input name="file" type="file" required /></label>
            <button class="primary" type="submit">上传合同</button>
          </form>
        </div>
      </section>
      <section>${renderFileTable(contractFiles)}</section>
    </div>
  `;
}

function renderFileTable(files) {
  if (!files.length) return `<div class="panel"><div class="empty">暂无文件</div></div>`;
  return `
    <div class="table-wrap">
      <table>
        <thead><tr><th>文件名</th><th>模块</th><th>上传时间</th><th></th></tr></thead>
        <tbody>
          ${files.map(file => `
            <tr>
              <td>${h(file.original_name)}</td>
              <td>${h(file.module)}</td>
              <td>${h(file.created_at)}</td>
              <td class="row-action"><button class="secondary" type="button" data-action="download-file" data-url="${attr(file.url)}" data-name="${attr(file.original_name)}">下载</button></td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderPayment() {
  const payment = state.payment || {};
  const percent = payment.progress_percent || 0;
  return `
    ${renderModuleHead("回款")}
    <div class="split">
      <section class="panel">
        <div class="panel-body">
          <form class="form" data-form="payment">
            <label class="field"><span>总金额</span><input name="total_amount" type="number" min="0" step="0.01" value="${attr(payment.total_amount || 0)}" /></label>
            <label class="field"><span>到账金额</span><input name="received_amount" type="number" min="0" step="0.01" value="${attr(payment.received_amount || 0)}" /></label>
            <label class="field"><span>备注</span><textarea name="payment_note">${h(payment.payment_note)}</textarea></label>
            <button class="primary" type="submit">保存回款</button>
          </form>
        </div>
      </section>
      <section class="panel">
        <div class="panel-body">
          <div class="metric-grid">
            <div class="metric"><b>${h(payment.total_amount || 0)}</b><span>总金额</span></div>
            <div class="metric"><b>${h(payment.received_amount || 0)}</b><span>到账金额</span></div>
            <div class="metric"><b>${percent}%</b><span>回款进度</span></div>
            <div class="metric"><b>${Math.max(0, Number(payment.total_amount || 0) - Number(payment.received_amount || 0)).toFixed(2)}</b><span>未回款</span></div>
          </div>
        </div>
      </section>
    </div>
  `;
}

function renderFinish() {
  const completed = state.project?.status === "completed";
  return `
    ${renderModuleHead("项目结束")}
    <section class="panel">
      <div class="panel-body stack">
        ${renderMetrics()}
        <div class="toolbar">
          <button class="${completed ? "ghost" : "primary"}" type="button" data-action="complete-project" ${completed ? "disabled" : ""}>
            ${completed ? "项目已结束" : "标记项目结束"}
          </button>
        </div>
      </div>
    </section>
  `;
}

function renderMetrics() {
  const libraryCount = state.samples.filter(item => item.sub_library_code).length;
  const storageCount = state.samples.filter(item => item.data_storage_path).length;
  const completedCount = state.samples.filter(item => item.sample_status === "已完成").length;
  return `
    <div class="metric-grid">
      <div class="metric"><b>${state.samples.length}</b><span>样本数</span></div>
      <div class="metric"><b>${libraryCount}</b><span>建库记录</span></div>
      <div class="metric"><b>${storageCount}</b><span>数据存放</span></div>
      <div class="metric"><b>${completedCount}</b><span>已完成样本</span></div>
    </div>
  `;
}

function renderSampleMiniTable(fields) {
  if (!state.samples.length) return `<div class="panel"><div class="empty">暂无样本</div></div>`;
  const labels = {
    sample_code: "样本编号",
    patient_name: "患者姓名",
    sample_category: "样本大类",
    sample_subcategory: "样本小类",
    group_name: "分组",
    sample_status: "状态",
    lab_code: "实验室编号",
    sub_library_code: "子文库编号",
    adapter_code: "接头编号",
    extraction_concentration: "提取浓度",
    library_concentration: "文库浓度",
    operator_name: "操作人",
  };
  return `
    <div class="table-wrap">
      <table>
        <thead><tr>${fields.map(field => `<th>${h(labels[field] || field)}</th>`).join("")}</tr></thead>
        <tbody>
          ${state.samples.map(sample => `
            <tr>${fields.map(field => `<td>${h(sample[field])}</td>`).join("")}</tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderEditableTable(fields, labels) {
  if (!state.samples.length) return `<div class="panel"><div class="empty">暂无样本</div></div>`;
  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>${fields.map(field => `<th>${h(labels[field] || field)}</th>`).join("")}<th></th></tr>
        </thead>
        <tbody>
          ${state.samples.map(sample => `
            <tr data-sample-id="${sample.id}">
              ${fields.map(field => {
                if (["sample_code", "lab_code", "sub_library_code"].includes(field)) return `<td>${h(sample[field])}</td>`;
                return `<td><input class="compact-input" name="${field}" value="${attr(sample[field])}" /></td>`;
              }).join("")}
              <td class="row-action"><button class="secondary" type="button" data-action="save-row" data-id="${sample.id}">保存</button></td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

async function handleClick(event) {
  const button = event.target.closest("[data-action]");
  if (!button) return;
  const action = button.dataset.action;

  try {
    if (action === "auth-mode") {
      state.authMode = button.dataset.mode;
      renderAuth();
    }
    if (action === "logout") logout();
    if (action === "toggle-project-form") {
      state.showProjectForm = !state.showProjectForm;
      renderProjects();
    }
    if (action === "open-project") await openProject(Number(button.dataset.id));
    if (action === "back-projects") {
      state.view = "projects";
      await loadProjects();
      renderProjects();
    }
    if (action === "stage") await setStage(button.dataset.stage);
    if (action === "export-summary") await downloadFile("/api/projects/export-summary", "项目信息汇总.xlsx");
    if (action === "export-bioinfo") await downloadFile(`/api/projects/${state.project.id}/bioinfo/export`, "生信信息表.xlsx");
    if (action === "new-sample") {
      state.editingSampleId = null;
      renderDetail();
    }
    if (action === "edit-sample") {
      state.editingSampleId = Number(button.dataset.id);
      renderDetail();
    }
    if (action === "cancel-sample-edit") {
      state.editingSampleId = null;
      renderDetail();
    }
    if (action === "delete-sample") await deleteSample(Number(button.dataset.id));
    if (action === "save-row") await saveEditableRow(button);
    if (action === "download-file") await downloadFile(button.dataset.url, button.dataset.name);
    if (action === "complete-project") await completeProject();
  } catch (error) {
    showToast(error.message, true);
  }
}

async function handleSubmit(event) {
  const form = event.target.closest("form[data-form]");
  if (!form) return;
  event.preventDefault();
  const name = form.dataset.form;
  try {
    if (name === "login") await submitLogin(form);
    if (name === "register") await submitRegister(form);
    if (name === "forgot") await submitForgot(form);
    if (name === "project") await submitProject(form);
    if (name === "sample") await submitSample(form);
    if (name === "sample-import") await submitFileForm(form, `/api/projects/${state.project.id}/samples/import`, "样本导入完成");
    if (name === "library-import") await submitFileForm(form, `/api/projects/${state.project.id}/library/import`, "建库信息导入完成");
    if (name === "contract-upload") await submitFileForm(form, `/api/projects/${state.project.id}/files/contract`, "合同上传完成");
    if (name === "bioinfo-task") await submitBioinfoTask(form);
    if (name === "payment") await submitPayment(form);
  } catch (error) {
    showToast(error.message, true);
  }
}

function handleInput(event) {
  const input = event.target.closest("[data-action='search-projects']");
  if (!input) return;
  state.search = input.value;
  renderProjects();
  const nextInput = document.querySelector("[data-action='search-projects']");
  if (nextInput) {
    nextInput.focus();
    nextInput.setSelectionRange(nextInput.value.length, nextInput.value.length);
  }
}

async function submitLogin(form) {
  const data = formData(form);
  const result = await api("/api/auth/login", { method: "POST", body: data });
  state.token = result.token;
  state.user = result.user;
  localStorage.setItem("biotech_pm_token", state.token);
  await loadProjects();
  renderProjects();
}

async function submitRegister(form) {
  await api("/api/auth/register-request", { method: "POST", body: formData(form) });
  showToast("申请已提交");
  state.authMode = "login";
  renderAuth();
}

async function submitForgot(form) {
  const result = await api("/api/auth/forgot-password", { method: "POST", body: formData(form) });
  showToast(result.message || "已提交");
  state.authMode = "login";
  renderAuth();
}

async function submitProject(form) {
  await api("/api/projects", { method: "POST", body: formData(form) });
  state.showProjectForm = false;
  await loadProjects();
  renderProjects();
  showToast("项目已创建");
}

async function submitSample(form) {
  const data = formData(form);
  const id = Number(data.id);
  delete data.id;
  if (id) {
    await api(`/api/samples/${id}`, { method: "PUT", body: data });
    showToast("样本已更新");
  } else {
    await api(`/api/projects/${state.project.id}/samples`, { method: "POST", body: data });
    showToast("样本已新增");
  }
  state.editingSampleId = null;
  await refreshDetail();
}

async function submitFileForm(form, path, message) {
  const input = form.querySelector("input[type='file']");
  if (!input?.files?.[0]) throw new Error("请选择文件");
  const body = new FormData();
  body.append("file", input.files[0]);
  const result = await api(path, { method: "POST", body });
  const suffix = result.imported !== undefined ? `，新增 ${result.imported}，更新 ${result.updated}` : "";
  showToast(`${message}${suffix}`);
  await refreshDetail();
}

async function submitBioinfoTask(form) {
  await api(`/api/projects/${state.project.id}/bioinfo/tasks`, { method: "POST", body: formData(form) });
  showToast("生信任务已新增");
  await loadStageData("生信分析");
  await refreshDetail(false);
}

async function submitPayment(form) {
  await api(`/api/projects/${state.project.id}/payment`, { method: "PUT", body: formData(form) });
  showToast("回款已保存");
  await loadStageData("回款");
  await refreshDetail(false);
}

async function deleteSample(id) {
  if (!window.confirm("确认删除这个样本？")) return;
  await api(`/api/samples/${id}`, { method: "DELETE" });
  showToast("样本已删除");
  await refreshDetail();
}

async function saveEditableRow(button) {
  const row = button.closest("tr");
  const id = Number(button.dataset.id);
  const data = {};
  row.querySelectorAll("input[name]").forEach(input => {
    data[input.name] = input.value;
  });
  await api(`/api/samples/${id}`, { method: "PUT", body: data });
  showToast("已保存");
  await refreshDetail();
}

async function completeProject() {
  await api(`/api/projects/${state.project.id}/complete`, { method: "POST" });
  showToast("项目已结束");
  await refreshDetail();
}

async function loadProjects() {
  const data = await api("/api/projects");
  state.projects = data.projects || [];
}

async function openProject(id) {
  state.view = "detail";
  state.activeStage = "样本入库";
  state.editingSampleId = null;
  await loadProject(id);
  await loadSamples(id);
  await loadStageData(state.activeStage);
  renderDetail();
}

async function setStage(stage) {
  state.activeStage = stage;
  state.editingSampleId = null;
  await loadStageData(stage);
  renderDetail();
}

async function refreshDetail(loadStage = true) {
  await loadProject(state.project.id);
  await loadSamples(state.project.id);
  if (loadStage) await loadStageData(state.activeStage);
  renderDetail();
}

async function loadProject(id) {
  const data = await api(`/api/projects/${id}`);
  state.project = data.project;
}

async function loadSamples(id) {
  const data = await api(`/api/projects/${id}/samples`);
  state.samples = data.samples || [];
}

async function loadStageData(stage) {
  if (!state.project) return;
  if (stage === "生信分析") {
    const data = await api(`/api/projects/${state.project.id}/bioinfo/tasks`);
    state.tasks = data.tasks || [];
  }
  if (stage === "合同资料") {
    const data = await api(`/api/projects/${state.project.id}/files`);
    state.files = data.files || [];
  }
  if (stage === "回款") {
    const data = await api(`/api/projects/${state.project.id}/payment`);
    state.payment = data.payment;
  }
}

function logout() {
  state.token = "";
  state.user = null;
  localStorage.removeItem("biotech_pm_token");
  renderAuth();
}

function formData(form) {
  const data = {};
  new FormData(form).forEach((value, key) => {
    if (!(value instanceof File)) data[key] = value;
  });
  return data;
}

function filteredProjects() {
  const term = state.search.trim().toLowerCase();
  if (!term) return state.projects;
  return state.projects.filter(project => {
    return [project.name, project.oa_project_code, project.order_code, project.hospital_name, project.customer_name]
      .filter(Boolean)
      .join(" ")
      .toLowerCase()
      .includes(term);
  });
}

function renderProgress(percent = 0, label = "") {
  const value = Math.max(0, Math.min(100, Number(percent || 0)));
  return `
    <div class="progress">
      <div class="progress-track"><div class="progress-fill" style="width:${value}%"></div></div>
      <div class="progress-text">${value}% · ${h(label || "")}</div>
    </div>
  `;
}

function priorityClass(priority) {
  if (priority === "紧急") return "priority-high";
  if (priority === "重要") return "priority-mid";
  return "priority-normal";
}

function showToast(message, isError = false) {
  toast.textContent = message;
  toast.classList.toggle("error", isError);
  toast.hidden = false;
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => {
    toast.hidden = true;
  }, 3200);
}

function h(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function attr(value) {
  return h(value);
}
