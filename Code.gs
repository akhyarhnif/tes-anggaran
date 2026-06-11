// ============================================================
// Code.gs — Main Router & Entry Point
// Budget vs Realisasi System
// ============================================================

const SS_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const ss = SpreadsheetApp.openById(SS_ID);

// ---------- Web App Entry Points ----------

function doGet(e) {
  const page = e.parameter.page || 'dashboard';
  const template = HtmlService.createTemplateFromFile(page);
  template.user = Session.getActiveUser().getEmail();
  return template.evaluate()
    .setTitle('Budget Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------- Dispatcher (dipanggil dari frontend via google.script.run) ----------

function dispatch(action, payload) {
  const actions = {
    // Department & Category
    getDepartments:     () => DepartmentService.getAll(),
    getCategories:      (p) => DepartmentService.getCategoriesByDept(p.deptId),
    saveDepartment:     (p) => DepartmentService.save(p),
    saveCategory:       (p) => DepartmentService.saveCategory(p),

    // Budget
    getBudgets:         (p) => BudgetService.getAll(p),
    saveBudget:         (p) => BudgetService.save(p),
    deleteBudget:       (p) => BudgetService.delete(p.id),

    // Realisasi
    getRealisasi:       (p) => BudgetService.getRealisasi(p),
    inputRealisasi:     (p) => BudgetService.inputRealisasi(p),

    // Lock
    checkLock:          (p) => LockService.check(p),
    getOverbudgetItems: ()  => LockService.getOverbudgetItems(),

    // Approval
    submitApproval:     (p) => ApprovalService.submit(p),
    getApprovals:       (p) => ApprovalService.getAll(p),
    approveRequest:     (p) => ApprovalService.approve(p),
    rejectRequest:      (p) => ApprovalService.reject(p),

    // Dashboard
    getDashboardStats:  ()  => DashboardService.getStats(),
    getPerformanceData: (p) => DashboardService.getPerformance(p),
  };

  if (!actions[action]) throw new Error(`Action tidak ditemukan: ${action}`);

  try {
    return { success: true, data: actions[action](payload) };
  } catch (err) {
    Logger.log(`[ERROR] ${action}: ${err.message}`);
    return { success: false, error: err.message };
  }
}

// ---------- Setup (jalankan sekali) ----------

function setupSpreadsheet() {
  const sheets = [
    { name: 'Departments',  headers: ['id','name','code','createdAt'] },
    { name: 'Categories',   headers: ['id','deptId','name','code','createdAt'] },
    { name: 'Budgets',      headers: ['id','deptId','categoryId','period','amount','status','createdAt','createdBy'] },
    { name: 'Realisasi',    headers: ['id','budgetId','deptId','categoryId','period','amount','description','inputAt','inputBy'] },
    { name: 'Locks',        headers: ['id','budgetId','deptId','categoryId','period','reason','lockedAt'] },
    { name: 'Approvals',    headers: ['id','budgetId','deptId','categoryId','period','requestedAmount','currentAmount','reason','status','requestedBy','requestedAt','reviewedBy','reviewedAt','notes'] },
  ];

  sheets.forEach(({ name, headers }) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    }
  });

  // Seed data departemen awal
  DepartmentService.seedInitialData();

  SpreadsheetApp.getUi().alert('✅ Setup selesai! Spreadsheet sudah siap.');
}
