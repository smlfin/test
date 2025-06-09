// Assuming this script is placed in script.js and attached to the HTML

// Global Maps
const employeeMap = {}; // { empCode: {name, branch, designation, code} }
const performanceMap = {}; // { empCode: {calls, visits, references, leads, callTarget, ...} }
const empBranchMap = {}; // { empCode: branch }

document.addEventListener("DOMContentLoaded", () => {
    initializeDashboard();
});

function initializeDashboard() {
    setupTabSwitching();
    google.script.run.withSuccessHandler(loadEmployeeData).getEmployeeData();
    google.script.run.withSuccessHandler(loadPerformanceData).getPerformanceData();
}

function setupTabSwitching() {
    document.getElementById("allBranchSnapshotTabBtn").addEventListener("click", () => switchTab("allBranchSnapshotContainer"));
    document.getElementById("allStaffOverallPerformanceTabBtn").addEventListener("click", () => switchTab("allStaffOverallPerformanceContainer"));
    document.getElementById("nonParticipatingBranchesTabBtn").addEventListener("click", () => switchTab("nonParticipatingBranchesContainer"));
    document.getElementById("employeeManagementTabBtn").addEventListener("click", () => switchTab("employeeManagementSection"));
}

function switchTab(containerId) {
    const containers = document.querySelectorAll(".report-container, .form-section-container");
    containers.forEach(el => el.style.display = "none");
    document.getElementById(containerId).style.display = "block";
}

function loadEmployeeData(data) {
    data.forEach(emp => {
        employeeMap[emp.code] = emp;
        empBranchMap[emp.code] = emp.branch;
    });
}

function loadPerformanceData(data) {
    data.forEach(perf => {
        performanceMap[perf.code] = perf;
    });
    renderAllBranchSnapshot();
    renderAllStaffPerformance();
    renderNonParticipatingBranches();
}

function renderAllBranchSnapshot() {
    const tbody = document.getElementById('allBranchSnapshotTableBody');
    tbody.innerHTML = '';

    const branchStats = {};
    Object.values(employeeMap).forEach(emp => {
        if (!branchStats[emp.branch]) {
            branchStats[emp.branch] = { employees: new Set(), calls: 0, visits: 0, references: 0, leads: 0 };
        }
        branchStats[emp.branch].employees.add(emp.code);

        const perf = performanceMap[emp.code];
        if (perf) {
            branchStats[emp.branch].calls += perf.calls || 0;
            branchStats[emp.branch].visits += perf.visits || 0;
            branchStats[emp.branch].references += perf.references || 0;
            branchStats[emp.branch].leads += perf.leads || 0;
        }
    });

    let totalEmployees = new Set();
    let totalBranches = 0;
    let overall = { calls: 0, visits: 0, references: 0, leads: 0 };

    Object.entries(branchStats).forEach(([branch, stats]) => {
        if (stats.employees.size > 0) totalBranches++;
        stats.employees.forEach(e => totalEmployees.add(e));

        overall.calls += stats.calls;
        overall.visits += stats.visits;
        overall.references += stats.references;
        overall.leads += stats.leads;

        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${branch}</td>
            <td>${stats.employees.size}</td>
            <td>${stats.calls}</td>
            <td>${stats.visits}</td>
            <td>${stats.references}</td>
            <td>${stats.leads}</td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById("totalBranchesParticipated").textContent = totalBranches;
    document.getElementById("overallEmployeesParticipated").textContent = totalEmployees.size;
    document.getElementById("overallCalls").textContent = overall.calls;
    document.getElementById("overallVisits").textContent = overall.visits;
    document.getElementById("overallReferences").textContent = overall.references;
    document.getElementById("overallNewCustomerLeads").textContent = overall.leads;
}

function renderAllStaffPerformance() {
    const tbody = document.getElementById("allStaffPerformanceTableBody");
    tbody.innerHTML = "";

    Object.values(employeeMap).forEach(emp => {
        const perf = performanceMap[emp.code] || {
            calls: 0, visits: 0, references: 0, leads: 0,
            callTarget: 0, visitTarget: 0, referenceTarget: 0, newCustomerLeadTarget: 0
        };

        const row = document.createElement("tr");
        row.innerHTML = `
            <td class="employee-name-cell" data-code="${emp.code}">${emp.name}</td>
            <td>${emp.code}</td>
            <td>${emp.branch}</td>
            <td>${emp.designation}</td>
            <td>${perf.calls}</td>
            <td>${perf.visits}</td>
            <td>${perf.references}</td>
            <td>${perf.leads}</td>
            <td>${perf.callTarget}</td>
            <td>${perf.visitTarget}</td>
            <td>${perf.referenceTarget}</td>
            <td>${perf.newCustomerLeadTarget}</td>
        `;
        tbody.appendChild(row);
    });

    addClickListenerToEmployeeNames();
}

function renderNonParticipatingBranches() {
    const list = document.getElementById("nonParticipatingBranchesList");
    const message = document.getElementById("noParticipationMessage");

    const allBranches = new Set();
    const participatingBranches = new Set();

    Object.values(employeeMap).forEach(emp => allBranches.add(emp.branch));
    Object.keys(performanceMap).forEach(code => {
        const emp = employeeMap[code];
        if (emp) participatingBranches.add(emp.branch);
    });

    const nonParticipating = [...allBranches].filter(branch => !participatingBranches.has(branch));

    list.innerHTML = "";
    if (nonParticipating.length === 0) {
        message.textContent = "All branches have participating employees.";
    } else {
        message.textContent = "The following branches have no participation:";
        nonParticipating.forEach(branch => {
            const li = document.createElement("li");
            li.textContent = branch;
            list.appendChild(li);
        });
    }
}

function addClickListenerToEmployeeNames() {
    // Optional: Implement to handle employee detail view on name click
}
