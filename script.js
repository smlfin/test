document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    // This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
    // NOTE: If you are still getting 404, this URL is the problem.
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv"; 

    // IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
    // NOTE: If you are getting errors sending data, this URL is the problem.
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE

    const MONTHLY_WORKING_DAYS = 22; // Common approximation for a month's working days

    const TARGETS = {
        'Branch Manager': {
            'Visit': 10,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Investment Staff': { // Added Investment Staff with custom Visit target
            'Visit': 30, // Example: Higher visit target for Investment Staff
            'Call': 5 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 2 // Example: Lower leads target
        },
        'Sales Executive': {
            'Visit': 5 * MONTHLY_WORKING_DAYS,
            'Call': 10 * MONTHLY_WORKING_DAYS,
            'Reference': 2 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 50
        }
    };

    // --- Headers for Canvassing Data (from DATA_URL) ---
    const HEADER_ACTIVITY_DATE = 'Activity Date';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_VISITS = 'Visits';
    const HEADER_CALLS = 'Calls';
    const HEADER_REFERENCES = 'References';
    const HEADER_NEW_CUSTOMER_LEADS = 'New Customer Leads';

    // *** IMPORTANT: UPDATE THIS LIST WITH YOUR ACTUAL BRANCH NAMES FROM LIV).XLSX - SHEET1.CSV ***
    // This list should contain ALL expected branch names for non-participation reporting.
    const allBranches = [
        "Your Actual Branch Name 1", 
        "Your Actual Branch Name 2", 
        "Your Actual Branch Name 3",
        // ... add all other unique branch names here, exactly as they appear in your CSV
    ]; 

    let canvassingData = []; // Raw data from CSV
    let currentFilteredData = []; // Data currently displayed after filters

    // *** DOM Elements ***
    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const startDateInput = document.getElementById('startDate');
    const endDateInput = document.getElementById('endDate');
    const applyFilterBtn = document.getElementById('applyFilterBtn');
    const reportsSection = document.getElementById('reportsSection');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const viewOptions = document.getElementById('viewOptions');

    // Tab Buttons
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');

    // Report Sections and Table Bodies
    const allBranchSnapshotSection = document.getElementById('allBranchSnapshotSection');
    const allStaffOverallPerformanceSection = document.getElementById('allStaffOverallPerformanceSection');
    const employeeDetailsPerformanceSection = document.getElementById('employeeDetailsPerformanceSection');
    const nonParticipatingBranchesSection = document.getElementById('nonParticipatingBranchesSection'); // New
    const allBranchSnapshotTableBody = document.getElementById('allBranchSnapshotTableBody');
    const allStaffOverallPerformanceTableBody = document.getElementById('allStaffOverallPerformanceTableBody');
    const employeeDetailsPerformanceBody = document.getElementById('employeeDetailsPerformanceBody');

    // Elements for Employee Performance Details Summary
    const totalVisitsElement = document.getElementById('totalVisits');
    const totalCallsElement = document.getElementById('totalCalls');
    const totalReferencesElement = document.getElementById('totalReferences');
    const totalNewCustomerLeadsElement = document.getElementById('totalNewCustomerLeads');

    // Employee Management Section Elements
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    const statusMessage = document.getElementById('statusMessage');
    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeDetailsTextarea = document.getElementById('bulkEmployeeDetails');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');

    // Non-participating Branches elements
    const noParticipationMessage = document.getElementById('noParticipationMessage');
    const nonParticipatingBranchList = document.getElementById('nonParticipatingBranchList');


    // Helper to display messages for employee management
    function displayEmployeeManagementMessage(message, isError = false) {
        statusMessage.textContent = message;
        statusMessage.className = `message-container ${isError ? 'error-message' : 'success-message'}`;
        statusMessage.style.display = 'block';
        setTimeout(() => {
            statusMessage.style.display = 'none';
        }, 5000); // Hide after 5 seconds
    }

    // Function to send data to Google Apps Script
    async function sendDataToGoogleAppsScript(action, data) {
        const payload = {
            action: action,
            data: data
        };

        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Crucial for cross-origin requests
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(payload)
            });

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(result.message || 'Operation successful!');
                return true;
            } else {
                displayEmployeeManagementMessage(result.message || 'Operation failed!', true);
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Apps Script:', error);
            displayEmployeeManagementMessage('Network error or Apps Script issue. Check console.', true);
            return false;
        }
    }


    // Function to fetch and parse Canvassing Data
    async function fetchData() {
        try {
            PapaParse.parse(DATA_URL, {
                download: true,
                header: true,
                complete: function(results) {
                    // Filter out rows without a valid activity date or with all empty values
                    canvassingData = results.data.filter(row => 
                        row[HEADER_ACTIVITY_DATE] && Object.values(row).some(val => val !== null && val !== '')
                    );

                    // Initial filter and render
                    currentFilteredData = filterDataByDate(canvassingData);
                    renderReports(currentFilteredData);
                    populateBranchDropdown(canvassingData);
                    populateEmployeeDropdown(canvassingData);
                    // Show the "All Branch Snapshot" tab by default on initial load
                    showTab('allBranchSnapshotTabBtn');
                },
                error: function(err) {
                    console.error("Error parsing Canvassing Data CSV:", err);
                    alert("Failed to load canvassing data. Please check the DATA_URL and your network connection.");
                }
            });
        } catch (error) {
            console.error("Error fetching data:", error);
            alert("An error occurred while fetching data.");
        }
    }

    // Function to filter data by date range
    function filterDataByDate(data) {
        const startDate = startDateInput.value ? new Date(startDateInput.value) : null;
        const endDate = endDateInput.value ? new Date(endDateInput.value) : null;

        return data.filter(entry => {
            const entryDate = new Date(entry[HEADER_ACTIVITY_DATE]);
            // Ensure entryDate is valid before comparison
            if (isNaN(entryDate.getTime())) {
                return false; // Skip invalid dates
            }
            return (!startDate || entryDate >= startDate) &&
                   (!endDate || entryDate <= endDate);
        });
    }

    // Function to populate the branch dropdown
    function populateBranchDropdown(data) {
        const branches = new Set();
        data.forEach(entry => {
            if (entry[HEADER_BRANCH_NAME]) {
                branches.add(entry[HEADER_BRANCH_NAME]);
            }
        });

        branchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        Array.from(branches).sort().forEach(branch => {
            const option = document.createElement('option');
            option.value = branch;
            option.textContent = branch;
            branchSelect.appendChild(option);
        });
    }

    // Function to populate the employee dropdown based on selected branch or all data
    function populateEmployeeDropdown(data, selectedBranch = '') {
        const employees = new Map(); // Use Map to store unique employees with their designations

        const filteredByBranch = selectedBranch ?
            data.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch) :
            data;

        filteredByBranch.forEach(entry => {
            if (entry[HEADER_EMPLOYEE_CODE] && entry[HEADER_EMPLOYEE_NAME]) {
                // Store employee name and designation, overwrite if a more recent designation is found (or just take the first)
                if (!employees.has(entry[HEADER_EMPLOYEE_CODE])) {
                    employees.set(entry[HEADER_EMPLOYEE_CODE], {
                        name: entry[HEADER_EMPLOYEE_NAME],
                        designation: entry[HEADER_DESIGNATION] || 'N/A'
                    });
                }
            }
        });

        employeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';
        Array.from(employees.entries())
            .sort((a, b) => a[1].name.localeCompare(b[1].name)) // Sort by employee name
            .forEach(([code, details]) => {
                const option = document.createElement('option');
                option.value = code;
                option.textContent = `${details.name} (${code}) - ${details.designation}`;
                employeeSelect.appendChild(option);
            });
    }


    // Function to update the All Branch Snapshot Table
    function updateAllBranchSnapshotTable(filteredData) {
        const branchData = new Map(); // branchName -> {employees, visits, calls, references, newLeads}

        filteredData.forEach(entry => {
            const branchName = entry[HEADER_BRANCH_NAME];
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];

            if (!branchName) return;

            if (!branchData.has(branchName)) {
                branchData.set(branchName, {
                    employees: new Set(), // Use a Set to count unique employees per branch
                    visits: 0,
                    calls: 0,
                    references: 0,
                    newLeads: 0
                });
            }

            const currentBranch = branchData.get(branchName);
            currentBranch.employees.add(employeeCode); // Add employee code to the set

            currentBranch.visits += parseInt(entry[HEADER_VISITS] || 0);
            currentBranch.calls += parseInt(entry[HEADER_CALLS] || 0);
            currentBranch.references += parseInt(entry[HEADER_REFERENCES] || 0);
            currentBranch.newLeads += parseInt(entry[HEADER_NEW_CUSTOMER_LEADS] || 0);
        });

        allBranchSnapshotTableBody.innerHTML = ''; // Clear existing rows

        Array.from(branchData.entries()).sort((a, b) => a[0].localeCompare(b[0])).forEach(([branchName, totals]) => {
            const row = allBranchSnapshotTableBody.insertRow();
            row.insertCell().textContent = branchName;
            row.insertCell().textContent = totals.employees.size; // Count unique employees
            row.insertCell().textContent = totals.visits;
            row.insertCell().textContent = totals.calls;
            row.insertCell().textContent = totals.references;
            row.insertCell().textContent = totals.newLeads;
        });

        if (branchData.size === 0) {
            allBranchSnapshotTableBody.innerHTML = '<tr><td colspan="6">No data available for the selected period.</td></tr>';
        }
    }

    // Function to update the All Staff Overall Performance Table
    function updateAllStaffOverallPerformanceTable(filteredData) {
        const employeePerformance = new Map(); // employeeCode -> {name, branch, designation, visits, calls, references, newLeads}

        filteredData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const branchName = entry[HEADER_BRANCH_NAME];
            const designation = entry[HEADER_DESIGNATION];

            if (!employeeCode || !employeeName) return;

            if (!employeePerformance.has(employeeCode)) {
                employeePerformance.set(employeeCode, {
                    name: employeeName,
                    branch: branchName || 'N/A',
                    designation: designation || 'N/A',
                    visits: 0,
                    calls: 0,
                    references: 0,
                    newLeads: 0
                });
            }

            const currentEmployee = employeePerformance.get(employeeCode);
            currentEmployee.visits += parseInt(entry[HEADER_VISITS] || 0);
            currentEmployee.calls += parseInt(entry[HEADER_CALLS] || 0);
            currentEmployee.references += parseInt(entry[HEADER_REFERENCES] || 0);
            currentEmployee.newLeads += parseInt(entry[HEADER_NEW_CUSTOMER_LEADS] || 0);
        });

        allStaffOverallPerformanceTableBody.innerHTML = ''; // Clear existing rows

        Array.from(employeePerformance.entries()).sort((a, b) => a[1].name.localeCompare(b[1].name)).forEach(([, employee]) => {
            const row = allStaffOverallPerformanceTableBody.insertRow();
            row.insertCell().textContent = employee.name;
            row.insertCell().textContent = employee.branch;
            row.insertCell().textContent = employee.designation;
            row.insertCell().textContent = employee.visits;
            row.insertCell().textContent = employee.calls;
            row.insertCell().textContent = employee.references;
            row.insertCell().textContent = employee.newLeads;
        });

        if (employeePerformance.size === 0) {
            allStaffOverallPerformanceTableBody.innerHTML = '<tr><td colspan="7">No data available for the selected period.</td></tr>';
        }
    }


    // Function to update individual employee performance details (when an employee is selected)
    function updateEmployeeDetailsPerformance(filteredData, employeeCode) {
        const employeeData = filteredData.find(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);

        if (!employeeData) {
            employeeDetailsPerformanceSection.style.display = 'none';
            return;
        }

        let totalVisits = 0;
        let totalCalls = 0;
        let totalReferences = 0;
        let totalNewCustomerLeads = 0;

        filteredData.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode).forEach(entry => {
            totalVisits += parseInt(entry[HEADER_VISITS] || 0);
            totalCalls += parseInt(entry[HEADER_CALLS] || 0);
            totalReferences += parseInt(entry[HEADER_REFERENCES] || 0);
            totalNewCustomerLeads += parseInt(entry[HEADER_NEW_CUSTOMER_LEADS] || 0);
        });

        totalVisitsElement.textContent = totalVisits;
        totalCallsElement.textContent = totalCalls;
        totalReferencesElement.textContent = totalReferences;
        totalNewCustomerLeadsElement.textContent = totalNewCustomerLeads;

        // Calculate targets based on designation
        const designation = employeeData[HEADER_DESIGNATION] || 'Sales Executive'; // Default to Sales Executive if not found
        const targets = TARGETS[designation] || TARGETS['Sales Executive']; // Fallback

        employeeDetailsPerformanceBody.innerHTML = ''; // Clear previous rows

        const activities = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
        const headers = {
            'Visit': HEADER_VISITS,
            'Call': HEADER_CALLS,
            'Reference': HEADER_REFERENCES,
            'New Customer Leads': HEADER_NEW_CUSTOMER_LEADS
        };

        activities.forEach(activity => {
            const count = parseInt(employeeData[headers[activity]] || 0);
            const target = targets[activity] || 0;
            const percentageOfTarget = target > 0 ? ((count / target) * 100).toFixed(2) : (count > 0 ? '100.00' : '0.00');

            const row = employeeDetailsPerformanceBody.insertRow();
            row.insertCell().textContent = activity;
            row.insertCell().textContent = count;
            row.insertCell().textContent = target;
            row.insertCell().textContent = `${percentageOfTarget}%`;
        });

        employeeDetailsPerformanceSection.style.display = 'block';
    }


    // Function to update non-participating branches
    function updateNonParticipatingBranches(filteredData) {
        const participatingBranches = new Set();
        filteredData.forEach(entry => {
            if (entry[HEADER_BRANCH_NAME]) {
                participatingBranches.add(entry[HEADER_BRANCH_NAME]);
            }
        });

        const nonParticipating = allBranches.filter(branch => !participatingBranches.has(branch));

        if (nonParticipating.length > 0) {
            noParticipationMessage.textContent = 'Branches with No Activity in the Selected Period:';
            nonParticipatingBranchList.innerHTML = ''; // Clear previous list
            nonParticipating.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                nonParticipatingBranchList.appendChild(li);
            });
            noParticipationMessage.style.display = 'block';
            nonParticipatingBranchList.style.display = 'block';
        } else {
            noParticipationMessage.textContent = 'All branches have activity in the selected period.';
            nonParticipatingBranchList.innerHTML = '';
            noParticipationMessage.style.display = 'block';
            nonParticipatingBranchList.style.display = 'none';
        }
    }

    // Function to render all reports based on filtered data
    function renderReports(dataToRender) {
        const selectedBranch = branchSelect.value;
        const selectedEmployee = employeeSelect.value;

        // Update all branch snapshot table only if no specific employee is selected
        if (!selectedEmployee) {
            updateAllBranchSnapshotTable(dataToRender);
            updateAllStaffOverallPerformanceTable(dataToRender);
            employeeDetailsPerformanceSection.style.display = 'none'; // Hide if no employee selected
        } else {
            // If an employee is selected, filter data for that employee and update their performance table
            const employeeSpecificData = dataToRender.filter(entry => entry[HEADER_EMPLOYEE_CODE] === selectedEmployee);
            updateEmployeeDetailsPerformance(employeeSpecificData, selectedEmployee);
            allBranchSnapshotTableBody.innerHTML = '<tr><td colspan="6">Select "All Staff Performance (Overall)" or a Branch to see branch/overall reports.</td></tr>';
            allStaffOverallPerformanceTableBody.innerHTML = '<tr><td colspan="7">Select "All Staff Performance (Overall)" or a Branch to see branch/overall reports.</td></tr>';
        }

        // Always update non-participating branches regardless of selected branch/employee
        // as this report is about overall branches without activity
        updateNonParticipatingBranches(dataToRender);
    }


    // Event Listeners for Tab Buttons
    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));

    // Event Listener for Apply Filter Button
    if (applyFilterBtn) {
        applyFilterBtn.addEventListener('click', () => {
            currentFilteredData = filterDataByDate(canvassingData);
            renderReports(currentFilteredData);
            // After applying filter, ensure the currently active tab's section is shown
            // This is handled by renderReports itself and showTab is not needed here
            // but if an employee or branch is selected, re-filter relevant dropdowns too
            const selectedBranch = branchSelect.value;
            populateEmployeeDropdown(canvassingData, selectedBranch);
        });
    }

    // Event Listener for Branch Select
    if (branchSelect) {
        branchSelect.addEventListener('change', () => {
            const selectedBranch = branchSelect.value;
            // When branch changes, re-filter employee dropdown
            populateEmployeeDropdown(canvassingData, selectedBranch);

            // Re-filter data based on selected branch and dates, then render reports
            let tempFiltered = filterDataByDate(canvassingData);
            if (selectedBranch) {
                tempFiltered = tempFiltered.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch);
            }
            // If employee is selected, further filter
            const selectedEmployee = employeeSelect.value;
            if (selectedEmployee) {
                tempFiltered = tempFiltered.filter(entry => entry[HEADER_EMPLOYEE_CODE] === selectedEmployee);
            }
            currentFilteredData = tempFiltered;
            renderReports(currentFilteredData);
        });
    }

    // Event Listener for Employee Select
    if (employeeSelect) {
        employeeSelect.addEventListener('change', () => {
            const selectedBranch = branchSelect.value;
            const selectedEmployee = employeeSelect.value;
            // Re-filter data based on selected branch, employee, and dates, then render reports
            let tempFiltered = filterDataByDate(canvassingData);
            if (selectedBranch) {
                tempFiltered = tempFiltered.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch);
            }
            if (selectedEmployee) {
                tempFiltered = tempFiltered.filter(entry => entry[HEADER_EMPLOYEE_CODE] === selectedEmployee);
            }
            currentFilteredData = tempFiltered;
            renderReports(currentFilteredData);
        });
    }

    // Function to handle tab visibility and content display
    function showTab(tabId) {
        // Hide all main sections initially
        reportsSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        // Hide all report subsections within reportsSection
        allBranchSnapshotSection.style.display = 'none';
        allStaffOverallPerformanceSection.style.display = 'none';
        employeeDetailsPerformanceSection.style.display = 'none';
        nonParticipatingBranchesSection.style.display = 'none';

        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Show the active tab content and activate the clicked button
        switch (tabId) {
            case 'allBranchSnapshotTabBtn':
            case 'allStaffOverallPerformanceTabBtn':
            case 'nonParticipatingBranchesTabBtn':
                reportsSection.style.display = 'block'; // Always show reports section for these tabs
                if (tabId === 'allBranchSnapshotTabBtn') {
                    allBranchSnapshotSection.style.display = 'block';
                } else if (tabId === 'allStaffOverallPerformanceTabBtn') {
                    allStaffOverallPerformanceSection.style.display = 'block';
                } else if (tabId === 'nonParticipatingBranchesTabBtn') {
                    nonParticipatingBranchesSection.style.display = 'block';
                }
                document.getElementById(tabId).classList.add('active');
                // Re-render reports for the selected tab with current filtered data
                // This ensures the displayed report is up-to-date when switching tabs
                renderReports(currentFilteredData); 
                break;
            case 'employeeManagementTabBtn':
                employeeManagementSection.style.display = 'block';
                document.getElementById(tabId).classList.add('active');
                break;
        }
    }

    // Existing: Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const employeeDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName || !employeeDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required.', true);
                return;
            }

            const employeesToAdd = [];
            const lines = employeeDetails.split('\n');
            lines.forEach(line => {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length >= 2) { // Expecting at least Name,Code
                    const employeeData = {
                        [HEADER_EMPLOYEE_NAME]: parts[0],
                        [HEADER_EMPLOYEE_CODE]: parts[1],
                        [HEADER_BRANCH_NAME]: branchName,
                        [HEADER_DESIGNATION]: parts[2] || '' // Designation is optional
                    };
                    employeesToAdd.push(employeeData);
                }
            });

            if (employeesToAdd.length > 0) {
                const success = await sendDataToGoogleAppsScript('add_bulk_employees', employeesToAdd);
                if (success) {
                    bulkAddEmployeeForm.reset(); // Clear form after submission
                    fetchData(); // Refresh data from Google Sheet after adding employees
                }
            } else {
                displayEmployeeManagementMessage('No valid employee entries found in the bulk details.', true);
            }
        });
    }

    // Existing: Event Listener for Delete Employee Form
    if (deleteEmployeeForm) {
        deleteEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeCodeToDelete = deleteEmployeeCodeInput.value.trim();

            if (!employeeCodeToDelete) {
                displayEmployeeManagementMessage('Employee Code is required for deletion.', true);
                return;
            }

            const deleteData = { [HEADER_EMPLOYEE_CODE]: employeeCodeToDelete };
            const success = await sendDataToGoogleAppsScript('delete_employee', deleteData);

            if (success) {
                deleteEmployeeForm.reset(); // Clear form after submission
                fetchData(); // Refresh data from Google Sheet after deleting employee
            }
        });
    }

    // Initial data fetch when the page loads
    fetchData(); // This will also call showTab('allBranchSnapshotTabBtn') upon completion
});
