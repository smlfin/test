document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv";
    // IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec"; // Ensure this is YOUR Apps Script URL

    const MONTHLY_WORKING_DAYS = 22; // Common approximation for a month's working days

    const TARGETS = {
        'Branch Manager': {
            'Visit': 10,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Investment Staff': { // Added Investment Staff with custom Visit target
            'Visit': 30,
            'Call': 5 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Default': { // For all other designations not explicitly defined
            'Visit': 5,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        }
    };

    // *** DOM Elements ***
    const branchSelect = document.getElementById('branchSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const employeeSelect = document.getElementById('employeeSelect');
    const viewOptions = document.getElementById('viewOptions');
    const viewBranchSummaryBtn = document.getElementById('viewBranchSummaryBtn');
    const viewBranchPerformanceReportBtn = document.getElementById('viewBranchPerformanceReportBtn');
    const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
    const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
    const viewPerformanceReportBtn = document.getElementById('viewPerformanceReportBtn');

    // Main Report Display Area
    const reportDisplay = document.getElementById('reportDisplay'); // Moved this here for clarity

    // Tab buttons for main navigation
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn'); // NEW: Employee Management Tab Button

    // Main Content Sections to toggle
    const reportsSection = document.getElementById('reportsSection'); // NEW: The main reports section
    const employeeManagementSection = document.getElementById('employeeManagementSection'); // NEW: The employee management section

    // Employee Management Form Elements
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const newEmployeeNameInput = document.getElementById('newEmployeeName');
    const newEmployeeCodeInput = document.getElementById('newEmployeeCode');
    const newBranchNameInput = document.getElementById('newBranchName');
    const newDesignationInput = document.getElementById('newDesignation');
    const employeeManagementMessage = document.getElementById('employeeManagementMessage'); // Central message area for employee management forms

    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsTextarea = document.getElementById('bulkEmployeeDetails');

    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');


    // Global variables to store fetched data
    let allCanvassingData = [];
    let allUniqueBranches = [];
    let allUniqueEmployees = [];
    let selectedBranchEntries = [];
    let selectedEmployeeEntries = [];

    // Utility to format date to YYYY-MM-DD
    const formatDate = (dateString) => {
        const date = new Date(dateString);
        return date.toISOString().split('T')[0];
    };

    // Helper to clear and display messages in a specific div
    function displayMessage(message, type = 'info', targetDiv = reportDisplay) {
        if (targetDiv) {
            targetDiv.innerHTML = `<div class="message ${type}">${message}</div>`;
            targetDiv.style.display = 'block';
            setTimeout(() => {
                targetDiv.style.display = 'none';
            }, 5000); // Hide after 5 seconds
        }
    }

    // Specific message display for employee management forms
    function displayEmployeeManagementMessage(message, isError = false) {
        if (employeeManagementMessage) {
            employeeManagementMessage.textContent = message;
            employeeManagementMessage.style.color = isError ? 'red' : 'green';
            employeeManagementMessage.style.display = 'block';
            setTimeout(() => {
                employeeManagementMessage.style.display = 'none';
            }, 5000);
        }
    }

    // Function to fetch data from Google Sheet
    async function fetchData() {
        displayMessage("Fetching data...", 'info');
        try {
            const response = await fetch(DATA_URL);
            const csvText = await response.text();
            allCanvassingData = parseCSV(csvText);
            processData();
            displayMessage("Data loaded successfully!", 'success');
        } catch (error) {
            console.error('Error fetching data:', error);
            displayMessage('Failed to load data. Please check the DATA_URL and your network connection.', 'error');
        }
    }

    // CSV parsing function
    function parseCSV(csv) {
        const lines = csv.split('\n').filter(line => line.trim() !== ''); // Filter out empty lines
        const headers = lines[0].split(',').map(header => header.trim());
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = lines[i].split(',').map(value => value.trim());
            const entry = {};
            headers.forEach((header, index) => {
                entry[header] = values[index];
            });
            data.push(entry);
        }
        return data;
    }

    // Process fetched data to populate filters
    function processData() {
        allUniqueBranches = [...new Set(allCanvassingData.map(entry => entry['Branch Name']))].sort();
        allUniqueEmployees = [...new Set(allCanvassingData.map(entry => entry['Employee Name']))].sort();

        populateDropdown(branchSelect, allUniqueBranches);
        // Employee dropdown will be populated dynamically based on branch selection
    }

    // Populate dropdown utility
    function populateDropdown(selectElement, items) {
        selectElement.innerHTML = '<option value="">-- Select --</option>'; // Default option
        items.forEach(item => {
            const option = document.createElement('option');
            option.value = item;
            option.textContent = item;
            selectElement.appendChild(option);
        });
    }

    // Filter employees based on selected branch
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            employeeFilterPanel.style.display = 'block';
            selectedBranchEntries = allCanvassingData.filter(entry => entry['Branch Name'] === selectedBranch);
            const employeesInBranch = [...new Set(selectedBranchEntries.map(entry => entry['Employee Name']))].sort();
            populateDropdown(employeeSelect, employeesInBranch);
            viewOptions.style.display = 'flex'; // Show view options
            // Reset employee selection and employee-specific display when branch changes
            employeeSelect.value = "";
            selectedEmployeeEntries = [];
            reportDisplay.innerHTML = '<p>Select an employee or choose a report option.</p>';
        } else {
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none'; // Hide view options
            reportDisplay.innerHTML = '<p>Please select a branch from the dropdown above to view reports.</p>';
            selectedBranchEntries = [];
            selectedEmployeeEntries = [];
        }
    });

    // Handle employee selection
    employeeSelect.addEventListener('change', () => {
        const selectedEmployee = employeeSelect.value;
        if (selectedEmployee) {
            selectedEmployeeEntries = allCanvassingData.filter(entry => entry['Employee Name'] === selectedEmployee && entry['Branch Name'] === branchSelect.value);
            reportDisplay.innerHTML = `<p>Ready to view reports for ${selectedEmployee}.</p>`;
        } else {
            selectedEmployeeEntries = [];
            reportDisplay.innerHTML = '<p>Select an employee or choose a report option.</p>';
        }
    });

    // Helper to calculate total activity
    function calculateTotalActivity(entries) {
        const totalActivity = { 'Visit': 0, 'Call': 0, 'Reference': 0, 'New Customer Leads': 0 };
        entries.forEach(entry => {
            totalActivity['Visit'] += parseInt(entry['No of Visit (Client)']) || 0;
            totalActivity['Call'] += parseInt(entry['No of Call']) || 0;
            totalActivity['Reference'] += parseInt(entry['No of Reference']) || 0;
            totalActivity['New Customer Leads'] += parseInt(entry['No of New Customer Leads']) || 0;
        });
        return totalActivity;
    }

    // Render All Branch Snapshot
    function renderAllBranchSnapshot() {
        reportDisplay.innerHTML = '<h2>All Branch Snapshot</h2>';
        const uniqueBranches = [...new Set(allCanvassingData.map(entry => entry['Branch Name']))];

        uniqueBranches.forEach(branch => {
            const branchEntries = allCanvassingData.filter(entry => entry['Branch Name'] === branch);
            const totalActivity = calculateTotalActivity(branchEntries);

            const branchDiv = document.createElement('div');
            branchDiv.className = 'branch-snapshot';
            branchDiv.innerHTML = `<h3>${branch}</h3>
                                   <p>Total Employees: ${[...new Set(branchEntries.map(e => e['Employee Name']))].length}</p>
                                   <ul class="summary-list">
                                       <li><strong>Total Visits:</strong> ${totalActivity['Visit']}</li>
                                       <li><strong>Total Calls:</strong> ${totalActivity['Call']}</li>
                                       <li><strong>Total References:</strong> ${totalActivity['Reference']}</li>
                                       <li><strong>Total New Customer Leads:</strong> ${totalActivity['New Customer Leads']}</li>
                                   </ul>`;
            reportDisplay.appendChild(branchDiv);
        });
    }

    // Render All Staff Overall Performance Report
    function renderOverallStaffPerformanceReport() {
        reportDisplay.innerHTML = '<h2>All Staff Overall Performance Report</h2>';
        const uniqueEmployees = [...new Set(allCanvassingData.map(entry => entry['Employee Name']))];

        uniqueEmployees.forEach(employeeName => {
            const employeeEntries = allCanvassingData.filter(entry => entry['Employee Name'] === employeeName);
            const totalActivity = calculateTotalActivity(employeeEntries);
            const designation = employeeEntries[0] ? employeeEntries[0]['Designation'] : 'Default'; // Get designation from first entry
            const targets = TARGETS[designation] || TARGETS['Default'];
            const performance = calculatePerformance(totalActivity, targets);

            const employeeDiv = document.createElement('div');
            employeeDiv.className = 'employee-performance';
            employeeDiv.innerHTML = `<h3>${employeeName} (${designation})</h3>
                                    <ul class="summary-list">
                                        <li><strong>Visits:</strong> ${totalActivity['Visit']} / ${targets['Visit']} (${performance['Visit']}%)</li>
                                        <li><strong>Calls:</strong> ${totalActivity['Call']} / ${targets['Call']} (${performance['Call']}%)</li>
                                        <li><strong>References:</strong> ${totalActivity['Reference']} / ${targets['Reference']} (${performance['Reference']}%)</li>
                                        <li><strong>New Customer Leads:</strong> ${totalActivity['New Customer Leads']} / ${targets['New Customer Leads']} (${performance['New Customer Leads']}%)</li>
                                    </ul>`;
            reportDisplay.appendChild(employeeDiv);
        });
    }

    // Function to calculate performance percentage
    function calculatePerformance(actual, target) {
        const performance = {};
        for (const key in actual) {
            if (target[key] !== undefined && target[key] > 0) {
                performance[key] = ((actual[key] / target[key]) * 100).toFixed(2);
            } else {
                performance[key] = 'N/A'; // No target or target is zero
            }
        }
        return performance;
    }

    // Render Branch Summary
    function renderBranchSummary(branchEntries) {
        if (branchEntries.length === 0) {
            reportDisplay.innerHTML = '<p>No data for this branch.</p>';
            return;
        }

        const branchName = branchEntries[0]['Branch Name'];
        reportDisplay.innerHTML = `<h2>Branch Summary for ${branchName}</h2>`;

        const employeesInBranch = [...new Set(branchEntries.map(entry => entry['Employee Name']))];

        employeesInBranch.forEach(employeeName => {
            const employeeEntries = branchEntries.filter(entry => entry['Employee Name'] === employeeName);
            const totalActivity = calculateTotalActivity(employeeEntries);

            const employeeSummaryDiv = document.createElement('div');
            employeeSummaryDiv.className = 'employee-summary-card';
            employeeSummaryDiv.innerHTML = `<h3>${employeeName}</h3>
                                           <ul class="summary-list">
                                               <li><strong>Visits:</strong> ${totalActivity['Visit']}</li>
                                               <li><strong>Calls:</strong> ${totalActivity['Call']}</li>
                                               <li><strong>References:</strong> ${totalActivity['Reference']}</li>
                                               <li><strong>New Customer Leads:</strong> ${totalActivity['New Customer Leads']}</li>
                                           </ul>`;
            reportDisplay.appendChild(employeeSummaryDiv);
        });
    }

    // Render Employee Detailed Entries
    function renderEmployeeDetailedEntries(employeeEntries) {
        if (employeeEntries.length === 0) {
            reportDisplay.innerHTML = '<p>No detailed entries for this employee.</p>';
            return;
        }

        const employeeName = employeeEntries[0]['Employee Name'];
        reportDisplay.innerHTML = `<h2>Detailed Entries for ${employeeName}</h2>`;

        const table = document.createElement('table');
        table.className = 'report-table';
        // Create table header
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const headers = Object.keys(employeeEntries[0]); // Use keys from the first entry as headers
        headers.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });

        // Create table body
        const tbody = table.createTBody();
        employeeEntries.forEach(entry => {
            const row = tbody.insertRow();
            headers.forEach(header => {
                const cell = row.insertCell();
                cell.textContent = entry[header];
            });
        });

        reportDisplay.appendChild(table);
    }

    // Render Employee Summary (Current Month)
    function renderEmployeeSummary(employeeEntries) {
        if (employeeEntries.length === 0) {
            reportDisplay.innerHTML = '<p>No data for this employee.</p>';
            return;
        }

        const employeeName = employeeEntries[0]['Employee Name'];
        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        const currentMonthEntries = employeeEntries.filter(entry => {
            const entryDate = new Date(entry['Date']);
            return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
        });

        const totalActivity = calculateTotalActivity(currentMonthEntries);

        reportDisplay.innerHTML = `<h2>Employee Summary for ${employeeName} (Current Month)</h2>
                                   <ul class="summary-list">
                                       <li><strong>Total Visits:</strong> ${totalActivity['Visit']}</li>
                                       <li><strong>Total Calls:</strong> ${totalActivity['Call']}</li>
                                       <li><strong>Total References:</strong> ${totalActivity['Reference']}</li>
                                       <li><strong>Total New Customer Leads:</strong> ${totalActivity['New Customer Leads']}</li>
                                   </ul>`;

        // Optional: Show detailed entries for the current month below the summary
        if (currentMonthEntries.length > 0) {
            reportDisplay.innerHTML += '<h3>Current Month\'s Detailed Entries</h3>';
            renderEmployeeDetailedEntries(currentMonthEntries);
        } else {
            reportDisplay.innerHTML += '<p>No entries for the current month.</p>';
        }
    }

    // Render Employee Performance Report
    function renderPerformanceReport(employeeEntries, employeeName, designation) {
        reportDisplay.innerHTML = `<h2>Performance Report for ${employeeName} (${designation})</h2>`;

        const totalActivity = calculateTotalActivity(employeeEntries);
        const targets = TARGETS[designation] || TARGETS['Default'];
        const performance = calculatePerformance(totalActivity, targets);

        const performanceDiv = document.createElement('div');
        performanceDiv.className = 'performance-report';
        performanceDiv.innerHTML = `
            <h3>Overall Performance</h3>
            <ul class="summary-list">
                <li><strong>Visits:</strong> ${totalActivity['Visit']} / ${targets['Visit']} (${performance['Visit']}%)</li>
                <li><strong>Calls:</strong> ${totalActivity['Call']} / ${targets['Call']} (${performance['Call']}%)</li>
                <li><strong>References:</strong> ${totalActivity['Reference']} / ${targets['Reference']} (${performance['Reference']}%)</li>
                <li><strong>New Customer Leads:</strong> ${totalActivity['New Customer Leads']} / ${targets['New Customer Leads']} (${performance['New Customer Leads']}%)</li>
            </ul>
            <div class="summary-details-container">
                <div>
                    <h4>Activity Breakdown by Date</h4>
                    <ul class="summary-list">
                        ${employeeEntries.map(entry => `
                            <li>${formatDate(entry['Date'])}:
                                V:${entry['No of Visit (Client)']} |
                                C:${entry['No of Call']} |
                                R:${entry['No of Reference']} |
                                L:${entry['No of New Customer Leads']}
                            </li>`).join('')}
                    </ul>
                </div>
            </div>
        `;
        reportDisplay.appendChild(performanceDiv);
    }

    // Function to send data to Google Apps Script
    // This is a generic function that can be used for add, delete, and bulk operations
    async function sendDataToGoogleAppsScript(actionType, data = {}, messageDiv = employeeManagementMessage) {
        if (messageDiv) {
            messageDiv.innerHTML = 'Sending data...';
            messageDiv.style.color = 'blue';
            messageDiv.style.display = 'block';
        }

        try {
            const formData = new URLSearchParams();
            formData.append('actionType', actionType);
            formData.append('data', JSON.stringify(data)); // Data is sent as a JSON string

            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                },
                body: formData
            });

            if (!response.ok) {
                // Try to parse JSON error first, then fallback to text
                let errorDetails = `HTTP error! status: ${response.status}`;
                try {
                    const errorJson = await response.json();
                    errorDetails += `, message: ${errorJson.message || JSON.stringify(errorJson)}`;
                } catch {
                    const errorText = await response.text();
                    errorDetails += `, message: ${errorText}`;
                }
                throw new Error(errorDetails);
            }

            const result = await response.json(); // Assuming Apps Script returns JSON

            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(result.message, false); // Not an error
                return true;
            } else {
                displayEmployeeManagementMessage(`Error: ${result.message}`, true); // Is an error
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Apps Script:', error);
            displayEmployeeManagementMessage(`Error sending data: ${error.message}`, true);
            return false;
        } finally {
            // Optional: Hide message after a delay even on success/failure
            setTimeout(() => {
                if (messageDiv) messageDiv.style.display = 'none';
            }, 5000);
        }
    }


    // Event Listeners for main report buttons
    viewBranchSummaryBtn.addEventListener('click', () => {
        if (selectedBranchEntries.length > 0) {
            renderBranchSummary(selectedBranchEntries);
        } else {
            displayMessage("No branch selected or no data for this branch.");
        }
    });

    viewBranchPerformanceReportBtn.addEventListener('click', () => {
        if (selectedBranchEntries.length > 0) {
            renderOverallStaffPerformanceReportForBranch(selectedBranchEntries);
        } else {
            displayMessage("No branch selected or no data for this branch to show performance report.");
        }
    });

    function renderOverallStaffPerformanceReportForBranch(branchEntries) {
        if (branchEntries.length === 0) {
            reportDisplay.innerHTML = '<p>No data for this branch.</p>';
            return;
        }

        const branchName = branchEntries[0]['Branch Name'];
        reportDisplay.innerHTML = `<h2>All Staff Performance for ${branchName}</h2>`;
        const uniqueEmployeesInBranch = [...new Set(branchEntries.map(entry => entry['Employee Name']))];

        uniqueEmployeesInBranch.forEach(employeeName => {
            const employeeEntries = branchEntries.filter(entry => entry['Employee Name'] === employeeName);
            const totalActivity = calculateTotalActivity(employeeEntries);
            const designation = employeeEntries[0] ? employeeEntries[0]['Designation'] : 'Default'; // Get designation from first entry
            const targets = TARGETS[designation] || TARGETS['Default'];
            const performance = calculatePerformance(totalActivity, targets);

            const employeeDiv = document.createElement('div');
            employeeDiv.className = 'employee-performance';
            employeeDiv.innerHTML = `<h3>${employeeName} (${designation})</h3>
                                    <ul class="summary-list">
                                        <li><strong>Visits:</strong> ${totalActivity['Visit']} / ${targets['Visit']} (${performance['Visit']}%)</li>
                                        <li><strong>Calls:</strong> ${totalActivity['Call']} / ${targets['Call']} (${performance['Call']}%)</li>
                                        <li><strong>References:</strong> ${totalActivity['Reference']} / ${targets['Reference']} (${performance['Reference']}%)</li>
                                        <li><strong>New Customer Leads:</strong> ${totalActivity['New Customer Leads']} / ${targets['New Customer Leads']} (${performance['New Customer Leads']}%)</li>
                                    </ul>`;
            reportDisplay.appendChild(employeeDiv);
        });
    }

    viewAllEntriesBtn.addEventListener('click', () => {
        if (selectedEmployeeEntries.length > 0) {
            renderEmployeeDetailedEntries(selectedEmployeeEntries);
        } else {
            displayMessage("No employee selected or no data for this employee to show detailed entries.");
        }
    });

    viewEmployeeSummaryBtn.addEventListener('click', () => {
        if (selectedEmployeeEntries.length > 0) {
            renderEmployeeSummary(selectedEmployeeEntries);
        } else {
            displayMessage("No employee selected or no data for this employee to show summary.");
        }
    });

    viewPerformanceReportBtn.addEventListener('click', () => {
        if (selectedEmployeeEntries.length > 0) {
            const employeeDesignation = selectedEmployeeEntries[0]['Designation'] || 'Default';
            renderPerformanceReport(selectedEmployeeEntries, selectedEmployeeEntries[0]['Employee Name'], employeeDesignation);
        } else {
            displayMessage("No employee selected or no data for this employee for performance report.");
        }
    });

    // Function to manage tab visibility
    function showTab(tabButtonId) {
        // Remove 'active' from all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });
        // Add 'active' to the clicked tab button
        document.getElementById(tabButtonId).classList.add('active');

        // Hide all major content sections first
        reportsSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        // Show the relevant section and adjust controls visibility
        if (tabButtonId === 'allBranchSnapshotTabBtn' || tabButtonId === 'allStaffOverallPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            document.querySelector('.controls-panel').style.display = 'block'; // Show controls for reports
            if (tabButtonId === 'allBranchSnapshotTabBtn') {
                renderAllBranchSnapshot();
            } else if (tabButtonId === 'allStaffOverallPerformanceTabBtn') {
                renderOverallStaffPerformanceReport();
            }
        } else if (tabButtonId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
            document.querySelector('.controls-panel').style.display = 'none'; // Hide controls when in employee management
            // Optionally clear existing messages when switching to employee management tab
            displayEmployeeManagementMessage('', false);
        }
    }


    // Event listeners for main tab buttons
    if (allBranchSnapshotTabBtn) {
        allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    }
    if (allStaffOverallPerformanceTabBtn) {
        allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    }
    // NEW: Employee Management tab button listener
    if (employeeManagementTabBtn) {
        employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    }

    // NEW: Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault(); // Prevent default form submission

            const employeeName = newEmployeeNameInput.value.trim();
            const employeeCode = newEmployeeCodeInput.value.trim();
            const branchName = newBranchNameInput.value.trim();
            const designation = newDesignationInput.value.trim(); // Optional

            if (!employeeName || !employeeCode || !branchName) {
                displayEmployeeManagementMessage('Please fill in Employee Name, Code, and Branch Name.', true);
                return;
            }

            const employeeData = {
                "Employee Name": employeeName,
                "Employee Code": employeeCode,
                "Branch Name": branchName,
                "Designation": designation // Include even if empty
            };

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);

            if (success) {
                addEmployeeForm.reset(); // Clear form after submission
                fetchData(); // Refresh data to update dropdowns etc.
            }
        });
    }

    // Existing: Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const bulkDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName || !bulkDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk entry.', true);
                return;
            }

            const employeeLines = bulkDetails.split('\n').filter(line => line.trim() !== '');
            const employeesToAdd = [];

            for (const line of employeeLines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length < 2) { // Need at least Name and Code
                    displayEmployeeManagementMessage(`Skipping invalid line: \"${line}\". Each line must have at least Employee Name and Employee Code.`, true);
                    continue;
                }

                const employeeData = {
                    'Employee Name': parts[0],
                    'Employee Code': parts[1],
                    'Branch Name': branchName, // Apply the common branch name
                    'Designation': parts[2] || '' // Designation is optional
                };
                employeesToAdd.push(employeeData);
            }

            if (employeesToAdd.length > 0) {
                const success = await sendDataToGoogleAppsScript('add_bulk_employees', employeesToAdd);
                if (success) {
                    bulkAddEmployeeForm.reset(); // Clear form after submission
                    fetchData(); // Refresh data
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

            const deleteData = { 'Employee Code': employeeCodeToDelete };
            const success = await sendDataToGoogleAppsScript('delete_employee', deleteData);

            if (success) {
                deleteEmployeeForm.reset(); // Clear form after submission
                fetchData(); // Refresh data
            }
        });
    }

    // Initial data fetch when the page loads
    fetchData();

    // Show the "All Branch Snapshot" tab by default on load
    showTab('allBranchSnapshotTabBtn');
});
