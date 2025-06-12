document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    // This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
    // NOTE: If you are still getting 404, this URL is the problem.
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv"; 

    // IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
    // NOTE: If you are getting errors sending data, this URL is the problem.
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE

    // We will IGNORE MasterEmployees sheet for data fetching and report generation
    // Employee management functions in Apps Script still use the MASTER_SHEET_ID you've set up in code.gs
    // For front-end reporting, all employee and branch data will come from Canvassing Data and predefined list.
    const EMPLOYEE_MASTER_DATA_URL = "UNUSED"; // This URL is not used for client-side data fetching.

    // Header constants for CSV/Sheet columns
    const HEADER_DATE_CANVASSED = 'Date Canvassed';
    const HEADER_CUSTOMER_NAME = 'Customer Name';
    const HEADER_CUSTOMER_ID = 'Customer ID';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_STATUS = 'Status';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_REMARKS = 'Remarks';

    // *** DOM Elements ***
    const statusMessageDiv = document.getElementById('statusMessage');
    const employeeManagementStatusMessageDiv = document.getElementById('employeeManagementStatusMessage');

    const tabButtons = document.querySelectorAll('.tab-button');
    const tabContents = document.querySelectorAll('.tab-content');

    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const monthSelect = document.getElementById('monthSelect');
    const yearSelect = document.getElementById('yearSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const monthYearFilterPanel = document.getElementById('monthYearFilterPanel');

    const totalCustomersSpan = document.getElementById('totalCustomers');
    const totalCanvassedSpan = document.getElementById('totalCanvassed');
    const totalAchievedSpan = document.getElementById('totalAchieved');
    const totalAchievedPercentageSpan = document.getElementById('totalAchievedPercentage');
    const branchPerformanceTableBody = document.getElementById('branchPerformanceTable').querySelector('tbody');
    const employeePerformanceTableBody = document.getElementById('employeePerformanceTable').querySelector('tbody');
    const nonParticipatingBranchesTableBody = document.getElementById('nonParticipatingBranchesTable').querySelector('tbody');
    const nonParticipatingMessageDiv = document.getElementById('nonParticipatingMessage');
    const customerDetailsTableBody = document.getElementById('customerDetailsTable').querySelector('tbody');
    const noCustomerDataMessageDiv = document.getElementById('noCustomerDataMessage');

    // Employee Management Forms
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const employeeNameInput = document.getElementById('employeeName');
    const employeeCodeInput = document.getElementById('employeeCode');
    const employeeBranchNameInput = document.getElementById('employeeBranchName');
    const employeeDesignationInput = document.getElementById('employeeDesignation');

    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsInput = document.getElementById('bulkEmployeeDetails');

    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');

    // *** Data Storage ***
    let allCanvassingData = [];
    let allEmployees = []; // Stores unique employees with their branch and designation
    let allBranches = []; // Stores unique branch names

    // Function to show the loading overlay
    function showLoadingIndicator() {
        const loadingOverlay = document.getElementById('loadingOverlay');
        if (loadingOverlay) {
            loadingOverlay.style.display = 'flex'; // Use flex to center spinner
        }
    }

    // Function to hide the loading overlay
    function hideLoadingIndicator() {
        const loadingOverlay = document.getElementById('loadingOverlay');
        if (loadingOverlay) {
            loadingOverlay.style.display = 'none';
        }
    }

    // Helper function to send data to Google Apps Script
    async function sendDataToGoogleAppsScript(action, data) {
        showLoadingIndicator(); // Show loading indicator before fetch
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ action, data }),
            });

            if (!response.ok) {
                const errorText = await response.text();
                displayEmployeeManagementMessage(`Error: ${errorText}`, true);
                return false;
            }

            const result = await response.json();
            displayEmployeeManagementMessage(result.message, !result.success);
            return result.success;

        } catch (error) {
            console.error('Error sending data to Google Apps Script:', error);
            displayEmployeeManagementMessage(`Failed to send data: ${error.message}`, true);
            return false;
        } finally {
            hideLoadingIndicator(); // Hide loading indicator after fetch (success or failure)
        }
    }

    // Helper function to display messages
    function displayMessage(element, message, isError = false) {
        element.textContent = message;
        element.className = 'message-container'; // Reset classes
        if (isError) {
            element.classList.add('error');
        } else {
            element.classList.add('success');
        }
        element.style.display = 'block';

        // Hide message after 5 seconds
        setTimeout(() => {
            element.style.display = 'none';
            element.textContent = '';
        }, 5000);
    }

    function displayStatusMessage(message, isError = false) {
        displayMessage(statusMessageDiv, message, isError);
    }

    function displayEmployeeManagementMessage(message, isError = false) {
        displayMessage(employeeManagementStatusMessageDiv, message, isError);
    }

    // Function to parse CSV data
    function parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        if (lines.length === 0) return [];

        const headers = lines[0].split(',').map(header => header.trim());
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = lines[i].split(',').map(value => value.trim());
            const entry = {};
            headers.forEach((header, index) => {
                entry[header] = values[index] || '';
            });
            data.push(entry);
        }
        return data;
    }

    // Function to fetch and process data
    async function processData() {
        showLoadingIndicator(); // Show loading indicator before data fetch
        try {
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const csvText = await response.text();
            allCanvassingData = parseCSV(csvText);
            
            // Extract unique employees and branches
            const employeeSet = new Set();
            const branchSet = new Set();
            allCanvassingData.forEach(entry => {
                if (entry[HEADER_EMPLOYEE_CODE] && entry[HEADER_EMPLOYEE_NAME]) {
                    employeeSet.add(JSON.stringify({
                        code: entry[HEADER_EMPLOYEE_CODE],
                        name: entry[HEADER_EMPLOYEE_NAME],
                        branch: entry[HEADER_BRANCH_NAME] || 'N/A',
                        designation: entry[HEADER_DESIGNATION] || 'N/A'
                    }));
                }
                if (entry[HEADER_BRANCH_NAME]) {
                    branchSet.add(entry[HEADER_BRANCH_NAME]);
                }
            });
            
            allEmployees = Array.from(employeeSet).map(str => JSON.parse(str));
            allBranches = Array.from(branchSet).sort();

            populateFilters();
            // No need to call filterAndDisplayReports here, showTab will call it
            displayStatusMessage('Data loaded successfully!', false);

        } catch (error) {
            console.error('Error fetching or processing data:', error);
            displayEmployeeManagementMessage(`Error loading data: ${error.message}`, true);
        } finally {
            hideLoadingIndicator(); // Hide loading indicator after data processing
        }
    }

    // Function to populate branch, employee, month, and year filters
    function populateFilters() {
        // Populate Branch Select
        branchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        allBranches.forEach(branch => {
            const option = document.createElement('option');
            option.value = branch;
            option.textContent = branch;
            branchSelect.appendChild(option);
        });

        // Populate Employee Select (ensure unique employees only once)
        employeeSelect.innerHTML = '<option value="">-- All Employees --</option>';
        const uniqueEmployeesForFilter = new Set();
        allEmployees.forEach(emp => {
             // Use a combination of code and name to ensure uniqueness if codes are not strictly unique globally
            uniqueEmployeesForFilter.add(JSON.stringify({ code: emp.code, name: emp.name }));
        });
        
        Array.from(uniqueEmployeesForFilter).map(str => JSON.parse(str)).sort((a,b) => a.name.localeCompare(b.name)).forEach(emp => {
            const option = document.createElement('option');
            option.value = emp.code; // Use code as value
            option.textContent = `${emp.name} (${emp.code})`;
            employeeSelect.appendChild(option);
        });

        // Populate Month and Year Select
        const currentYear = new Date().getFullYear();
        const currentMonth = new Date().getMonth() + 1; // getMonth() is 0-indexed

        monthSelect.innerHTML = '';
        const months = ["January", "February", "March", "April", "May", "June",
                        "July", "August", "September", "October", "November", "December"];
        months.forEach((month, index) => {
            const option = document.createElement('option');
            option.value = String(index + 1).padStart(2, '0'); // "01", "02", etc.
            option.textContent = month;
            if (index + 1 === currentMonth) {
                option.selected = true;
            }
            monthSelect.appendChild(option);
        });

        yearSelect.innerHTML = '';
        for (let i = currentYear; i >= currentYear - 5; i--) { // Last 5 years
            const option = document.createElement('option');
            option.value = i;
            option.textContent = i;
            if (i === currentYear) {
                option.selected = true;
            }
            yearSelect.appendChild(option);
        }
    }

    // Function to filter data based on selected filters
    function getFilteredData() {
        const selectedBranch = branchSelect.value;
        const selectedEmployeeCode = employeeSelect.value;
        const selectedMonth = monthSelect.value;
        const selectedYear = yearSelect.value;

        return allCanvassingData.filter(entry => {
            const entryDate = new Date(entry[HEADER_DATE_CANVASSED]);
            const entryMonth = String(entryDate.getMonth() + 1).padStart(2, '0');
            const entryYear = entryDate.getFullYear().toString();

            const matchesBranch = !selectedBranch || entry[HEADER_BRANCH_NAME] === selectedBranch;
            const matchesEmployee = !selectedEmployeeCode || entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode;
            const matchesMonth = !selectedMonth || entryMonth === selectedMonth;
            const matchesYear = !selectedYear || entryYear === selectedYear;

            return matchesBranch && matchesEmployee && matchesMonth && matchesYear;
        });
    }

    // Function to filter employees for a specific branch
    function getEmployeesByBranch(branchName) {
        // Filter unique employees for the selected branch
        const employeesInBranch = allEmployees.filter(emp => emp.branch === branchName);
        // Ensure no duplicate entries for the same employee code (e.g., if designation differs slightly)
        const uniqueEmployees = new Map();
        employeesInBranch.forEach(emp => {
            if (!uniqueEmployees.has(emp.code)) {
                uniqueEmployees.set(emp.code, emp);
            }
        });
        return Array.from(uniqueEmployees.values()).sort((a,b) => a.name.localeCompare(b.name));
    }

    // Function to display reports based on filtered data and active tab
    function filterAndDisplayReports() {
        const filteredData = getFilteredData();
        // Get the active tab button's ID
        const activeTabButton = document.querySelector('.tab-button.active');
        const activeTabId = activeTabButton ? activeTabButton.id : null; // Default to null if no active tab

        // Only display report content if a tab is active
        if (activeTabId) {
            switch (activeTabId) {
                case 'allBranchSnapshotTabBtn':
                    displayAllBranchSnapshot(filteredData);
                    break;
                case 'allStaffOverallPerformanceTabBtn':
                    displayAllStaffOverallPerformance(filteredData);
                    break;
                case 'nonParticipatingBranchesTabBtn':
                    displayNonParticipatingBranches(filteredData);
                    break;
                case 'detailedCustomerViewTabBtn':
                    displayDetailedCustomerView(filteredData);
                    break;
                case 'employeeManagementTabBtn':
                    // No data processing needed for this tab as it's form-based
                    break;
            }
        }
    }

    // *** Report Display Functions ***

    function displayAllBranchSnapshot(data) {
        const branchPerformance = {};
        const allPossibleBranchesInPeriod = new Set(data.map(entry => entry[HEADER_BRANCH_NAME]).filter(Boolean));

        // Initialize all branches (even those with no canvassing data in this period but might exist)
        allBranches.forEach(branch => {
            if (!branchPerformance[branch]) {
                branchPerformance[branch] = {
                    totalCustomers: 0,
                    canvassed: 0,
                    achieved: 0,
                    nonParticipating: false // Reset for current period
                };
            }
        });

        data.forEach(entry => {
            const branch = entry[HEADER_BRANCH_NAME];
            if (branch) {
                if (!branchPerformance[branch]) {
                    branchPerformance[branch] = { totalCustomers: 0, canvassed: 0, achieved: 0 };
                }
                branchPerformance[branch].totalCustomers++;
                if (entry[HEADER_STATUS] === 'Canvassed') {
                    branchPerformance[branch].canvassed++;
                }
                if (entry[HEADER_STATUS] === 'Achieved') {
                    branchPerformance[branch].achieved++;
                }
            }
        });

        // Determine non-participating branches for the *selected month and year*
        // A branch is non-participating if it has *no entries* for the selected month/year.
        const selectedMonth = monthSelect.value;
        const selectedYear = yearSelect.value;
        
        const participatingBranchesInSelectedPeriod = new Set(data.filter(entry => {
            const entryDate = new Date(entry[HEADER_DATE_CANVASSED]);
            return String(entryDate.getMonth() + 1).padStart(2, '0') === selectedMonth &&
                   entryDate.getFullYear().toString() === selectedYear;
        }).map(entry => entry[HEADER_BRANCH_NAME]).filter(Boolean));

        // Overall Summary
        let totalCustomersOverall = 0;
        let totalCanvassedOverall = 0;
        let totalAchievedOverall = 0;

        // Iterate through all branches that exist in the master list or had data in the filtered period
        const branchesToReport = new Set([...allBranches, ...Array.from(allPossibleBranchesInPeriod)]);

        branchPerformanceTableBody.innerHTML = '';
        const sortedBranches = Array.from(branchesToReport).sort();

        sortedBranches.forEach(branch => {
            const stats = branchPerformance[branch] || { totalCustomers: 0, canvassed: 0, achieved: 0 };
            const achievedPercentage = stats.canvassed > 0 ? ((stats.achieved / stats.canvassed) * 100).toFixed(2) : '0.00';
            const isNonParticipating = !participatingBranchesInSelectedPeriod.has(branch);

            totalCustomersOverall += stats.totalCustomers;
            totalCanvassedOverall += stats.canvassed;
            totalAchievedOverall += stats.achieved;

            const row = branchPerformanceTableBody.insertRow();
            row.insertCell().textContent = branch;
            row.insertCell().textContent = stats.totalCustomers;
            row.insertCell().textContent = stats.canvassed;
            row.insertCell().textContent = stats.achieved;
            row.insertCell().textContent = `${achievedPercentage}%`;
            row.insertCell().textContent = isNonParticipating ? 'Yes' : 'No';

            // Add data-label for responsive tables
            row.cells[0].setAttribute('data-label', 'Branch Name');
            row.cells[1].setAttribute('data-label', 'Total Customers');
            row.cells[2].setAttribute('data-label', 'Canvassed');
            row.cells[3].setAttribute('data-label', 'Achieved');
            row.cells[4].setAttribute('data-label', 'Achieved %');
            row.cells[5].setAttribute('data-label', 'Non-Participating');
        });

        totalCustomersSpan.textContent = totalCustomersOverall;
        totalCanvassedSpan.textContent = totalCanvassedOverall;
        totalAchievedSpan.textContent = totalAchievedOverall;
        totalAchievedPercentageSpan.textContent = totalCanvassedOverall > 0 ? ((totalAchievedOverall / totalCanvassedOverall) * 100).toFixed(2) + '%' : '0.00%';
    }

    function displayAllStaffOverallPerformance(data) {
        const employeePerformance = {};

        // Aggregate data per employee
        data.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            if (employeeCode) {
                if (!employeePerformance[employeeCode]) {
                    employeePerformance[employeeCode] = {
                        name: entry[HEADER_EMPLOYEE_NAME] || 'N/A',
                        code: employeeCode,
                        designation: entry[HEADER_DESIGNATION] || 'N/A',
                        branch: entry[HEADER_BRANCH_NAME] || 'N/A',
                        totalCustomers: 0,
                        canvassed: 0,
                        achieved: 0
                    };
                }
                employeePerformance[employeeCode].totalCustomers++;
                if (entry[HEADER_STATUS] === 'Canvassed') {
                    employeePerformance[employeeCode].canvassed++;
                }
                if (entry[HEADER_STATUS] === 'Achieved') {
                    employeePerformance[employeeCode].achieved++;
                }
            }
        });

        // Ensure all known employees are included, even if they have no data in the filtered period
        allEmployees.forEach(emp => {
            if (!employeePerformance[emp.code]) {
                employeePerformance[emp.code] = {
                    name: emp.name,
                    code: emp.code,
                    designation: emp.designation || 'N/A',
                    branch: emp.branch || 'N/A',
                    totalCustomers: 0,
                    canvassed: 0,
                    achieved: 0
                };
            }
        });

        employeePerformanceTableBody.innerHTML = '';
        const sortedEmployees = Object.values(employeePerformance).sort((a, b) => a.name.localeCompare(b.name));

        sortedEmployees.forEach(emp => {
            const achievedPercentage = emp.canvassed > 0 ? ((emp.achieved / emp.canvassed) * 100).toFixed(2) : '0.00';
            const row = employeePerformanceTableBody.insertRow();
            row.insertCell().textContent = emp.name;
            row.insertCell().textContent = emp.code;
            row.insertCell().textContent = emp.designation;
            row.insertCell().textContent = emp.branch;
            row.insertCell().textContent = emp.totalCustomers;
            row.insertCell().textContent = emp.canvassed;
            row.insertCell().textContent = emp.achieved;
            row.insertCell().textContent = `${achievedPercentage}%`;

            // Add data-label for responsive tables
            row.cells[0].setAttribute('data-label', 'Employee Name');
            row.cells[1].setAttribute('data-label', 'Employee Code');
            row.cells[2].setAttribute('data-label', 'Designation');
            row.cells[3].setAttribute('data-label', 'Branch');
            row.cells[4].setAttribute('data-label', 'Total Customers');
            row.cells[5].setAttribute('data-label', 'Canvassed');
            row.cells[6].setAttribute('data-label', 'Achieved');
            row.cells[7].setAttribute('data-label', 'Achieved %');
        });
    }

    function displayNonParticipatingBranches(data) {
        nonParticipatingBranchesTableBody.innerHTML = '';
        nonParticipatingMessageDiv.style.display = 'none';

        const selectedMonth = monthSelect.value;
        const selectedYear = yearSelect.value;

        // Get all branches that had *any* activity within the currently selected month/year
        const participatingBranchesInSelectedPeriod = new Set(data.filter(entry => {
            const entryDate = new Date(entry[HEADER_DATE_CANVASSED]);
            return String(entryDate.getMonth() + 1).padStart(2, '0') === selectedMonth &&
                   entryDate.getFullYear().toString() === selectedYear;
        }).map(entry => entry[HEADER_BRANCH_NAME]).filter(Boolean));

        const nonParticipatingBranchesList = allBranches.filter(branch => !participatingBranchesInSelectedPeriod.has(branch));
        
        if (nonParticipatingBranchesList.length === 0) {
            nonParticipatingMessageDiv.style.display = 'block';
            return;
        }

        nonParticipatingBranchesList.sort().forEach(branch => {
            // Find the most recent canvassed date for this branch from all data (not just filtered)
            const branchCanvassingData = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branch && entry[HEADER_DATE_CANVASSED]);
            let lastCanvassedDate = 'N/A';
            if (branchCanvassingData.length > 0) {
                const dates = branchCanvassingData.map(entry => new Date(entry[HEADER_DATE_CANVASSED]));
                const mostRecentDate = new Date(Math.max(...dates));
                lastCanvassedDate = mostRecentDate.toLocaleDateString('en-GB'); // DD/MM/YYYY
            }

            const row = nonParticipatingBranchesTableBody.insertRow();
            row.insertCell().textContent = branch;
            row.insertCell().textContent = lastCanvassedDate;

            // Add data-label for responsive tables
            row.cells[0].setAttribute('data-label', 'Branch Name');
            row.cells[1].setAttribute('data-label', 'Last Canvassed Date');
        });
    }

    function displayDetailedCustomerView(data) {
        customerDetailsTableBody.innerHTML = '';
        noCustomerDataMessageDiv.style.display = 'none';

        if (data.length === 0) {
            noCustomerDataMessageDiv.style.display = 'block';
            return;
        }

        // Sort by Date Canvassed (newest first), then Customer Name
        data.sort((a, b) => {
            const dateA = new Date(a[HEADER_DATE_CANVASSED]);
            const dateB = new Date(b[HEADER_DATE_CANVASSED]);
            if (dateA < dateB) return 1;
            if (dateA > dateB) return -1;
            return a[HEADER_CUSTOMER_NAME].localeCompare(b[HEADER_CUSTOMER_NAME]);
        });

        data.forEach(entry => {
            const row = customerDetailsTableBody.insertRow();
            row.insertCell().textContent = entry[HEADER_CUSTOMER_NAME] || 'N/A';
            row.insertCell().textContent = entry[HEADER_CUSTOMER_ID] || 'N/A';
            row.insertCell().textContent = entry[HEADER_DATE_CANVASSED] ? new Date(entry[HEADER_DATE_CANVASSED]).toLocaleDateString('en-GB') : 'N/A';
            const statusCell = row.insertCell();
            statusCell.textContent = entry[HEADER_STATUS] || 'N/A';
            if (entry[HEADER_STATUS] === 'Achieved') {
                statusCell.classList.add('status-achieved');
            } else if (entry[HEADER_STATUS] === 'Not Canvassed') {
                statusCell.classList.add('status-not-canvassed');
            }
            row.insertCell().textContent = entry[HEADER_EMPLOYEE_NAME] || 'N/A';
            row.insertCell().textContent = entry[HEADER_BRANCH_NAME] || 'N/A';
            row.insertCell().textContent = entry[HEADER_REMARKS] || 'N/A';

            // Add data-label for responsive tables
            row.cells[0].setAttribute('data-label', 'Customer Name');
            row.cells[1].setAttribute('data-label', 'Customer ID');
            row.cells[2].setAttribute('data-label', 'Date Canvassed');
            row.cells[3].setAttribute('data-label', 'Status');
            row.cells[4].setAttribute('data-label', 'Employee Name');
            row.cells[5].setAttribute('data-label', 'Branch Name');
            row.cells[6].setAttribute('data-label', 'Remarks');
        });
    }

    // *** New showTab Function ***
    function showTab(tabButtonId) {
        // Remove active class from all buttons and content
        tabButtons.forEach(btn => btn.classList.remove('active'));
        tabContents.forEach(content => content.classList.remove('active'));

        // Add active class to the specified button and corresponding content
        const button = document.getElementById(tabButtonId);
        if (button) {
            button.classList.add('active');
            const targetTabId = tabButtonId.replace('Btn', ''); // e.g., 'allBranchSnapshotTab'
            const targetContent = document.getElementById(targetTabId);
            if (targetContent) {
                targetContent.classList.add('active');
            }

            // Hide/Show filter panels based on active tab
            if (tabButtonId === 'employeeManagementTabBtn') {
                employeeFilterPanel.style.display = 'none';
                monthYearFilterPanel.style.display = 'none';
                branchSelect.value = ''; // Clear branch selection when in employee management
                employeeSelect.value = ''; // Clear employee selection
                statusMessageDiv.style.display = 'none'; // Hide general status message
            } else if (tabButtonId === 'allStaffOverallPerformanceTabBtn' || tabButtonId === 'detailedCustomerViewTabBtn') {
                employeeFilterPanel.style.display = 'block';
                monthYearFilterPanel.style.display = 'block';
            } else {
                employeeFilterPanel.style.display = 'none';
                monthYearFilterPanel.style.display = 'block';
            }
            // Always refilter and display reports when tab changes
            filterAndDisplayReports();
        }
    }


    // *** Event Listeners ***

    // Tab switching logic (now calls showTab)
    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            showTab(button.id); // Call the new showTab function
        });
    });

    // Filter change listeners
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        employeeSelect.innerHTML = '<option value="">-- All Employees --</option>'; // Reset employee filter

        if (selectedBranch) {
            const employeesInSelectedBranch = getEmployeesByBranch(selectedBranch);
            employeesInSelectedBranch.forEach(emp => {
                const option = document.createElement('option');
                option.value = emp.code;
                option.textContent = `${emp.name} (${emp.code})`;
                employeeSelect.appendChild(option);
            });
        }
        filterAndDisplayReports();
    });
    employeeSelect.addEventListener('change', filterAndDisplayReports);
    monthSelect.addEventListener('change', filterAndDisplayReports);
    yearSelect.addEventListener('change', filterAndDisplayReports);

    // Event Listener for Add Single Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeName = employeeNameInput.value.trim();
            const employeeCode = employeeCodeInput.value.trim();
            const employeeBranchName = employeeBranchNameInput.value.trim();
            const employeeDesignation = employeeDesignationInput.value.trim();

            if (!employeeName || !employeeCode || !employeeBranchName) {
                displayEmployeeManagementMessage('Employee Name, Code, and Branch Name are required.', true);
                return;
            }

            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: employeeName,
                [HEADER_EMPLOYEE_CODE]: employeeCode,
                [HEADER_BRANCH_NAME]: employeeBranchName,
                [HEADER_DESIGNATION]: employeeDesignation
            };

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);

            if (success) {
                // Optionally re-fetch data or update local lists if employee added successfully
                await processData(); // Re-fetch to ensure new employee is available in filters
                addEmployeeForm.reset();
            }
        });
    }

    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const bulkDetails = bulkEmployeeDetailsInput.value.trim();

            if (!branchName || !bulkDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk addition.', true);
                return;
            }

            const employeesToAdd = [];
            const lines = bulkDetails.split('\n');
            for (const line of lines) {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length >= 2) { // Name, Code are mandatory
                    const employeeData = {
                        [HEADER_EMPLOYEE_NAME]: parts[0],
                        [HEADER_EMPLOYEE_CODE]: parts[1],
                        [HEADER_BRANCH_NAME]: branchName,
                        [HEADER_DESIGNATION]: parts[2] || ''
                    };
                    employeesToAdd.push(employeeData);
                }
            }

            if (employeesToAdd.length > 0) {
                const success = await sendDataToGoogleAppsScript('add_bulk_employees', employeesToAdd);
                if (success) {
                    await processData(); // Re-fetch to ensure new employees are available in filters
                    bulkAddEmployeeForm.reset();
                }
            } else {
                displayEmployeeManagementMessage('No valid employee entries found in the bulk details.', true);
            }
        });
    }

    // Event Listener for Delete Employee Form
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
                await processData(); // Re-fetch to update lists after deletion
                deleteEmployeeForm.reset();
            }
        });
    }

    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn'); // Call showTab to set initial active tab and display report
});
