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
        'Investment Staff': {
            'Visit': 30,
            'Call': 5 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Seniors': {
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
    // Predefined list of branches for the dropdown and "no participation" check
    const PREDEFINED_BRANCHES = [
        "Angamaly", "Corporate Office", "Edappally", "Harippad", "Koduvayur", "Kuzhalmannam",
        "Mattanchery", "Mavelikara", "Nedumkandom", "Nenmara", "Paravoor", "Perumbavoor",
        "Thiruwillamala", "Thodupuzha", "Chengannur", "Alathur", "Kottayam", "Kattapana",
        "Muvattupuzha", "Thiruvalla", "Pathanamthitta", "HO KKM"
    ].sort();

    // --- Column Headers Mapping (IMPORTANT: These must EXACTLY match the column names in your "Form Responses 2" Google Sheet) ---
    const HEADER_TIMESTAMP = 'Timestamp';
    const HEADER_DATE = 'Date';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_ACTIVITY_TYPE = 'Activity Type';
    const HEADER_TYPE_OF_CUSTOMER = 'Type of Customer';
    const HEADER_R_LEAD_SOURCE = 'rLead Source';
    const HEADER_HOW_CONTACTED = 'How Contacted';
    const HEADER_PROSPECT_NAME = 'Prospect Name';
    const HEADER_PHONE_NUMBER_WHATSAPP = 'Phone Numebr(Whatsapp)';
    const HEADER_ADDRESS = 'Address';
    const HEADER_PROFESSION = 'Profession';
    const HEADER_DOB_WD = 'DOB/WD';
    const HEADER_PRODUCT_INTERESTED = 'Prodcut Interested';
    const HEADER_REMARKS = 'Remarks';
    const HEADER_NEXT_FOLLOW_UP_DATE = 'Next Follow-up Date';
    const HEADER_RELATION_WITH_STAFF = 'Relation With Staff';


    // *** DOM Elements ***
    const branchSelect = document.getElementById('branchSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const employeeSelect = document.getElementById('employeeSelect');
    const dateFilterPanel = document.getElementById('dateFilterPanel');
    const dateFilter = document.getElementById('dateFilter');

    // Main Report Display Area (now within tab contents)
    const allBranchSnapshotTableBody = document.getElementById('branchSnapshotTableBody');

    // All Staff Performance (Overall) specific elements
    const employeeSummaryViewBtn = document.getElementById('employeeSummaryViewBtn');
    const customerActivityLogBtn = document.getElementById('customerActivityLogBtn');
    const employeePerformanceView = document.getElementById('employeePerformanceView');
    const customerActivityLogView = document.getElementById('customerActivityLogView');

    const staffOverallPerformanceTableBody = document.getElementById('staffOverallPerformanceTableBody'); // For employee summary
    const overallPerformanceTableBody = document.getElementById('overallPerformanceTableBody'); // For customer activity log

    const customerDetailView = document.getElementById('customerDetailView');
    const backToOverallListBtn = document.getElementById('backToOverallListBtn');

    // Specific detail elements for the customer report
    const detailDate = document.getElementById('detailDate');
    const detailBranchName = document.getElementById('detailBranchName');
    const detailEmployeeName = document.getElementById('detailEmployeeName');
    const detailDesignation = document.getElementById('detailDesignation');
    const detailCustomerName = document.getElementById('detailCustomerName');
    const detailAddress = document.getElementById('detailAddress');
    const detailPhoneNumber = document.getElementById('detailPhoneNumber');
    const detailEmailId = document.getElementById('detailEmailId');
    const detailRemarks = document.getElementById('detailRemarks');
    const detailFollowupDate = document.getElementById('detailFollowupDate');


    // NEW: Dedicated message area element
    const statusMessageDiv = document.getElementById('statusMessage');


    // Tab buttons for main navigation
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    // Main Content Sections to toggle
    const reportsSection = document.getElementById('reportsSection');
    const allBranchSnapshotTabContent = document.getElementById('allBranchSnapshotTabContent');
    const allStaffOverallPerformanceTabContent = document.getElementById('allStaffOverallPerformanceTabContent');
    const nonParticipatingBranchesTabContent = document.getElementById('nonParticipatingBranchesTabContent');
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    const noParticipationMessage = document.getElementById('noParticipationMessage');
    const nonParticipatingBranchesList = document.getElementById('nonParticipatingBranchesList');


    // Employee Management Form Elements
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const employeeNameInput = document.getElementById('employeeName');
    const employeeCodeInput = document.getElementById('employeeCode');
    const employeeBranchInput = document.getElementById('employeeBranch');
    const employeeDesignationInput = document.getElementById('employeeDesignation');
    const employeeManagementMessage = document.getElementById('employeeManagementMessage');

    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsTextarea = document.getElementById('bulkEmployeeDetails');

    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');


    // Global variables to store fetched data
    let allCanvassingData = []; // Raw activity data from Form Responses 2
    let allUniqueBranches = []; // Will be populated from PREDEFINED_BRANCHES
    let allUniqueEmployees = []; // Employee codes from Canvassing Data
    let employeeCodeToNameMap = {}; // {code: name} from Canvassing Data
    let employeeCodeToBranchMap = {}; // {code: branch} from Canvassing Data
    let employeeCodeToDesignationMap = {}; // {code: designation} from Canvassing Data

    let currentStaffSubView = 'employeeSummary'; // 'employeeSummary' or 'customerActivity'


    // Utility to format date to ISO-MM-DD
    const formatDate = (dateString) => {
        if (!dateString) return '';
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString; // Return original if invalid date
        return date.toISOString().split('T')[0];
    };

    // Helper to clear and display messages in a specific div
    function displayMessage(message, type = 'info') {
        if (statusMessageDiv) {
            statusMessageDiv.innerHTML = `<div class="message ${type}">${message}</div>`;
            statusMessageDiv.style.display = 'block';
            setTimeout(() => {
                statusMessageDiv.innerHTML = ''; // Clear message
                statusMessageDiv.style.display = 'none';
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
                employeeManagementMessage.textContent = ''; // Clear content
            }, 5000);
        }
    }

    // Function to fetch activity data from Google Sheet (Form Responses 2)
    async function fetchCanvassingData() {
        displayMessage("Fetching activity data...", 'info');
        try {
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                const errorText = await response.text();
                console.error(`HTTP error fetching Canvassing Data! Status: ${response.status}. Details: ${errorText}`);
                throw new Error(`Failed to fetch canvassing data. Status: ${response.status}. Please check DATA_URL.`);
            }
            const csvText = await response.text();
            allCanvassingData = parseCSV(csvText);
            console.log('--- Fetched Canvassing Data: ---');
            console.log(allCanvassingData); // Log canvassing data for debugging
            if (allCanvassingData.length > 0) {
                console.log('Canvassing Data Headers (first entry):', Object.keys(allCanvassingData[0]));
            }
            displayMessage("Activity data loaded successfully!", 'success');
        } catch (error) {
            console.error('Error fetching canvassing data:', error);
            displayMessage(`Failed to load activity data: ${error.message}. Please ensure the sheet is published correctly to CSV and the URL is accurate.`, 'error');
            allCanvassingData = [];
        }
    }

    // CSV parsing function (handles commas within quoted strings)
    function parseCSV(csv) {
        const lines = csv.split('\n').filter(line => line.trim() !== '');
        if (lines.length === 0) return [];

        const headers = parseCSVLine(lines[0]); // Headers can also contain commas in quotes
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = parseCSVLine(lines[i]);
            if (values.length !== headers.length) {
                console.warn(`Skipping malformed row ${i + 1}: Expected ${headers.length} columns, got ${values.length}. Line: "${lines[i]}"`);
                continue;
            }
            const entry = {};
            headers.forEach((header, index) => {
                entry[header] = values[index];
            });
            data.push(entry);
        }
        return data;
    }

    // Helper to parse a single CSV line safely
    function parseCSVLine(line) {
        const result = [];
        let inQuote = false;
        let currentField = '';
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                inQuote = !inQuote;
            } else if (char === ',' && !inQuote) {
                result.push(currentField.trim());
                currentField = '';
            } else {
                currentField += char;
            }
        }
        result.push(currentField.trim());
        return result;
    }


    // Process fetched data to populate filters and prepare for reports
    async function processData() {
        await fetchCanvassingData();

        allUniqueBranches = [...PREDEFINED_BRANCHES].sort();

        employeeCodeToNameMap = {};
        employeeCodeToBranchMap = {};
        employeeCodeToDesignationMap = {};
        allCanvassingData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const designation = entry[HEADER_DESIGNATION];
            const branchName = entry[HEADER_BRANCH_NAME];

            if (employeeCode) {
                employeeCodeToNameMap[employeeCode] = employeeName || employeeCode;
                employeeCodeToDesignationMap[employeeCode] = designation || 'Default';
                employeeCodeToBranchMap[employeeCode] = branchName || 'N/A';
            }
        });

        // Filter out employees without a code or name for unique employee list
        allUniqueEmployees = [...new Set(allCanvassingData
            .filter(entry => entry[HEADER_EMPLOYEE_CODE] && employeeCodeToNameMap[entry[HEADER_EMPLOYEE_CODE]])
            .map(entry => entry[HEADER_EMPLOYEE_CODE]))
        ].sort((codeA, codeB) => {
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        });


        populateDropdown(branchSelect, allUniqueBranches);
        populateDropdown(employeeSelect, allUniqueEmployees, true); // Populate employee dropdown

        console.log('Final All Unique Branches (Predefined):', allUniqueBranches);
        console.log('Final Employee Code To Name Map (from Canvassing Data):', employeeCodeToNameMap);
        console.log('Final Employee Code To Designation Map (from Canvassing Data):', employeeCodeToDesignationMap);
        console.log('Final Employee Code To Branch Map (from Canvassing Data):', employeeCodeToBranchMap);
        console.log('Final All Unique Employees (Codes from Canvassing Data):', allUniqueEmployees);

        // After data is loaded and maps are populated, render the initial report
        renderAllBranchSnapshot(); // Still render this by default on load.
    }

    // Populate dropdown utility
    function populateDropdown(selectElement, items, useCodeForValue = false) {
        selectElement.innerHTML = '<option value="">-- Select --</option>'; // Default option
        items.forEach(item => {
            const option = document.createElement('option');
            if (useCodeForValue) {
                // Display name from map or code itself
                option.value = item; // item is employeeCode
                option.textContent = employeeCodeToNameMap[item] || item;
            } else {
                option.value = item; // item is branch name
                option.textContent = item;
            }
            selectElement.appendChild(option);
        });
    }

    // --- Filtering Logic ---
    function getFilteredCanvassingData() {
        let filteredData = allCanvassingData;
        const selectedBranch = branchSelect.value;
        const selectedEmployeeCode = employeeSelect.value;
        const selectedDate = dateFilter.value; // YYYY-MM-DD format from input

        if (selectedBranch) {
            filteredData = filteredData.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch);
        }
        if (selectedEmployeeCode) {
            filteredData = filteredData.filter(entry => entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode);
        }
        if (selectedDate) {
            filteredData = filteredData.filter(entry => formatDate(entry[HEADER_DATE]) === selectedDate);
        }
        return filteredData;
    }

    // Event listeners for filters
    branchSelect.addEventListener('change', () => {
        // When branch changes, re-populate employee dropdown for that branch
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            const employeesInBranch = allCanvassingData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]);
            const uniqueEmployeesInBranch = [...new Set(employeesInBranch)].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });
            populateDropdown(employeeSelect, uniqueEmployeesInBranch, true);
            employeeFilterPanel.style.display = 'block';
        } else {
            // If no branch selected, show all employees
            populateDropdown(employeeSelect, allUniqueEmployees, true);
            employeeFilterPanel.style.display = 'block'; // Keep it visible for "All Staff" tab
        }
        // Re-render the active staff performance sub-view
        const activeTabButton = document.querySelector('.tab-button.active');
        if (activeTabButton.id === 'allStaffOverallPerformanceTabBtn') {
            renderActiveStaffSubView();
        } else if (activeTabButton.id === 'allBranchSnapshotTabBtn') {
            renderAllBranchSnapshot();
        }
    });

    employeeSelect.addEventListener('change', () => {
        const activeTabButton = document.querySelector('.tab-button.active');
        if (activeTabButton.id === 'allStaffOverallPerformanceTabBtn') {
            renderActiveStaffSubView();
        } else if (activeTabButton.id === 'allBranchSnapshotTabBtn') {
            renderAllBranchSnapshot();
        }
    });

    dateFilter.addEventListener('change', () => {
        const activeTabButton = document.querySelector('.tab-button.active');
        if (activeTabButton.id === 'allStaffOverallPerformanceTabBtn') {
            renderActiveStaffSubView();
        }
    });


    // Helper to calculate total activity from a set of activity entries based on Activity Type
    function calculateTotalActivity(entries) {
        const totalActivity = { 'Visit': 0, 'Call': 0, 'Reference': 0, 'New Customer Leads': 0 }; // Initialize counters

        entries.forEach(entry => {
            let activityType = entry[HEADER_ACTIVITY_TYPE];
            let typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER];

            const trimmedActivityType = activityType ? activityType.trim().toLowerCase() : '';
            const trimmedTypeOfCustomer = typeOfCustomer ? typeOfCustomer.trim().toLowerCase() : '';

            if (trimmedActivityType === 'visit') {
                totalActivity['Visit']++;
            } else if (trimmedActivityType === 'calls') {
                totalActivity['Call']++;
            } else if (trimmedActivityType === 'referance') {
                totalActivity['Reference']++;
            }

            if (trimmedTypeOfCustomer === 'new') { // Assuming 'New' in 'Type of Customer' indicates a new lead
                totalActivity['New Customer Leads']++;
            }
        });
        return totalActivity;
    }

    // Render All Branch Snapshot
    function renderAllBranchSnapshot() {
        allBranchSnapshotTableBody.innerHTML = ''; // Clear existing content
        const filteredData = getFilteredCanvassingData(); // Use filters for branch snapshot

        const branchesToReport = branchSelect.value ? [branchSelect.value] : PREDEFINED_BRANCHES;

        branchesToReport.forEach(branch => {
            const branchActivityEntries = filteredData.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
            const totalActivity = calculateTotalActivity(branchActivityEntries);
            const employeeCodesInBranch = [...new Set(branchActivityEntries.map(entry => entry[HEADER_EMPLOYEE_CODE]))];
            const displayEmployeeCount = employeeCodesInBranch.length;

            const row = allBranchSnapshotTableBody.insertRow();
            row.insertCell().setAttribute('data-label', 'Branch Name');
            row.lastChild.textContent = branch;

            row.insertCell().setAttribute('data-label', 'Employees with Activity');
            row.lastChild.textContent = displayEmployeeCount;

            row.insertCell().setAttribute('data-label', 'Total Visits');
            row.lastChild.textContent = totalActivity['Visit'];

            row.insertCell().setAttribute('data-label', 'Total Calls');
            row.lastChild.textContent = totalActivity['Call'];

            row.insertCell().setAttribute('data-label', 'Total References');
            row.lastChild.textContent = totalActivity['Reference'];

            row.insertCell().setAttribute('data-label', 'Total New Customer Leads');
            row.lastChild.textContent = totalActivity['New Customer Leads'];
        });
    }

    // Render Non-Participating Branches Report
    function renderNonParticipatingBranches() {
        nonParticipatingBranchesList.innerHTML = '';
        noParticipationMessage.style.display = 'none';

        const filteredData = getFilteredCanvassingData(); // Apply filters if any

        const nonParticipatingBranches = [];
        PREDEFINED_BRANCHES.forEach(branch => {
            // Check if ANY activity exists for this branch in the filtered data
            const hasActivity = filteredData.some(entry => entry[HEADER_BRANCH_NAME] === branch);
            if (!hasActivity) {
                nonParticipatingBranches.push(branch);
            }
        });

        if (nonParticipatingBranches.length > 0) {
            nonParticipatingBranches.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                nonParticipatingBranchesList.appendChild(li);
            });
        } else {
            noParticipationMessage.style.display = 'block';
        }
    }


    // --- All Staff Performance (Overall) Sub-Views ---

    // Render Employee Performance Summary
    function renderEmployeePerformanceSummary() {
        staffOverallPerformanceTableBody.innerHTML = ''; // Clear previous entries

        const filteredData = getFilteredCanvassingData();

        // Get unique employees from the filtered data, or all unique employees if no specific employee is filtered
        const employeesToReport = employeeSelect.value ? [employeeSelect.value] : allUniqueEmployees;

        if (filteredData.length === 0 || employeesToReport.length === 0) {
            const row = staffOverallPerformanceTableBody.insertRow();
            const cell = row.insertCell();
            cell.colSpan = 15; // Span across all columns
            cell.textContent = 'No employee canvassing data found for the selected filters.';
            cell.classList.add('no-activity-message');
            return;
        }

        employeesToReport.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const employeeDesignation = employeeCodeToDesignationMap[employeeCode] || 'Default';
            const employeeBranch = employeeCodeToBranchMap[employeeCode] || 'N/A';

            // Filter entries specifically for this employee
            const employeeEntries = filteredData.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);

            // If no entries for this employee under current filters, skip them
            if (employeeEntries.length === 0 && employeeSelect.value && employeeSelect.value === employeeCode) {
                 // If a specific employee is selected and they have no data, show message
                const row = staffOverallPerformanceTableBody.insertRow();
                const cell = row.insertCell();
                cell.colSpan = 15;
                cell.textContent = `No activity found for employee ${employeeName} under the current filters.`;
                cell.classList.add('no-activity-message');
                return;
            } else if (employeeEntries.length === 0 && !employeeSelect.value) {
                // If no specific employee selected, and current employee has no data, skip showing them in the list.
                // This prevents showing employees with 0 actuals if no activity for them in the filtered data.
                return;
            }

            const actuals = calculateTotalActivity(employeeEntries);
            const targets = TARGETS[employeeDesignation] || TARGETS['Default'];

            const row = staffOverallPerformanceTableBody.insertRow();

            // Employee Info
            row.insertCell().setAttribute('data-label', 'Employee Name');
            row.lastChild.textContent = employeeName;
            row.insertCell().setAttribute('data-label', 'Designation');
            row.lastChild.textContent = employeeDesignation;
            row.insertCell().setAttribute('data-label', 'Branch');
            row.lastChild.textContent = employeeBranch;

            // Visit Metrics
            row.insertCell().setAttribute('data-label', 'Target Visits');
            row.lastChild.textContent = targets['Visit'];
            row.insertCell().setAttribute('data-label', 'Actual Visits');
            row.lastChild.textContent = actuals['Visit'];
            const visitProgressCell = row.insertCell();
            visitProgressCell.setAttribute('data-label', 'Visit Progress');
            visitProgressCell.appendChild(createProgressBar(actuals['Visit'], targets['Visit']));

            // Call Metrics
            row.insertCell().setAttribute('data-label', 'Target Calls');
            row.lastChild.textContent = targets['Call'];
            row.insertCell().setAttribute('data-label', 'Actual Calls');
            row.lastChild.textContent = actuals['Call'];
            const callProgressCell = row.insertCell();
            callProgressCell.setAttribute('data-label', 'Call Progress');
            callProgressCell.appendChild(createProgressBar(actuals['Call'], targets['Call']));

            // Reference Metrics
            row.insertCell().setAttribute('data-label', 'Target References');
            row.lastChild.textContent = targets['Reference'];
            row.insertCell().setAttribute('data-label', 'Actual References');
            row.lastChild.textContent = actuals['Reference'];
            const referenceProgressCell = row.insertCell();
            referenceProgressCell.setAttribute('data-label', 'Reference Progress');
            referenceProgressCell.appendChild(createProgressBar(actuals['Reference'], targets['Reference']));

            // New Leads Metrics
            row.insertCell().setAttribute('data-label', 'Target New Leads');
            row.lastChild.textContent = targets['New Customer Leads'];
            row.insertCell().setAttribute('data-label', 'Actual New Leads');
            row.lastChild.textContent = actuals['New Customer Leads'];
            const newLeadProgressCell = row.insertCell();
            newLeadProgressCell.setAttribute('data-label', 'New Lead Progress');
            newLeadProgressCell.appendChild(createProgressBar(actuals['New Customer Leads'], targets['New Customer Leads']));
        });
    }

    // Render Customer Activity Log
    function renderCustomerActivityLog() {
        overallPerformanceTableBody.innerHTML = ''; // Clear previous entries
        customerDetailView.style.display = 'none'; // Ensure detail view is hidden
        // overallPerformanceListView.style.display = 'block'; // Make sure list is visible

        const filteredData = getFilteredCanvassingData();

        if (filteredData.length === 0) {
            const row = overallPerformanceTableBody.insertRow();
            const cell = row.insertCell();
            cell.colSpan = 5; // Span across all columns in this table
            cell.textContent = 'No customer activity entries found for the selected filters.';
            cell.classList.add('no-activity-message');
            return;
        }

        filteredData.forEach((entry, index) => {
            const row = overallPerformanceTableBody.insertRow();
            const dateCell = row.insertCell();
            const employeeNameCell = row.insertCell();
            const branchCell = row.insertCell();
            const customerNameCell = row.insertCell();
            const activityTypeCell = row.insertCell();

            dateCell.textContent = formatDate(entry[HEADER_DATE]);
            dateCell.setAttribute('data-label', 'Date');

            employeeNameCell.textContent = entry[HEADER_EMPLOYEE_NAME] || 'N/A';
            employeeNameCell.setAttribute('data-label', 'Employee Name');

            branchCell.textContent = entry[HEADER_BRANCH_NAME] || 'N/A';
            branchCell.setAttribute('data-label', 'Branch');

            const customerNameLink = document.createElement('a');
            customerNameLink.href = '#';
            customerNameLink.textContent = entry[HEADER_PROSPECT_NAME] || 'N/A';
            customerNameLink.classList.add('customer-name-link');
            customerNameLink.dataset.entryIndex = allCanvassingData.indexOf(entry); // Use original index for retrieval

            customerNameLink.addEventListener('click', (e) => {
                e.preventDefault();
                const clickedIndex = parseInt(e.target.dataset.entryIndex);
                if (!isNaN(clickedIndex) && allCanvassingData[clickedIndex]) {
                    renderCustomerDetailReport(allCanvassingData[clickedIndex]);
                }
            });

            customerNameCell.appendChild(customerNameLink);
            customerNameCell.setAttribute('data-label', 'Customer Name');

            activityTypeCell.textContent = entry[HEADER_ACTIVITY_TYPE] || 'N/A';
            activityTypeCell.setAttribute('data-label', 'Activity Type');
        });
    }

    // Render Detailed Customer Report
    function renderCustomerDetailReport(entry) {
        customerActivityLogView.querySelector('#overallPerformanceListView').style.display = 'none'; // Hide the list view
        customerDetailView.style.display = 'block'; // Show the detail view

        detailDate.textContent = formatDate(entry[HEADER_DATE]);
        detailBranchName.textContent = entry[HEADER_BRANCH_NAME] || 'N/A';
        detailEmployeeName.textContent = entry[HEADER_EMPLOYEE_NAME] || 'N/A';
        detailDesignation.textContent = entry[HEADER_DESIGNATION] || 'N/A';
        detailCustomerName.textContent = entry[HEADER_PROSPECT_NAME] || 'N/A';
        detailAddress.textContent = entry[HEADER_ADDRESS] || 'N/A';
        detailPhoneNumber.textContent = entry[HEADER_PHONE_NUMBER_WHATSAPP] || 'N/A';
        detailEmailId.textContent = entry['Email ID'] || 'N/A'; // Assuming 'Email ID' is the header for email if it exists
        detailRemarks.textContent = entry[HEADER_REMARKS] || 'N/A';
        detailFollowupDate.textContent = formatDate(entry[HEADER_NEXT_FOLLOW_UP_DATE]);
    }

    // Event Listener for "Back to List" button for the detailed customer report
    if (backToOverallListBtn) {
        backToOverallListBtn.addEventListener('click', () => {
            customerActivityLogView.querySelector('#overallPerformanceListView').style.display = 'block';
            customerDetailView.style.display = 'none';
        });
    }


    // Helper to create a progress bar element
    function createProgressBar(actual, target) {
        const progressBarContainer = document.createElement('div');
        progressBarContainer.classList.add('progress-bar-container');

        const progressBar = document.createElement('div');
        progressBar.classList.add('progress-bar');

        let percentage = 0;
        if (target > 0) {
            percentage = (actual / target) * 100;
        } else if (actual > 0) {
            // If no target but actuals exist, consider 100% or higher, capped at 100% visual width
            percentage = 100;
        }

        progressBar.style.width = `${Math.min(100, percentage)}%`; // Cap at 100% width
        progressBar.classList.add(getProgressBarClass(percentage));

        const progressText = document.createElement('span');
        progressText.classList.add('progress-text');
        progressText.textContent = `${percentage.toFixed(0)}%`;

        progressBarContainer.appendChild(progressBar);
        progressBarContainer.appendChild(progressText);

        return progressBarContainer;
    }

    // Helper for progress bar styling
    function getProgressBarClass(percentage) {
        if (isNaN(percentage)) return 'no-activity';
        const p = parseFloat(percentage);
        if (p >= 100) return 'overachieved';
        if (p >= 75) return 'success';
        if (p >= 50) return 'warning';
        return 'danger';
    }

    // Function to send data to Google Apps Script (retained original functionality)
    async function sendDataToGoogleAppsScript(actionType, data = {}) {
        displayEmployeeManagementMessage('Processing request...', false);

        try {
            const formData = new URLSearchParams();
            formData.append('actionType', actionType);
            formData.append('data', JSON.stringify(data));

            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                },
                body: formData
            });

            if (!response.ok) {
                const errorText = await response.text();
                console.error(`HTTP error from Apps Script! Status: ${response.status}. Details: ${errorText}`);
                throw new Error(`Failed to send data to Apps Script. Status: ${response.status}. Please check WEB_APP_URL and Apps Script deployment.`);
            }

            const result = await response.json();

            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(result.message, false);
                return true;
            } else {
                displayEmployeeManagementMessage(`Error: ${result.message}`, true);
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Apps Script:', error);
            displayEmployeeManagementMessage(`Error sending data: ${error.message}. Please check WEB_APP_URL and Apps Script deployment.`, true);
            return false;
        } finally {
            // Re-fetch all data to ensure reports are up-to-date after any employee management action
            await processData(); // Re-fetch canvassing data and re-populate maps/dropdowns
            // Re-render the current report or provide a message
            const activeTabButton = document.querySelector('.tab-button.active');
            if (activeTabButton) {
                if (activeTabButton.id === 'allBranchSnapshotTabBtn') {
                    renderAllBranchSnapshot();
                } else if (activeTabButton.id === 'allStaffOverallPerformanceTabBtn') {
                    renderActiveStaffSubView(); // Re-render the currently active sub-view
                } else if (activeTabButton.id === 'nonParticipatingBranchesTabBtn') {
                    renderNonParticipatingBranches();
                }
                // No re-render for Employee Management tab, just clear messages
            }
        }
    }


    // Function to manage tab visibility
    function showTab(tabButtonId) {
        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });
        // Activate the clicked tab button
        document.getElementById(tabButtonId).classList.add('active');

        // Hide all main content sections
        allBranchSnapshotTabContent.style.display = 'none';
        allStaffOverallPerformanceTabContent.style.display = 'none';
        nonParticipatingBranchesTabContent.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        // Hide/Show filtering controls based on tab
        const controlsPanel = document.querySelector('.controls-panel');
        employeeFilterPanel.style.display = 'none'; // Start hidden
        dateFilterPanel.style.display = 'none'; // Start hidden

        // Reset dropdowns if we are switching away from a branch/employee dependent tab
        branchSelect.value = '';
        employeeSelect.value = '';
        dateFilter.value = ''; // Reset date filter too

        // Show relevant content based on tab selection
        if (tabButtonId === 'allBranchSnapshotTabBtn') {
            reportsSection.style.display = 'block';
            allBranchSnapshotTabContent.style.display = 'block';
            controlsPanel.style.display = 'flex'; // Show controls
            employeeFilterPanel.style.display = 'block'; // Show employee filter
            dateFilterPanel.style.display = 'none'; // Hide date filter for this tab
            populateDropdown(employeeSelect, allUniqueEmployees, true); // Reset employee dropdown to all
            renderAllBranchSnapshot();
        } else if (tabButtonId === 'allStaffOverallPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            allStaffOverallPerformanceTabContent.style.display = 'block';
            controlsPanel.style.display = 'flex'; // Show controls
            employeeFilterPanel.style.display = 'block'; // Show employee filter
            dateFilterPanel.style.display = 'block'; // Show date filter for this tab
            populateDropdown(employeeSelect, allUniqueEmployees, true); // Reset employee dropdown to all

            // Initialize sub-view to employee summary
            currentStaffSubView = 'employeeSummary';
            employeeSummaryViewBtn.classList.add('active');
            customerActivityLogBtn.classList.remove('active');
            employeePerformanceView.style.display = 'block';
            customerActivityLogView.style.display = 'none';
            customerDetailView.style.display = 'none'; // Ensure detail view is hidden
            
            renderEmployeePerformanceSummary(); // Render the default view for this tab
        } else if (tabButtonId === 'nonParticipatingBranchesTabBtn') {
            reportsSection.style.display = 'block';
            nonParticipatingBranchesTabContent.style.display = 'block';
            controlsPanel.style.display = 'flex'; // Show controls (branch filter is useful here)
            employeeFilterPanel.style.display = 'none'; // Employee filter not directly relevant here
            dateFilterPanel.style.display = 'none'; // Date filter not directly relevant here
            renderNonParticipatingBranches();
        } else if (tabButtonId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
            controlsPanel.style.display = 'none'; // Hide controls for this tab
            displayEmployeeManagementMessage('', false); // Clear messages
        }
    }

    // Function to render the currently active sub-view within All Staff Performance (Overall)
    function renderActiveStaffSubView() {
        if (currentStaffSubView === 'employeeSummary') {
            renderEmployeePerformanceSummary();
        } else if (currentStaffSubView === 'customerActivity') {
            renderCustomerActivityLog();
        }
    }

    // Event listeners for main tab buttons
    if (allBranchSnapshotTabBtn) {
        allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    }
    if (allStaffOverallPerformanceTabBtn) {
        allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    }
    if (nonParticipatingBranchesTabBtn) {
        nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    }
    if (employeeManagementTabBtn) {
        employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    }

    // Event listeners for sub-view toggle buttons
    if (employeeSummaryViewBtn) {
        employeeSummaryViewBtn.addEventListener('click', () => {
            currentStaffSubView = 'employeeSummary';
            employeeSummaryViewBtn.classList.add('active');
            customerActivityLogBtn.classList.remove('active');
            employeePerformanceView.style.display = 'block';
            customerActivityLogView.style.display = 'none';
            customerDetailView.style.display = 'none'; // Hide detail view when switching
            renderEmployeePerformanceSummary();
        });
    }

    if (customerActivityLogBtn) {
        customerActivityLogBtn.addEventListener('click', () => {
            currentStaffSubView = 'customerActivity';
            customerActivityLogBtn.classList.add('active');
            employeeSummaryViewBtn.classList.remove('active');
            employeePerformanceView.style.display = 'none';
            customerActivityLogView.style.display = 'block';
            renderCustomerActivityLog(); // This will automatically hide detail view
        });
    }

    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();

            const employeeName = employeeNameInput.value.trim();
            const employeeCode = employeeCodeInput.value.trim();
            const branchName = employeeBranchInput.value.trim();
            const designation = employeeDesignationInput.value.trim();

            if (!employeeName || !employeeCode || !branchName) {
                displayEmployeeManagementMessage('Please fill in Employee Name, Code, and Branch Name.', true);
                return;
            }

            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: employeeName,
                [HEADER_EMPLOYEE_CODE]: employeeCode,
                [HEADER_BRANCH_NAME]: branchName,
                [HEADER_DESIGNATION]: designation
            };

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);

            if (success) {
                addEmployeeForm.reset();
            }
        });
    }

    // Event Listener for Bulk Add Employee Form
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
                if (parts.length < 2) {
                    displayEmployeeManagementMessage(`Skipping invalid line: "${line}". Each line must have at least Employee Name and Employee Code.`, true);
                    continue;
                }

                const employeeData = {
                    [HEADER_EMPLOYEE_NAME]: parts[0],
                    [HEADER_EMPLOYEE_CODE]: parts[1],
                    [HEADER_BRANCH_NAME]: branchName,
                    [HEADER_DESIGNATION]: parts[2] || ''
                };
                employeesToAdd.push(employeeData);
            }

            if (employeesToAdd.length > 0) {
                const success = await sendDataToGoogleAppsScript('add_bulk_employees', employeesToAdd);
                if (success) {
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
                deleteEmployeeForm.reset();
            }
        });
    }

    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn'); // Default tab to display on load
});
