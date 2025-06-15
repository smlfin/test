document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    // This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv";

    // IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE

    // NEW: URL for your published MasterEmployees sheet (as CSV).
    // THIS URL MUST COME FROM THE SPREADSHEET WITH ID '1Za1CrlzzXpQjB3yZHjL2ZpRkjXgkVmLHH_LtXJq9K5o'
    // THE SPREADSHEET THAT YOUR APPS SCRIPT (code.gs) IS UPDATING.
    const EMPLOYEE_MASTER_DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTQg3wL9eQf1sY6x7zB5bK8dC5oV0XpQjB3yZHjL2ZpRkjXgkVmLHH_LtXJq9K5o/pub?gid=0&single=true&output=csv"; // Example URL, ensure yours is correct

    // --- Header Constants ---
    // These must EXACTLY match your column headers in the Google Sheet CSV outputs
    const HEADER_BRANCH_NAME = "Branch Name";
    const HEADER_EMPLOYEE_NAME = "Employee Name";
    const HEADER_EMPLOYEE_CODE = "Employee Code";
    const HEADER_DESIGNATION = "Designation";
    const HEADER_CUSTOMER_NAME = "Customer Name";
    const HEADER_LAST_CONTACT = "Last Contact";
    const HEADER_PRODUCTS_DISCUSSED = "Products Discussed";
    const HEADER_STATUS = "Status";
    const HEADER_REMARKS = "Remarks";
    const HEADER_FAMILY_DETAILS = "Family Details"; // Assuming a column for family details
    const HEADER_CUSTOMER_PROFILE = "Customer Profile"; // Assuming a column for customer profile


    // *** DOM Elements ***
    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const reportDisplay = document.getElementById('reportDisplay');
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');
    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection');
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    const statusMessageContainer = document.getElementById('statusMessage');

    const customerViewBranchSelect = document.getElementById('customerViewBranchSelect');
    const customerViewEmployeeSelect = document.getElementById('customerViewEmployeeSelect');
    const customerCanvassedList = document.getElementById('customerCanvassedList');
    const customerOverviewSection = document.getElementById('customerOverviewSection');
    const canvassingActivitySection = document.getElementById('canvassingActivitySection');
    const remarksSection = document.getElementById('remarksSection');
    const familyDetailsSection = document.getElementById('familyDetailsSection');
    const customerProfileSection = document.getElementById('customerProfileSection');

    // Employee Management DOM elements
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const newEmployeeNameInput = document.getElementById('newEmployeeName');
    const newEmployeeCodeInput = document.getElementById('newEmployeeCode');
    const newBranchNameInput = document.getElementById('newBranchName');
    const newDesignationInput = document.getElementById('newDesignation');
    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsInput = document.getElementById('bulkEmployeeDetails');
    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');
    const employeeManagementMessageDiv = document.getElementById('employeeManagementMessage');

    // --- Global Data Stores ---
    let fullCanvassingData = [];
    let employeeMasterData = [];
    let selectedCustomerData = null; // To store details of the currently selected customer

    // --- Helper Functions ---

    // Function to parse CSV text into an array of objects
    function parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        if (lines.length === 0) return [];

        const headers = lines[0].split(',').map(header => header.trim()); // Trim headers
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = lines[i].split(',').map(value => value.trim()); // Trim values
            // Skip empty lines or lines with incorrect number of columns
            if (values.length === 0 || values.every(v => v === '')) continue;
            if (values.length !== headers.length) {
                console.warn(`Skipping row ${i + 1} due to column mismatch. Expected ${headers.length}, got ${values.length}. Row: ${lines[i]}`);
                continue; // Skip malformed rows
            }
            const rowObject = {};
            headers.forEach((header, index) => {
                rowObject[header] = values[index];
            });
            data.push(rowObject);
        }
        return data;
    }

    // Function to fetch CSV data from a URL
    async function fetchCSV(url) {
        try {
            const response = await fetch(url);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const csvText = await response.text();
            return csvText;
        } catch (error) {
            console.error("Error fetching CSV:", error);
            displayStatusMessage(`Failed to load data: ${error.message}`, true);
            return null;
        }
    }

    // Function to display status messages
    function displayStatusMessage(message, isError = false) {
        statusMessageContainer.textContent = message;
        statusMessageContainer.className = `message-container message ${isError ? 'error' : 'success'}`;
        statusMessageContainer.style.display = 'block';

        // Hide message after a few seconds unless it's an error
        if (!isError) {
            setTimeout(() => {
                statusMessageContainer.style.display = 'none';
            }, 5000);
        }
    }

    // Function to display employee management messages
    function displayEmployeeManagementMessage(message, isError = false) {
        employeeManagementMessageDiv.textContent = message;
        employeeManagementMessageDiv.className = `message ${isError ? 'error' : 'success'}`;
        employeeManagementMessageDiv.style.display = 'block';

        setTimeout(() => {
            employeeManagementMessageDiv.style.display = 'none';
        }, 5000);
    }

    // Function to populate a dropdown
    function populateDropdown(selectElement, items, defaultText = null) {
        selectElement.innerHTML = ''; // Clear existing options
        const defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = defaultText || `-- Select a ${selectElement.id.includes('Branch') ? 'Branch' : 'Employee'} --`;
        selectElement.appendChild(defaultOption);

        items.forEach(item => {
            const option = document.createElement('option');
            option.value = item;
            option.textContent = item;
            selectElement.appendChild(option);
        });
    }

    // Function to show active tab content
    function showTab(tabId) {
        // Hide all sections
        reportsSection.style.display = 'none';
        detailedCustomerViewSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));

        // Show the selected section and activate its button
        switch (tabId) {
            case 'allBranchSnapshotTabBtn':
            case 'allStaffOverallPerformanceTabBtn':
            case 'nonParticipatingBranchesTabBtn':
                reportsSection.style.display = 'block';
                // Adjust display of employee filter based on report type
                employeeFilterPanel.style.display = (tabId === 'allStaffOverallPerformanceTabBtn') ? 'flex' : 'none';
                break;
            case 'detailedCustomerViewTabBtn':
                detailedCustomerViewSection.style.display = 'flex'; // Use flex for this section
                break;
            case 'employeeManagementTabBtn':
                employeeManagementSection.style.display = 'block';
                break;
        }
        document.getElementById(tabId).classList.add('active');
    }

    // Function to send data to Google Apps Script
    async function sendDataToGoogleAppsScript(action, data) {
        displayEmployeeManagementMessage('Processing request...', false);
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Required for cross-origin requests
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({ action: action, data: JSON.stringify(data) })
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Server error: ${response.status} - ${errorText}`);
            }

            const result = await response.json();
            if (result.status === 'success') {
                displayEmployeeManagementMessage(result.message || 'Operation successful!', false);
                await processData(); // Re-fetch data to update dropdowns
                return true;
            } else {
                displayEmployeeManagementMessage(result.message || 'Operation failed!', true);
                return false;
            }
        } catch (error) {
            console.error("Error sending data to Apps Script:", error);
            displayEmployeeManagementMessage(`Failed to complete operation: ${error.message}`, true);
            return false;
        }
    }


    // --- Core Data Processing ---
    async function processData() {
        displayStatusMessage('Loading data...', false);
        try {
            // Fetch employee master data
            const employeeCsv = await fetchCSV(EMPLOYEE_MASTER_DATA_URL);
            if (!employeeCsv) {
                displayStatusMessage('Failed to load employee master data. Check URL and publishing settings.', true);
                return;
            }
            employeeMasterData = parseCSV(employeeCsv);
            console.log("Employee Master Data:", employeeMasterData); // Debugging

            // Fetch canvassing data
            const canvassingCsv = await fetchCSV(DATA_URL);
            if (!canvassingCsv) {
                displayStatusMessage('Failed to load canvassing data. Check URL and publishing settings.', true);
                return;
            }
            fullCanvassingData = parseCSV(canvassingCsv);
            console.log("Full Canvassing Data:", fullCanvassingData); // Debugging


            // Extract unique branches and employees from Master Data for dropdowns
            const branches = new Set();
            const employees = new Set();
            employeeMasterData.forEach(row => {
                if (row[HEADER_BRANCH_NAME]) branches.add(row[HEADER_BRANCH_NAME]);
                if (row[HEADER_EMPLOYEE_NAME]) employees.add(row[HEADER_EMPLOYEE_NAME]);
            });

            // Populate main branch and employee dropdowns
            populateDropdown(branchSelect, Array.from(branches).sort());
            populateDropdown(employeeSelect, Array.from(employees).sort()); // Populate with all employees initially

            // Populate customer view branch and employee dropdowns
            populateDropdown(customerViewBranchSelect, Array.from(branches).sort());
            populateDropdown(customerViewEmployeeSelect, Array.from(employees).sort()); // Populate with all employees initially

            displayStatusMessage('Data loaded successfully!', false);

        } catch (error) {
            console.error("Error in processData:", error);
            displayStatusMessage(`Error processing data: ${error.message}`, true);
        }
    }

    // --- Report Generation Functions (Stubs for now) ---
    function generateAllBranchSnapshotReport() {
        let html = '<h2>All Branch Snapshot Report</h2>';
        const uniqueBranches = [...new Set(fullCanvassingData.map(d => d[HEADER_BRANCH_NAME]))];

        if (uniqueBranches.length === 0) {
            html += '<p>No branch data available.</p>';
            reportDisplay.innerHTML = html;
            return;
        }

        html += '<div class="data-table-container">';
        html += '<table class="all-branch-snapshot-table">';
        html += '<thead><tr><th>Branch Name</th><th>Total Customers Canvassed</th><th>Total Products Discussed</th><th>Average Success Rate (%)</th></tr></thead>';
        html += '<tbody>';

        uniqueBranches.forEach(branch => {
            const branchData = fullCanvassingData.filter(d => d[HEADER_BRANCH_NAME] === branch);
            const totalCustomers = new Set(branchData.map(d => d[HEADER_CUSTOMER_NAME])).size; // Count unique customers
            const totalProducts = branchData.length; // Assuming each row is a product discussion

            // Simple success rate calculation for demo: count 'Success' statuses
            const successfulCanvasses = branchData.filter(d => d[HEADER_STATUS] && d[HEADER_STATUS].toLowerCase() === 'success').length;
            const successRate = totalProducts > 0 ? ((successfulCanvasses / totalProducts) * 100).toFixed(2) : 0;

            html += `<tr>
                        <td data-label="Branch Name">${branch}</td>
                        <td data-label="Total Customers Canvassed">${totalCustomers}</td>
                        <td data-label="Total Products Discussed">${totalProducts}</td>
                        <td data-label="Average Success Rate (%)">
                            <div class="progress-bar-container-small">
                                <div class="progress-bar ${getProgressBarClass(parseFloat(successRate))}" style="width: ${successRate}%;">
                                    ${successRate}%
                                </div>
                            </div>
                        </td>
                    </tr>`;
        });

        html += '</tbody></table></div>';
        reportDisplay.innerHTML = html;
    }

    function generateAllStaffOverallPerformanceReport() {
        let html = '<h2>All Staff Performance (Overall) Report</h2>';
        const employees = [...new Set(employeeMasterData.map(e => e[HEADER_EMPLOYEE_NAME]))];

        if (employees.length === 0) {
            html += '<p>No employee data available.</p>';
            reportDisplay.innerHTML = html;
            return;
        }

        html += '<div class="data-table-container">';
        html += '<table class="performance-table">';
        html += '<thead>';
        html += '<tr>';
        html += '<th rowspan="2">Employee Name</th>';
        html += '<th rowspan="2">Branch</th>';
        html += '<th rowspan="2">Designation</th>';
        html += '<th colspan="3">Canvassing Performance (All Time)</th>';
        html += '</tr>';
        html += '<tr>';
        html += '<th>Customers Canvassed</th><th>Products Discussed</th><th>Conversion Rate (%)</th>';
        html += '</tr>';
        html += '</thead>';
        html += '<tbody>';

        employees.forEach(employeeName => {
            const employeeDetails = employeeMasterData.find(e => e[HEADER_EMPLOYEE_NAME] === employeeName);
            const branch = employeeDetails ? employeeDetails[HEADER_BRANCH_NAME] : 'N/A';
            const designation = employeeDetails ? employeeDetails[HEADER_DESIGNATION] : 'N/A';

            const employeeCanvassData = fullCanvassingData.filter(d => d[HEADER_EMPLOYEE_NAME] === employeeName);
            const totalCustomers = new Set(employeeCanvassData.map(d => d[HEADER_CUSTOMER_NAME])).size;
            const totalProductsDiscussed = employeeCanvassData.length;
            const convertedCustomers = employeeCanvassData.filter(d => d[HEADER_STATUS] && d[HEADER_STATUS].toLowerCase() === 'converted').length;
            const conversionRate = totalProductsDiscussed > 0 ? ((convertedCustomers / totalProductsDiscussed) * 100).toFixed(2) : 0;

            html += `<tr>
                        <td data-label="Employee Name">${employeeName}</td>
                        <td data-label="Branch">${branch}</td>
                        <td data-label="Designation">${designation}</td>
                        <td data-label="Customers Canvassed">${totalCustomers}</td>
                        <td data-label="Products Discussed">${totalProductsDiscussed}</td>
                        <td data-label="Conversion Rate (%)">
                            <div class="progress-bar-container-small">
                                <div class="progress-bar ${getProgressBarClass(parseFloat(conversionRate))}" style="width: ${conversionRate}%;">
                                    ${conversionRate}%
                                </div>
                            </div>
                        </td>
                    </tr>`;
        });

        html += '</tbody></table></div>';
        reportDisplay.innerHTML = html;
    }

    function generateNonParticipatingBranchesReport() {
        let html = '<h2>Non-Participating Branches Report</h2>';
        const allBranchesFromMaster = new Set(employeeMasterData.map(e => e[HEADER_BRANCH_NAME]));
        const participatingBranchesFromCanvassing = new Set(fullCanvassingData.map(d => d[HEADER_BRANCH_NAME]));

        const nonParticipatingBranches = [...allBranchesFromMaster].filter(branch =>
            !participatingBranchesFromCanvassing.has(branch)
        );

        if (nonParticipatingBranches.length === 0) {
            html += '<p class="no-participation-message">All branches have participated in canvassing.</p>';
        } else {
            html += '<p>The following branches have not recorded any canvassing activity:</p>';
            html += '<ul class="non-participating-branch-list">';
            nonParticipatingBranches.forEach(branch => {
                html += `<li>${branch}</li>`;
            });
            html += '</ul>';
        }
        reportDisplay.innerHTML = html;
    }

    // Example of a function to view branch performance for a selected branch (d3.PNG)
    function generateBranchPerformanceReport(branchName) {
        if (!branchName) {
            reportDisplay.innerHTML = '<p>Please select a branch to view its performance report.</p>';
            return;
        }

        let html = `<h2>Branch Performance Report: ${branchName}</h2>`;
        const branchEmployees = employeeMasterData.filter(e => e[HEADER_BRANCH_NAME] === branchName);

        if (branchEmployees.length === 0) {
            html += '<p>No employees found for this branch.</p>';
            reportDisplay.innerHTML = html;
            return;
        }

        html += '<div class="branch-performance-grid">';
        branchEmployees.forEach(employee => {
            const employeeName = employee[HEADER_EMPLOYEE_NAME];
            const employeeCanvassData = fullCanvassingData.filter(d => d[HEADER_EMPLOYEE_NAME] === employeeName);
            const totalCustomers = new Set(employeeCanvassData.map(d => d[HEADER_CUSTOMER_NAME])).size;
            const totalProducts = employeeCanvassData.length;
            const convertedCount = employeeCanvassData.filter(d => d[HEADER_STATUS] && d[HEADER_STATUS].toLowerCase() === 'converted').length;
            const conversionRate = totalProducts > 0 ? ((convertedCount / totalProducts) * 100).toFixed(2) : 0;

            html += `
                <div class="employee-performance-card">
                    <h4>${employeeName}</h4>
                    <table>
                        <tr><th>Customers Canvassed:</th><td>${totalCustomers}</td></tr>
                        <tr><th>Products Discussed:</th><td>${totalProducts}</td></tr>
                        <tr><th>Conversion Rate:</th><td>
                            <div class="progress-bar-container">
                                <div class="progress-bar ${getProgressBarClass(parseFloat(conversionRate))}" style="width: ${conversionRate}%;">
                                    ${conversionRate}%
                                </div>
                            </div>
                        </td></tr>
                    </table>
                </div>
            `;
        });
        html += '</div>';
        reportDisplay.innerHTML = html;
    }

    // Example of a function to view employee summary (d4.PNG)
    function generateEmployeeSummary(employeeName) {
        if (!employeeName) {
            reportDisplay.innerHTML = '<p>Please select an employee to view their summary.</p>';
            return;
        }

        let html = `<h2>Employee Summary: ${employeeName} (Current Month)</h2>`;
        const currentMonthData = fullCanvassingData.filter(d =>
            d[HEADER_EMPLOYEE_NAME] === employeeName &&
            new Date(d[HEADER_LAST_CONTACT]).getMonth() === new Date().getMonth() &&
            new Date(d[HEADER_LAST_CONTACT]).getFullYear() === new Date().getFullYear()
        );

        if (currentMonthData.length === 0) {
            html += '<p>No canvassing activity for this employee in the current month.</p>';
            reportDisplay.innerHTML = html;
            return;
        }

        const totalCustomers = new Set(currentMonthData.map(d => d[HEADER_CUSTOMER_NAME])).size;
        const totalProductsDiscussed = currentMonthData.length;
        const successfulCanvasses = currentMonthData.filter(d => d[HEADER_STATUS] && d[HEADER_STATUS].toLowerCase() === 'success').length;
        const pendingCanvasses = currentMonthData.filter(d => d[HEADER_STATUS] && d[HEADER_STATUS].toLowerCase() === 'pending').length;
        const failedCanvasses = currentMonthData.filter(d => d[HEADER_STATUS] && d[HEADER_STATUS].toLowerCase() === 'failed').length;

        const productInterest = {};
        currentMonthData.forEach(d => {
            const products = d[HEADER_PRODUCTS_DISCUSSED] ? d[HEADER_PRODUCTS_DISCUSSED].split(';').map(p => p.trim()) : [];
            products.forEach(p => {
                productInterest[p] = (productInterest[p] || 0) + 1;
            });
        });

        html += `
            <div class="summary-breakdown-card">
                <div>
                    <h4>Overall Activity</h4>
                    <ul class="summary-list">
                        <li><span>Customers Canvassed:</span> <strong>${totalCustomers}</strong></li>
                        <li><span>Products Discussed:</span> <strong>${totalProductsDiscussed}</strong></li>
                        <li><span>Successful Canvasses:</span> <strong>${successfulCanvasses}</strong></li>
                        <li><span>Pending Canvasses:</span> <strong>${pendingCanvasses}</strong></li>
                        <li><span>Failed Canvasses:</span> <strong>${failedCanvasses}</strong></li>
                    </ul>
                </div>
                <div>
                    <h4>Product Interest Breakdown</h4>
                    <ul class="product-interest-list">
                        ${Object.entries(productInterest).map(([product, count]) => `
                            <li><span>${product}:</span> <strong>${count}</strong></li>
                        `).join('')}
                    </ul>
                </div>
            </div>
        `;
        reportDisplay.innerHTML = html;
    }

    // Function to generate 'View All Entries' table for an employee
    function generateEmployeeAllEntriesReport(employeeName) {
        if (!employeeName) {
            reportDisplay.innerHTML = '<p>Please select an employee to view all entries.</p>';
            return;
        }

        let html = `<h2>All Canvassing Entries for: ${employeeName}</h2>`;
        const employeeEntries = fullCanvassingData.filter(d => d[HEADER_EMPLOYEE_NAME] === employeeName);

        if (employeeEntries.length === 0) {
            html += '<p>No canvassing entries found for this employee.</p>';
            reportDisplay.innerHTML = html;
            return;
        }

        html += '<div class="data-table-container">';
        html += '<table>';
        html += '<thead><tr><th>Customer Name</th><th>Last Contact</th><th>Products Discussed</th><th>Status</th><th>Remarks</th></tr></thead>';
        html += '<tbody>';

        employeeEntries.forEach(entry => {
            html += `<tr>
                        <td data-label="Customer Name">${entry[HEADER_CUSTOMER_NAME] || ''}</td>
                        <td data-label="Last Contact">${entry[HEADER_LAST_CONTACT] || ''}</td>
                        <td data-label="Products Discussed">${entry[HEADER_PRODUCTS_DISCUSSED] || ''}</td>
                        <td data-label="Status">${entry[HEADER_STATUS] || ''}</td>
                        <td data-label="Remarks">${entry[HEADER_REMARKS] || ''}</td>
                    </tr>`;
        });

        html += '</tbody></table></div>';
        reportDisplay.innerHTML = html;
    }


    // Function to view detailed customer information when a customer is selected from the list
    function displayCustomerDetails(customerName) {
        // Find the customer data. Assuming customerName is unique enough for demo
        selectedCustomerData = fullCanvassingData.find(d => d[HEADER_CUSTOMER_NAME] === customerName);

        if (!selectedCustomerData) {
            // Clear previous details if no customer found
            customerOverviewSection.innerHTML = '<h3>Customer Overview</h3><p>No customer selected or data not found.</p>';
            canvassingActivitySection.innerHTML = '<h3>Canvassing Activity</h3><p>No customer selected or data not found.</p>';
            familyDetailsSection.innerHTML = '<h3>Family Details</h3><p>No customer selected or data not found.</p>';
            customerProfileSection.innerHTML = '<h3>Customer Profile</h3><p>No customer selected or data not found.</p>';
            remarksSection.innerHTML = '<h3>Remarks</h3><p>No customer selected or data not found.</p>';
            return;
        }

        // Populate Customer Overview
        customerOverviewSection.innerHTML = `
            <h3>Customer Overview</h3>
            <div class="detail-row">
                <span class="detail-label">Name:</span>
                <span class="detail-value">${selectedCustomerData[HEADER_CUSTOMER_NAME] || 'N/A'}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Employee:</span>
                <span class="detail-value">${selectedCustomerData[HEADER_EMPLOYEE_NAME] || 'N/A'}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Branch:</span>
                <span class="detail-value">${selectedCustomerData[HEADER_BRANCH_NAME] || 'N/A'}</span>
            </div>
            `;

        // Populate Canvassing Activity
        canvassingActivitySection.innerHTML = `
            <h3>Canvassing Activity</h3>
            <div class="detail-row">
                <span class="detail-label">Last Contact:</span>
                <span class="detail-value">${selectedCustomerData[HEADER_LAST_CONTACT] || 'N/A'}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Products Discussed:</span>
                <span class="detail-value">${selectedCustomerData[HEADER_PRODUCTS_DISCUSSED] || 'N/A'}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Status:</span>
                <span class="detail-value">${selectedCustomerData[HEADER_STATUS] || 'N/A'}</span>
            </div>
            `;

        // Populate Family Details
        familyDetailsSection.innerHTML = `
            <h3>Family Details</h3>
            <p class="remark-text">${selectedCustomerData[HEADER_FAMILY_DETAILS] || 'No family details available.'}</p>
        `;

        // Populate Customer Profile
        customerProfileSection.innerHTML = `
            <h3>Customer Profile</h3>
            <p class="profile-text">${selectedCustomerData[HEADER_CUSTOMER_PROFILE] || 'No customer profile available.'}</p>
        `;

        // Populate Remarks
        remarksSection.innerHTML = `
            <h3>Remarks</h3>
            <p class="remark-text">${selectedCustomerData[HEADER_REMARKS] || 'No remarks available.'}</p>
        `;

        // Highlight the selected customer in the list
        document.querySelectorAll('.customer-list-item').forEach(item => {
            item.classList.remove('active');
            if (item.textContent === customerName) {
                item.classList.add('active');
            }
        });
    }


    // Function to get progress bar class based on percentage
    function getProgressBarClass(percentage) {
        if (percentage >= 80) return 'success';
        if (percentage >= 60) return 'warning-high';
        if (percentage >= 40) return 'warning-medium';
        if (percentage >= 20) return 'warning-low';
        if (percentage > 0) return 'danger';
        return 'no-activity'; // For 0 or N/A
    }

    // --- Event Listeners ---

    // Tab button event listeners
    allBranchSnapshotTabBtn.addEventListener('click', () => { showTab('allBranchSnapshotTabBtn'); generateAllBranchSnapshotReport(); });
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => { showTab('allStaffOverallPerformanceTabBtn'); generateAllStaffOverallPerformanceReport(); });
    nonParticipatingBranchesTabBtn.addEventListener('click', () => { showTab('nonParticipatingBranchesTabBtn'); generateNonParticipatingBranchesReport(); });
    detailedCustomerViewTabBtn.addEventListener('click', () => { showTab('detailedCustomerViewTabBtn'); });
    employeeManagementTabBtn.addEventListener('click', () => { showTab('employeeManagementTabBtn'); });


    // Report section dropdowns and buttons
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        const employeesInBranch = employeeMasterData.filter(e => e[HEADER_BRANCH_NAME] === selectedBranch).map(e => e[HEADER_EMPLOYEE_NAME]);
        populateDropdown(employeeSelect, employeesInBranch.sort());
        employeeSelect.value = ''; // Reset employee selection
        reportDisplay.innerHTML = '<p>Select an employee or view reports.</p>'; // Clear report
    });

    employeeSelect.addEventListener('change', () => {
        reportDisplay.innerHTML = '<p>View reports for selected employee.</p>'; // Clear report
    });

    // View Options buttons
    document.getElementById('viewBranchPerformanceReportBtn').addEventListener('click', () => {
        const selectedBranch = branchSelect.value;
        generateBranchPerformanceReport(selectedBranch);
    });

    document.getElementById('viewEmployeeSummaryBtn').addEventListener('click', () => {
        const selectedEmployee = employeeSelect.value;
        generateEmployeeSummary(selectedEmployee);
    });

    document.getElementById('viewAllEntriesBtn').addEventListener('click', () => {
        const selectedEmployee = employeeSelect.value;
        generateEmployeeAllEntriesReport(selectedEmployee);
    });

    // Add event listener for the "View Performance Report" button (placeholder for now)
    document.getElementById('viewPerformanceReportBtn').addEventListener('click', () => {
        const selectedEmployee = employeeSelect.value;
        if (selectedEmployee) {
            // This button can trigger the same summary or a more detailed performance specific to the employee
            // For now, let's make it show the employee summary. You can expand this later.
            generateEmployeeSummary(selectedEmployee);
        } else {
            reportDisplay.innerHTML = '<p>Please select an employee to view their performance report.</p>';
        }
    });


    // Detailed Customer View Filters and List Generation
    let filteredCustomerData = []; // Store data filtered by branch/employee

    customerViewBranchSelect.addEventListener('change', filterAndDisplayCustomers);
    customerViewEmployeeSelect.addEventListener('change', filterAndDisplayCustomers);

    function filterAndDisplayCustomers() {
        const selectedBranch = customerViewBranchSelect.value;
        const selectedEmployee = customerViewEmployeeSelect.value;

        filteredCustomerData = fullCanvassingData.filter(customer => {
            const matchesBranch = selectedBranch === '' || customer[HEADER_BRANCH_NAME] === selectedBranch;
            const matchesEmployee = selectedEmployee === '' || customer[HEADER_EMPLOYEE_NAME] === selectedEmployee;
            return matchesBranch && matchesEmployee;
        });

        populateCustomerList(filteredCustomerData);
        // Clear customer details when filters change
        displayCustomerDetails(null);
    }

    function populateCustomerList(customers) {
        customerCanvassedList.innerHTML = ''; // Clear existing list
        if (customers.length === 0) {
            customerCanvassedList.innerHTML = '<p>No customers found for the selected criteria.</p>';
            return;
        }

        const uniqueCustomers = new Set(); // To handle multiple entries for same customer
        customers.forEach(customer => {
            if (!uniqueCustomers.has(customer[HEADER_CUSTOMER_NAME])) {
                uniqueCustomers.add(customer[HEADER_CUSTOMER_NAME]);
                const listItem = document.createElement('li');
                listItem.className = 'customer-list-item';
                listItem.textContent = customer[HEADER_CUSTOMER_NAME];
                listItem.addEventListener('click', () => displayCustomerDetails(customer[HEADER_CUSTOMER_NAME]));
                customerCanvassedList.appendChild(listItem);
            }
        });
    }


    // Employee Management Form Submissions
    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: newEmployeeNameInput.value.trim(),
                [HEADER_EMPLOYEE_CODE]: newEmployeeCodeInput.value.trim(),
                [HEADER_BRANCH_NAME]: newBranchNameInput.value.trim(),
                [HEADER_DESIGNATION]: newDesignationInput.value.trim()
            };

            if (!employeeData[HEADER_EMPLOYEE_NAME] || !employeeData[HEADER_EMPLOYEE_CODE] || !employeeData[HEADER_BRANCH_NAME]) {
                displayEmployeeManagementMessage('Employee Name, Code, and Branch are required.', true);
                return;
            }

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);
            if (success) {
                addEmployeeForm.reset();
            }
        });
    }

    // Event Listener for Bulk Add Employees Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const detailsText = bulkEmployeeDetailsInput.value.trim();

            if (!branchName) {
                displayEmployeeManagementMessage('Branch Name for Bulk Entry is required.', true);
                return;
            }
            if (!detailsText) {
                displayEmployeeManagementMessage('Employee Details are required for bulk entry.', true);
                return;
            }

            const lines = detailsText.split('\n');
            const employeesToAdd = [];
            for (const line of lines) {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length >= 2 && parts[0] && parts[1]) { // Name, Code (Designation optional)
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
    showTab('allBranchSnapshotTabBtn');
});
