document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THrVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv";
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyr_LP1bpaoszUtWAlMeFxFIwzsUQF0USmdTRaTdqBHcn9HFo82nmBqQt4zokKCuHL5/exec"; // <--- Updated URL!

    const MONTHLY_WORKING_DAYS = 22; // Common approximation for a month's working days

    const TARGETS = {
        'Branch Manager': {
            'Visit': 10,
            'Call': 3 * MONTHLY_WORKING_DAYS,
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
    const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
    const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
    const viewPerformanceReportBtn = document.getElementById('viewPerformanceReportBtn');
    const reportDisplay = document.getElementById('reportDisplay');

    // Tab Elements
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');


    // *** Data Storage ***
    let allCanvassingData = []; // Stores all fetched data
    let filteredBranchData = []; // Stores data for the currently selected branch
    let selectedEmployeeEntries = []; // Stores entries for the currently selected employee

    // *** Utility Functions ***

    // Simple CSV to JSON converter
    function parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        if (lines.length <= 1) return []; // Only headers or no data

        const headers = lines[0].split(',').map(header => header.trim());
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = lines[i].split(',').map(value => value.trim());
            let entry = {};
            headers.forEach((header, index) => {
                if (header && values[index] !== undefined) {
                    entry[header] = values[index];
                }
            });
            data.push(entry);
        }
        return data;
    }

    // Displays a message in the report area
    function displayMessage(message) {
        reportDisplay.innerHTML = `<p>${message}</p>`;
    }

    // Populates a dropdown with unique values
    function populateDropdown(dropdownElement, dataArray, key) {
        dropdownElement.innerHTML = `<option value="">-- Select a ${key.replace(' Name', '')} --</option>`;
        const uniqueValues = new Set();
        dataArray.forEach(item => {
            if (item[key] && item[key].trim() !== '') {
                uniqueValues.add(item[key].trim());
            }
        });

        Array.from(uniqueValues).sort().forEach(value => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            dropdownElement.appendChild(option);
        });
    }

    // Filters data for the current month and year using string comparison
    function filterDataForCurrentMonth(data) {
        const now = new Date();
        const currentMonthString = String(now.getMonth() + 1).padStart(2, '0'); // '06' for June
        const currentYearString = String(now.getFullYear()); // '2025'

        console.log(`Filtering for Current Month (Target): ${currentMonthString}, Year: ${currentYearString}`); // Debugging

        return data.filter(entry => {
            if (entry['Date']) {
                const dateParts = entry['Date'].split('/'); // Assumes DD/MM/YYYY
                if (dateParts.length === 3) {
                    const entryMonthString = dateParts[1]; // '06'
                    const entryYearString = dateParts[2]; // '2025'

                    const isMatch = entryMonthString === currentMonthString && entryYearString === currentYearString;
                    return isMatch;
                }
            }
            return false;
        });
    }

    // Function to calculate key metrics for a given set of entries (used in multiple reports)
    function calculateMetrics(entries) {
        const metrics = {
            totalEntries: entries.length,
            visits: 0,
            calls: 0,
            references: 0,
            newCustomerLeads: 0,
            activityTypeCounts: {},
            customerTypeCounts: {},
            leadSourceCounts: {},
            productInterestedCounts: {},
            professionCounts: {}
        };

        entries.forEach(entry => {
            // Normalize activity and customer type strings
            const activity = entry['Activity Type'] ? entry['Activity Type'].trim().toLowerCase() : '';
            const customerType = entry['Type of Customer'] ? entry['Type of Customer'].trim().toLowerCase() : '';
            const leadSource = entry['Lead Source'];
            const product = entry['Prodcut Interested']; // Use original spelling from sheet
            const profession = entry['Profession'];

            // Counting based on your specific definitions (now case-insensitive and trimmed)
            // AND matching your exact predefined values (singular/plural)
            if (activity === 'visit') { // User specified 'Visit' (singular)
                metrics.visits++;
            }
            if (activity === 'calls') { // User specified 'Calls' (plural)
                metrics.calls++;
            }
            // User specified 'Referance' (typo). Check if customer is 'new' AND activity is 'referance'.
            if (customerType === 'new' && activity === 'referance') {
                metrics.references++;
            }
            if (customerType === 'new') {
                metrics.newCustomerLeads++;
            }

            // General counts for summaries (these were already working, but ensure consistency)
            if (entry['Activity Type']) metrics.activityTypeCounts[entry['Activity Type'].trim()] = (metrics.activityTypeCounts[entry['Activity Type'].trim()] || 0) + 1;
            if (entry['Type of Customer']) metrics.customerTypeCounts[entry['Type of Customer'].trim()] = (metrics.customerTypeCounts[entry['Type of Customer'].trim()] || 0) + 1;
            if (leadSource) metrics.leadSourceCounts[leadSource.trim()] = (metrics.leadSourceCounts[leadSource.trim()] || 0) + 1;
            if (product) metrics.productInterestedCounts[product.trim()] = (metrics.productInterestedCounts[product.trim()] || 0) + 1;
            if (profession) metrics.professionCounts[profession.trim()] = (metrics.professionCounts[profession.trim()] || 0) + 1;
        });

        return metrics;
    }


    // --- Rendering Functions ---

    // Renders all detailed entries for a given employee in a table
    function renderEmployeeDetailedEntries(entries) {
        if (entries.length === 0) {
            reportDisplay.innerHTML = `<p>No detailed entries found for this employee.</p>`;
            return;
        }

        const employeeName = entries[0]['Employee Name'];
        reportDisplay.innerHTML = `<h3>Detailed Entries for ${employeeName}</h3>`;

        const table = document.createElement('table');
        table.className = 'data-table';

        // Table Header
        const thead = table.createTHead();
        const headerRow = thead.insertCell();
        const columnsToDisplayInTable = [
            "Timestamp", "Date", "Employee Code", "Designation",
            "Activity Type", "Type of Customer", "Lead Source",
            "How Contacted", "Prospect Name", "Phone Numebr(Whatsapp)", "Address",
            "Profession", "DOB/WD", "Prodcut Interested", "Remarks",
            "Next Follow-up Date", "Relation With Staff"
        ];
        columnsToDisplayInTable.forEach(col => {
            const th = document.createElement('th');
            th.textContent = col;
            headerRow.appendChild(th);
        });

        // Table Body
        const tbody = table.createTBody();
        entries.forEach(entry => {
            const row = tbody.insertRow();
            columnsToDisplayInTable.forEach(col => {
                const cell = row.insertCell();
                cell.textContent = entry[col] || ''; // Display value or empty string
            });
        });
        reportDisplay.appendChild(table);
    }

    // Renders a summary for a given employee (activity counts)
    function renderEmployeeSummary(entries) {
        const employeeName = entries[0]['Employee Name'];
        reportDisplay.innerHTML = `<h3>Activity Summary for ${employeeName}</h3>`;
        if (entries.length === 0) {
            reportDisplay.innerHTML += `<p>No entries found for this employee.</p>`;
            return;
        }

        const metrics = calculateMetrics(entries);

        let summaryHtml = `<p><strong>Total Canvassing Entries:</strong> ${metrics.totalEntries}</p>`;

        // Wrap key activity counts and breakdowns for side-by-side layout
        summaryHtml += `<div class="summary-details-container">`; // Start new container

        summaryHtml += `<div><h4>Key Activity Counts:</h4><ul class="summary-list">`;
        summaryHtml += `<li>Visits: ${metrics.visits}</li>`;
        summaryHtml += `<li>Calls: ${metrics.calls}</li>`;
        summaryHtml += `<li>References: ${metrics.references}</li>`;
        summaryHtml += `<li>New Customer Leads: ${metrics.newCustomerLeads}</li>`;
        summaryHtml += `</ul></div>`;

        summaryHtml += `<div><h4>Activity Types Breakdown:</h4><ul class="summary-list">`;
        for (const type in metrics.activityTypeCounts) {
            summaryHtml += `<li>${type}: ${metrics.activityTypeCounts[type]}</li>`;
        }
        summaryHtml += `</ul></div>`;

        summaryHtml += `<div><h4>Customer Types Breakdown:</h4><ul class="summary-list">`;
        for (const type in metrics.customerTypeCounts) {
            summaryHtml += `<li>${type}: ${metrics.customerTypeCounts[type]}</li>`;
        }
        summaryHtml += `</ul></div>`;

        summaryHtml += `<div><h4>Products Interested Breakdown:</h4><ul class="summary-list">`;
        for (const product in metrics.productInterestedCounts) {
            summaryHtml += `<li>${product}: ${metrics.productInterestedCounts[product]}</li>`;
        }
        summaryHtml += `</ul></div>`;

        summaryHtml += `</div>`; // End new container

        reportDisplay.innerHTML += summaryHtml;
    }

    // Renders a summary for all staff in a branch with key metrics per employee
    function renderBranchSummary(branchData) {
        const branchName = branchData[0]['Branch Name'];
        reportDisplay.innerHTML = `<h3>All Staff Activity Summary for ${branchName} Branch</h3>`;

        if (branchData.length === 0) {
            reportDisplay.innerHTML += `<p>No entries found for this branch.</p>`;
            return;
        }

        // Group data by employee
        const employeesInBranch = {};
        branchData.forEach(entry => {
            const employee = entry['Employee Name'];
            if (!employeesInBranch[employee]) {
                employeesInBranch[employee] = [];
            }
            employeesInBranch[employee].push(entry);
        });

        let branchSummaryHtml = `<div class="branch-summary-grid">`; // Wrapper for side-by-side cards

        // Display summary for each employee
        for (const employeeName in employeesInBranch) {
            const empEntries = employeesInBranch[employeeName];
            const metrics = calculateMetrics(empEntries); // Calculate metrics for each employee

            let empHtml = `<div class="employee-summary-card">`;
            empHtml += `<h4>${employeeName} (Total Entries: ${metrics.totalEntries})</h4>`;
            empHtml += `<p><strong>Visits:</strong> ${metrics.visits}</p>`;
            empHtml += `<p><strong>Calls:</strong> ${metrics.calls}</p>`;
            empHtml += `<p><strong>References:</strong> ${metrics.references}</p>`;
            empHtml += `<p><strong>New Customer Leads:</strong> ${metrics.newCustomerLeads}</p>`;

            // Optional: add top 3 most common activity types
            const sortedActivities = Object.entries(metrics.activityTypeCounts).sort((a,b) => b[1] - a[1]);
            if(sortedActivities.length > 0) {
                empHtml += `<p><strong>Top Activities:</strong> ${sortedActivities.slice(0, 3).map(a => `${a[0]} (${a[1]})`).join(', ')}</p>`;
            }

            empHtml += `</div>`;
            branchSummaryHtml += empHtml;
        }
        branchSummaryHtml += `</div>`; // Close wrapper
        reportDisplay.innerHTML += branchSummaryHtml;
    }


    // Renders a performance report for all staff in a branch
    function renderBranchPerformanceReport(branchData) {
        const branchName = branchData[0]['Branch Name'];
        reportDisplay.innerHTML = `<h3>All Staff Performance Report for ${branchName} Branch (This Month)</h3>`;

        if (branchData.length === 0) {
            reportDisplay.innerHTML += `<p>No entries found for this branch to generate a performance report.</p>`;
            return;
        }

        // Group data by employee
        const employeesInBranch = {};
        branchData.forEach(entry => {
            const employee = entry['Employee Name'];
            if (!employeesInBranch[employee]) {
                employeesInBranch[employee] = [];
            }
            employeesInBranch[employee].push(entry);
        });

        let performanceReportHtml = `<div class="branch-performance-grid">`; // Wrapper for side-by-side performance cards

        // Display performance for each employee
        for (const employeeName in employeesInBranch) {
            const empEntries = employeesInBranch[employeeName];
            const employeeDesignation = empEntries[0]['Designation'] || 'Default';

            // Filter entries for the current month for target tracking
            const currentMonthEntries = filterDataForCurrentMonth(empEntries);
            const actuals = calculateMetrics(currentMonthEntries);

            // Get targets based on designation
            const employeeTargets = TARGETS[employeeDesignation] || TARGETS['Default'];

            performanceReportHtml += `<div class="employee-performance-card">`;
            performanceReportHtml += `<h4>${employeeName} (${employeeDesignation})</h4>`;

            let tableHtml = `<table class="performance-table">
                <thead>
                    <tr>
                        <th>Metric</th>
                        <th>Actual</th>
                        <th>Target</th>
                        <th>%</th>
                    </tr>
                </thead>
                <tbody>`;

            const metricsToDisplay = ['Visit', 'Call', 'Reference', 'New Customer Leads'];

            metricsToDisplay.forEach(metric => {
                let actualValue;
                switch(metric) {
                    case 'Visit': actualValue = actuals.visits; break;
                    case 'Call': actualValue = actuals.calls; break;
                    case 'Reference': actualValue = actuals.references; break;
                    case 'New Customer Leads': actualValue = actuals.newCustomerLeads; break;
                    default: actualValue = 0;
                }

                const target = employeeTargets[metric] || 0;
                const actual = actualValue || 0;

                const percentage = target > 0 ? ((actual / target) * 100).toFixed(0) : 'N/A';
                let progressBarWidth = 0;
                if (target > 0) {
                    progressBarWidth = (actual / target) * 100;
                }

                let progressBarClass = '';
                if (progressBarWidth >= 100) {
                    progressBarClass = 'overachieved';
                    progressBarWidth = 100;
                } else if (progressBarWidth >= 90) {
                    progressBarClass = 'success';
                } else if (progressBarWidth >= 50) {
                    progressBarClass = 'warning';
                } else {
                    progressBarClass = 'danger';
                }

                // Debugging log for targets and percentages in branch performance report
                console.log(`Branch Performance Debug: Employee=${employeeName}, Metric=${metric}, Actual=${actual}, Target=${target}, Percentage=${percentage}`);


                tableHtml += `
                    <tr>
                        <td class="performance-metric">${metric}</td>
                        <td>${actual}</td>
                        <td>${target}</td>
                        <td>
                            <div class="progress-bar-container-small">
                                <div class="progress-bar ${progressBarClass}" style="width: ${progressBarWidth}%;">
                                    ${percentage !== 'N/A' ? `${percentage}%` : 'N/A'}
                                </div>
                            </div>
                        </td>
                    </tr>`;
            });

            tableHtml += `</tbody></table>`;
            performanceReportHtml += tableHtml;
            performanceReportHtml += `</div>`; // Close employee-performance-card
        }
        performanceReportHtml += `</div>`; // Close branch-performance-grid
        reportDisplay.innerHTML += performanceReportHtml;
    }


    // Renders a performance report for a single employee against targets
    function renderPerformanceReport(entries, employeeName, designation) {
        reportDisplay.innerHTML = `<h3>Performance Report for ${employeeName} (${designation || 'N/A'})</h3>`;
        if (entries.length === 0) {
            reportDisplay.innerHTML += `<p>No entries for this employee to calculate performance.</p>`;
            return;
        }

        // Filter entries for the current month for target tracking
        const currentMonthEntries = filterDataForCurrentMonth(entries);
        console.log(`Current Month Entries for ${employeeName}:`, currentMonthEntries); // Debugging line

        // Calculate actuals for the current month using the new calculateMetrics function
        const actuals = calculateMetrics(currentMonthEntries);
        console.log("Calculated Actuals for current month:", actuals); // Debugging line

        // Get targets based on designation, default if not found
        const employeeTargets = TARGETS[designation] || TARGETS['Default'];

        // Display results in a table
        let tableHtml = `<table class="performance-table">
            <thead>
                <tr>
                    <th>Metric</th>
                    <th>Actual (This Month)</th>
                    <th>Target (Monthly)</th>
                    <th>Achievement (%)</th>
                    <th>Progress</th>
                </tr>
            </thead>
            <tbody>`;

        const metricsToDisplay = ['Visit', 'Call', 'Reference', 'New Customer Leads']; // Explicit order for display

        metricsToDisplay.forEach(metric => {
            let actualValue;
            // Map metric name to its actuals property
            switch(metric) {
                case 'Visit': actualValue = actuals.visits; break;
                case 'Call': actualValue = actuals.calls; break;
                case 'Reference': actualValue = actuals.references; break;
                case 'New Customer Leads': actualValue = actuals.newCustomerLeads; break;
                default: actualValue = 0;
            }

            const target = employeeTargets[metric] || 0;
            const actual = actualValue || 0; // Ensure it's a number

            const percentage = target > 0 ? ((actual / target) * 100).toFixed(0) : 'N/A';
            let progressBarWidth = 0;
            if (target > 0) {
                progressBarWidth = (actual / target) * 100;
            }

            let progressBarClass = '';
            if (progressBarWidth >= 100) {
                progressBarClass = 'overachieved'; // New class for over 100%
                progressBarWidth = 100; // Cap visual bar at 100%
            } else if (progressBarWidth >= 90) {
                progressBarClass = 'success';
            } else if (progressBarWidth >= 50) {
                progressBarClass = 'warning';
            } else {
                progressBarClass = 'danger';
            }

            // Debugging log for targets and percentages in single employee performance report
            console.log(`Single Employee Performance Debug: Employee=${employeeName}, Metric=${metric}, Actual=${actual}, Target=${target}, Percentage=${percentage}`);

            tableHtml += `
                <tr>
                    <td class="performance-metric">${metric}</td>
                    <td>${actual}</td>
                    <td>${target}</td>
                    <td>${percentage}%</td>
                    <td>
                        <div class="progress-bar-container">
                            <div class="progress-bar ${progressBarClass}" style="width: ${progressBarWidth}%;">
                                ${percentage !== 'N/A' ? `${percentage}%` : 'N/A'}
                            </div>
                        </div>
                    </td>
                </tr>`;
        });
        tableHtml += `</tbody></table>`;
        reportDisplay.innerHTML += tableHtml;
    }

    // Renders an "All Branch Snapshot" report
    function renderAllBranchSnapshot() {
        if (allCanvassingData.length === 0) {
            displayMessage("No data available to generate an All Branch Snapshot. Please ensure the Google Sheet is published and accessible.");
            return;
        }

        reportDisplay.innerHTML = `<h3>All Branch Snapshot (This Month)</h3>`;

        const currentMonthData = filterDataForCurrentMonth(allCanvassingData);
        if (currentMonthData.length === 0) {
            reportDisplay.innerHTML += `<p>No canvassing entries for the current month across all branches.</p>`;
            return;
        }

        // Group data by branch
        const branches = {};
        currentMonthData.forEach(entry => {
            const branchName = entry['Branch Name'];
            if (branchName) {
                if (!branches[branchName]) {
                    branches[branchName] = [];
                }
                branches[branchName].push(entry);
            }
        });

        let snapshotHtml = `<div class="branch-snapshot-grid">`; // Grid for branch cards

        for (const branchName in branches) {
            const branchEntries = branches[branchName];
            const metrics = calculateMetrics(branchEntries);

            snapshotHtml += `<div class="branch-snapshot-card">`;
            snapshotHtml += `<h4>${branchName}</h4>`;
            snapshotHtml += `<p><strong>Total Entries:</strong> ${metrics.totalEntries}</p>`;
            snapshotHtml += `<p><strong>Visits:</strong> ${metrics.visits}</p>`;
            snapshotHtml += `<p><strong>Calls:</strong> ${metrics.calls}</p>`;
            snapshotHtml += `<p><strong>References:</strong> ${metrics.references}</p>`;
            snapshotHtml += `<p><strong>New Customer Leads:</strong> ${metrics.newCustomerLeads}</p>`;

            // Add breakdown of activity types for the branch
            snapshotHtml += `<h5>Activity Breakdown:</h5><ul class="summary-list">`;
            for (const type in metrics.activityTypeCounts) {
                snapshotHtml += `<li>${type}: ${metrics.activityTypeCounts[type]}</li>`;
            }
            snapshotHtml += `</ul>`;

            snapshotHtml += `</div>`; // Close branch-snapshot-card
        }
        snapshotHtml += `</div>`; // Close branch-snapshot-grid
        reportDisplay.innerHTML += snapshotHtml;
    }


    // Renders an "Overall Staff Performance" report across all branches
    function renderOverallStaffPerformanceReport() {
        if (allCanvassingData.length === 0) {
            displayMessage("No data available to generate an Overall Staff Performance Report. Please ensure the Google Sheet is published and accessible.");
            return;
        }

        reportDisplay.innerHTML = `<h3>All Staff Overall Performance (This Month)</h3>`;

        const currentMonthData = filterDataForCurrentMonth(allCanvassingData);
        if (currentMonthData.length === 0) {
            reportDisplay.innerHTML += `<p>No canvassing entries for the current month across all staff.</p>`;
            return;
        }

        // Group data by employee across all branches
        const allEmployees = {};
        currentMonthData.forEach(entry => {
            const employeeName = entry['Employee Name'];
            if (employeeName) {
                if (!allEmployees[employeeName]) {
                    allEmployees[employeeName] = [];
                }
                allEmployees[employeeName].push(entry);
            }
        });

        let overallPerformanceHtml = `<div class="overall-performance-grid">`;

        for (const employeeName in allEmployees) {
            const empEntries = allEmployees[employeeName];
            const employeeDesignation = empEntries[0]['Designation'] || 'Default'; // Assuming designation is consistent per employee

            const actuals = calculateMetrics(empEntries);
            const employeeTargets = TARGETS[employeeDesignation] || TARGETS['Default'];

            overallPerformanceHtml += `<div class="employee-performance-card">`;
            overallPerformanceHtml += `<h4>${employeeName} (${employeeDesignation})</h4>`;

            let tableHtml = `<table class="performance-table">
                <thead>
                    <tr>
                        <th>Metric</th>
                        <th>Actual</th>
                        <th>Target</th>
                        <th>%</th>
                    </tr>
                </thead>
                <tbody>`;

            const metricsToDisplay = ['Visit', 'Call', 'Reference', 'New Customer Leads'];

            metricsToDisplay.forEach(metric => {
                let actualValue;
                switch(metric) {
                    case 'Visit': actualValue = actuals.visits; break;
                    case 'Call': actualValue = actuals.calls; break;
                    case 'Reference': actualValue = actuals.references; break;
                    case 'New Customer Leads': actualValue = actuals.newCustomerLeads; break;
                    default: actualValue = 0;
                }

                const target = employeeTargets[metric] || 0;
                const actual = actualValue || 0;

                const percentage = target > 0 ? ((actual / target) * 100).toFixed(0) : 'N/A';
                let progressBarWidth = 0;
                if (target > 0) {
                    progressBarWidth = (actual / target) * 100;
                }

                let progressBarClass = '';
                if (progressBarWidth >= 100) {
                    progressBarClass = 'overachieved';
                    progressBarWidth = 100;
                } else if (progressBarWidth >= 90) {
                    progressBarClass = 'success';
                } else if (progressBarWidth >= 50) {
                    progressBarClass = 'warning';
                } else {
                    progressBarClass = 'danger';
                }

                tableHtml += `
                    <tr>
                        <td class="performance-metric">${metric}</td>
                        <td>${actual}</td>
                        <td>${target}</td>
                        <td>
                            <div class="progress-bar-container-small">
                                <div class="progress-bar ${progressBarClass}" style="width: ${progressBarWidth}%;">
                                    ${percentage !== 'N/A' ? `${percentage}%` : 'N/A'}
                                </div>
                            </div>
                        </td>
                    </tr>`;
            });

            tableHtml += `</tbody></table>`;
            overallPerformanceHtml += tableHtml;
            overallPerformanceHtml += `</div>`; // Close employee-performance-card
        }
        overallPerformanceHtml += `</div>`; // Close overall-performance-grid
        reportDisplay.innerHTML += overallPerformanceHtml;
    }


    // Function to handle tab switching
    function showTab(activeTabId) {
        // Remove 'active' from all tab buttons
        const tabButtons = document.querySelectorAll('.tab-button');
        tabButtons.forEach(button => button.classList.remove('active'));

        // Add 'active' to the clicked tab button
        const activeTabButton = document.getElementById(activeTabId);
        if (activeTabButton) {
            activeTabButton.classList.add('active');
        }

        // Render the corresponding report
        if (activeTabId === 'allBranchSnapshotTabBtn') {
            renderAllBranchSnapshot();
        } else if (activeTabId === 'allStaffOverallPerformanceTabBtn') {
            renderOverallStaffPerformanceReport();
        }
    }


    // *** Main Data Fetching and Initialization ***
    async function fetchData() {
        displayMessage("Loading data...");
        try {
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const csvText = await response.text();
            allCanvassingData = parseCSV(csvText);
            console.log("Fetched Canvassing Data:", allCanvassingData); // Debugging

            if (allCanvassingData.length > 0) {
                populateDropdown(branchSelect, allCanvassingData, 'Branch Name');
                // Initially show the "All Branch Snapshot" report
                showTab('allBranchSnapshotTabBtn');
            } else {
                displayMessage("No data found in the Google Sheet. Please check the sheet and URL.");
            }
        } catch (error) {
            console.error("Error fetching or parsing data:", error);
            displayMessage(`Failed to load data: ${error.message}. Please ensure the Google Sheet is published to web as CSV.`);
        }
    }

    // *** Event Listeners ***

    // Branch selection change
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            filteredBranchData = allCanvassingData.filter(entry => entry['Branch Name'] === selectedBranch);
            populateDropdown(employeeSelect, filteredBranchData, 'Employee Name');
            employeeFilterPanel.style.display = 'block';
            viewOptions.style.display = 'flex'; // Show view options
            displayMessage("Select an employee or a report option.");
            employeeSelect.value = ""; // Reset employee selection
            selectedEmployeeEntries = []; // Clear selected employee entries
        } else {
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none';
            displayMessage("Please select a branch from the dropdown above to view reports.");
            // When no branch is selected, show the default tab report
            showTab('allBranchSnapshotTabBtn');
        }
    });

    // Employee selection change
    employeeSelect.addEventListener('change', () => {
        const selectedEmployee = employeeSelect.value;
        if (selectedEmployee) {
            selectedEmployeeEntries = filteredBranchData.filter(entry => entry['Employee Name'] === selectedEmployee);
            // Default to showing employee summary when selected
            if (selectedEmployeeEntries.length > 0) {
                renderEmployeeSummary(selectedEmployeeEntries);
            } else {
                displayMessage("No data found for the selected employee.");
            }
        } else {
            selectedEmployeeEntries = []; // Clear selected employee entries
            // If employee is unselected, show branch summary
            if (filteredBranchData.length > 0) {
                renderBranchSummary(filteredBranchData);
            } else {
                displayMessage("Select an employee or a report option.");
            }
        }
    });

    // View Report Buttons Event Listeners
    viewBranchSummaryBtn.addEventListener('click', () => {
        if (filteredBranchData.length > 0) {
            renderBranchSummary(filteredBranchData);
        } else {
            displayMessage("No branch selected or no data for this branch to show summary.");
        }
    });

    viewBranchPerformanceReportBtn.addEventListener('click', () => {
        if (filteredBranchData.length > 0) {
            renderBranchPerformanceReport(filteredBranchData);
        } else {
            displayMessage("No branch selected or no data for this branch for performance report.");
        }
    });

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

    // Event listeners for tab buttons
    if (allBranchSnapshotTabBtn) {
        allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    }
    if (allStaffOverallPerformanceTabBtn) {
        allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    }


    // Initial data fetch when the page loads
    fetchData();

    // --- New Code for Employee Management ---

    // DOM elements for the employee management form
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const employeeManagementMessage = document.getElementById('employeeManagementMessage');

    // Function to display messages in the employee management section
    function displayEmployeeManagementMessage(message, isError = false) {
        if (employeeManagementMessage) {
            employeeManagementMessage.textContent = message;
            employeeManagementMessage.style.color = isError ? 'red' : 'green';
            employeeManagementMessage.style.display = 'block';
            setTimeout(() => {
                employeeManagementMessage.style.display = 'none';
                employeeManagementMessage.textContent = '';
            }, 5000); // Hide after 5 seconds
        }
    }

    // Handle Add Employee form submission
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            const employeeName = document.getElementById('newEmployeeName').value;
            const employeeCode = document.getElementById('newEmployeeCode').value;
            const branchName = document.getElementById('newBranchName').value;
            const designation = document.getElementById('newDesignation').value || 'Default'; // Default designation

            if (!employeeName || !employeeCode || !branchName) {
                displayEmployeeManagementMessage('Please fill in all required fields.', true);
                return;
            }

            const dataToSend = {
                type: 'add',
                data: [{
                    'Employee Name': employeeName,
                    'Employee Code': employeeCode,
                    'Branch Name': branchName,
                    'Designation': designation
                }]
            };

            try {
                const response = await fetch(WEB_APP_URL, {
                    method: 'POST',
                    mode: 'no-cors', // Important for Apps Script Web App
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify(dataToSend)
                });

                // Since 'no-cors' mode, response.ok will always be true and status 0
                // We rely on the Apps Script to handle errors internally and assume success if no network error
                displayEmployeeManagementMessage('Employee added (or request sent). Check your Google Sheet for updates.', false);

                // Clear the form
                addEmployeeForm.reset();
                // Optionally refetch data to update dropdowns if relevant for other sections
                // fetchData();

            } catch (error) {
                console.error('Error adding employee:', error);
                displayEmployeeManagementMessage(`Failed to add employee: ${error.message}.`, true);
            }
        });
    }

    // Handle Delete Employee form submission
    if (deleteEmployeeForm) {
        deleteEmployeeForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            const employeeCodeToDelete = document.getElementById('deleteEmployeeCode').value;

            if (!employeeCodeToDelete) {
                displayEmployeeManagementMessage('Please enter an Employee Code to delete.', true);
                return;
            }

            const dataToSend = {
                type: 'delete',
                employeeCode: employeeCodeToDelete
            };

            try {
                const response = await fetch(WEB_APP_URL, {
                    method: 'POST',
                    mode: 'no-cors', // Important for Apps Script Web App
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify(dataToSend)
                });

                // Similar to add, due to 'no-cors', assume success or check sheet manually
                displayEmployeeManagementMessage('Delete request sent. Check your Google Sheet for updates.', false);

                // Clear the form
                deleteEmployeeForm.reset();
                // Optionally refetch data to update dropdowns if relevant for other sections
                // fetchData();

            } catch (error) {
                console.error('Error deleting employee:', error);
                displayEmployeeManagementMessage(`Failed to delete employee: ${error.message}.`, true);
            }
        });
    }

    // Toggle Employee Management Section Visibility
    const toggleEmployeeManagementBtn = document.getElementById('toggleEmployeeManagement');
    const employeeManagementSection = document.getElementById('employeeManagementSection');

    if (toggleEmployeeManagementBtn && employeeManagementSection) {
        toggleEmployeeManagementBtn.addEventListener('click', () => {
            if (employeeManagementSection.style.display === 'none' || employeeManagementSection.style.display === '') {
                employeeManagementSection.style.display = 'block';
                toggleEmployeeManagementBtn.textContent = 'Hide Employee Management';
            } else {
                employeeManagementSection.style.display = 'none';
                toggleEmployeeManagementBtn.textContent = 'Show Employee Management';
            }
        });
    }
});
