document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv";
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec";

    // Column Headers (Case-sensitive, must match your Google Sheet headers)
    const HEADER_TIMESTAMP = "Timestamp";
    const HEADER_BRANCH = "Branch";
    const HEADER_EMPLOYEE_NAME = "Employee Name";
    const HEADER_EMPLOYEE_CODE = "Employee Code";
    const HEADER_DESIGNATION = "Designation";
    const HEADER_CUSTOMER_NAME = "Customer Name";
    const HEADER_CUSTOMER_CONTACT = "Customer Contact";
    const HEADER_ACTIVITY_TYPE = "Activity Type";
    const HEADER_REMARKS = "Remarks";
    const HEADER_PRODUCT_INTEREST = "Product Interest";
    const HEADER_LEAD_SOURCE = "Lead Source";
    const HEADER_CUSTOMER_PROFESSION = "Customer Profession";
    const HEADER_CALL_OUTCOME = "Call Outcome";
    const HEADER_NEXT_ACTION = "Next Action Date/Time";
    const HEADER_STATUS = "Status"; // For Lead Status (e.g., Hot, Warm, Cold)
    const HEADER_VALUE_ACHIEVED = "Value Achieved"; // New column for monetary value

    // *** DOM Elements ***
    const statusMessage = document.getElementById('statusMessage');
    const reportDisplay = document.getElementById('reportDisplay');
    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const viewAllBranchSnapshotBtn = document.getElementById('viewAllBranchSnapshotBtn');
    const viewOverallStaffPerformanceBtn = document.getElementById('viewOverallStaffPerformanceBtn');
    const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
    const viewAllCanvassingEntriesBtn = document.getElementById('viewAllCanvassingEntriesBtn');

    // Tab Buttons
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    // Sections
    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection');
    const employeeManagementSection = document.getElementById('employeeManagementSection');

    // Detailed Customer View elements
    const customerBranchSelect = document.getElementById('customerBranchSelect');
    const customerEmployeeSelect = document.getElementById('customerEmployeeSelect');
    const customerList = document.getElementById('customerList');
    const customerDetailsCard = document.getElementById('customerDetailsCard');

    // Employee Management elements
    const employeeManagementMessage = document.getElementById('employeeManagementMessage');
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const employeeNameInput = document.getElementById('employeeName');
    const employeeCodeInput = document.getElementById('employeeCode');
    const branchNameInput = document.getElementById('branchName');
    const designationInput = document.getElementById('designation');
    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsInput = document.getElementById('bulkEmployeeDetails');
    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');


    // *** Global Data Variables ***
    let allCanvassingData = [];
    let processedBranches = new Set();
    let processedEmployees = new Set();
    let branchEmployeeMap = {}; // Maps branch to a Set of employees in that branch
    let allUniqueBranches = []; // For dropdowns
    let allUniqueEmployees = []; // For dropdowns
    let employeeDataMap = {}; // Stores full employee data by code

    // *** Chart Instances ***
    // Store chart instances globally to destroy them before redrawing
    let allBranchSnapshotChart = null;
    let overallStaffPerformanceChart = null;
    let employeeActivityChart = null;
    let employeeLeadSourceChart = null;
    let employeeProductInterestChart = null;

    // *** Helper Functions ***

    function displayMessage(message, isError = false) {
        statusMessage.innerHTML = `<div class="message ${isError ? 'error' : 'success'}">${message}</div>`;
        statusMessage.style.display = 'block';
        setTimeout(() => {
            statusMessage.style.display = 'none';
            statusMessage.innerHTML = '';
        }, 5000); // Message disappears after 5 seconds
    }

    function displayEmployeeManagementMessage(message, isError = false) {
        employeeManagementMessage.innerHTML = `<div class="message ${isError ? 'error' : 'success'}">${message}</div>`;
        employeeManagementMessage.style.display = 'block';
        setTimeout(() => {
            employeeManagementMessage.style.display = 'none';
            employeeManagementMessage.innerHTML = '';
        }, 5000); // Message disappears after 5 seconds
    }


    async function sendDataToGoogleAppsScript(action, data) {
        displayMessage('Processing...', false); // Show processing message
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({
                    action: action,
                    data: JSON.stringify(data)
                })
            });

            const result = await response.json();
            if (result.status === 'success') {
                displayMessage(result.message);
                // Re-fetch and process data after successful operation
                await processData();
                return true;
            } else {
                displayMessage(`Error: ${result.message}`, true);
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Apps Script:', error);
            displayMessage('Network error or server unreachable.', true);
            return false;
        }
    }

    async function fetchCanvassingData() {
        try {
            const response = await fetch(DATA_URL);
            const csvText = await response.text();
            return parseCSV(csvText);
        } catch (error) {
            console.error('Error fetching canvassing data:', error);
            displayMessage('Failed to load canvassing data. Please check the data source URL.', true);
            return null;
        }
    }

    function parseCSV(csv) {
        const lines = csv.split('\n').filter(line => line.trim() !== ''); // Filter out empty lines
        if (lines.length === 0) return [];

        const headers = lines[0].split(',').map(header => header.trim());
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = lines[i].split(',').map(value => value.trim());
            if (values.length === headers.length) {
                let row = {};
                headers.forEach((header, index) => {
                    row[header] = values[index];
                });
                data.push(row);
            }
        }
        return data;
    }

    // Function to populate dropdowns
    function populateDropdown(selectElement, items, addAllOption = false) {
        selectElement.innerHTML = ''; // Clear existing options
        if (addAllOption) {
            const allOption = document.createElement('option');
            allOption.value = '';
            allOption.textContent = `-- All ${selectElement.id.includes('branch') ? 'Branches' : 'Employees'} --`;
            selectElement.appendChild(allOption);
        } else {
             const defaultOption = document.createElement('option');
             defaultOption.value = '';
             defaultOption.textContent = `-- Select a ${selectElement.id.includes('branch') ? 'Branch' : 'Employee'} --`;
             selectElement.appendChild(defaultOption);
        }

        items.forEach(item => {
            const option = document.createElement('option');
            option.value = item;
            option.textContent = item;
            selectElement.appendChild(option);
        });
    }

    async function processData() {
        displayMessage('Loading data...');
        allCanvassingData = await fetchCanvassingData();
        if (!allCanvassingData) {
            // Data fetching failed, keep existing data or clear
            allCanvassingData = [];
            allUniqueBranches = [];
            allUniqueEmployees = [];
            branchEmployeeMap = {};
            employeeDataMap = {};
            populateDropdown(branchSelect, []);
            populateDropdown(employeeSelect, []);
            populateDropdown(customerBranchSelect, []);
            populateDropdown(customerEmployeeSelect, []);
            displayMessage('Failed to load data. Please check your data source.', true);
            return;
        }

        const uniqueBranches = new Set();
        const uniqueEmployees = new Set();
        const tempBranchEmployeeMap = {}; // Use temp map to build

        allCanvassingData.forEach(entry => {
            const branch = entry[HEADER_BRANCH];
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const designation = entry[HEADER_DESIGNATION];

            if (branch) uniqueBranches.add(branch);
            if (employeeName && employeeCode) {
                uniqueEmployees.add(employeeName);
                if (!tempBranchEmployeeMap[branch]) {
                    tempBranchEmployeeMap[branch] = new Set();
                }
                tempBranchEmployeeMap[branch].add(employeeName);
                employeeDataMap[employeeName] = { // Store full employee data
                    [HEADER_EMPLOYEE_CODE]: employeeCode,
                    [HEADER_DESIGNATION]: designation,
                    [HEADER_BRANCH]: branch
                };
            }
        });

        allUniqueBranches = Array.from(uniqueBranches).sort();
        allUniqueEmployees = Array.from(uniqueEmployees).sort();
        branchEmployeeMap = tempBranchEmployeeMap; // Assign the fully built map

        populateDropdown(branchSelect, allUniqueBranches, true);
        populateDropdown(employeeSelect, allUniqueEmployees, true);
        populateDropdown(customerBranchSelect, allUniqueBranches);
        populateDropdown(customerEmployeeSelect, allUniqueEmployees);

        displayMessage('Data loaded successfully!');
    }

    function calculateBranchMetrics(branchName) {
        const branchData = branchName ? allCanvassingData.filter(entry => entry[HEADER_BRANCH] === branchName) : allCanvassingData;

        const metrics = {
            totalVisits: 0,
            totalCalls: 0,
            totalProductsSold: 0,
            totalValueAchieved: 0,
            employeePerformance: {} // For individual employee metrics within the branch
        };

        const activityTypes = {
            'Visit': 'totalVisits',
            'Call': 'totalCalls',
            'Product Sold': 'totalProductsSold'
        };

        branchData.forEach(entry => {
            const activityType = entry[HEADER_ACTIVITY_TYPE];
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const valueAchieved = parseFloat(entry[HEADER_VALUE_ACHIEVED] || 0);

            if (activityTypes[activityType]) {
                metrics[activityTypes[activityType]]++;
            }
            metrics.totalValueAchieved += valueAchieved;

            // Employee specific metrics
            if (employeeName) {
                if (!metrics.employeePerformance[employeeName]) {
                    metrics.employeePerformance[employeeName] = {
                        visits: 0,
                        calls: 0,
                        productsSold: 0,
                        valueAchieved: 0,
                        employeeCode: entry[HEADER_EMPLOYEE_CODE],
                        designation: entry[HEADER_DESIGNATION],
                        branch: entry[HEADER_BRANCH]
                    };
                }
                if (activityTypes[activityType]) {
                    metrics.employeePerformance[employeeName][activityTypes[activityType].replace('total', '').toLowerCase()]++;
                }
                metrics.employeePerformance[employeeName].valueAchieved += valueAchieved;
            }
        });

        return metrics;
    }

    function calculateEmployeeMetrics(employeeName) {
        const employeeData = allCanvassingData.filter(entry => entry[HEADER_EMPLOYEE_NAME] === employeeName);

        const metrics = {
            totalVisits: 0,
            totalCalls: 0,
            totalProductsSold: 0,
            totalValueAchieved: 0,
            activityTypeBreakdown: {},
            leadSources: {},
            productInterest: {}
        };

        const activityTypes = {
            'Visit': 'totalVisits',
            'Call': 'totalCalls',
            'Product Sold': 'totalProductsSold'
        };

        employeeData.forEach(entry => {
            const activityType = entry[HEADER_ACTIVITY_TYPE];
            const leadSource = entry[HEADER_LEAD_SOURCE];
            const productInterest = entry[HEADER_PRODUCT_INTEREST];
            const valueAchieved = parseFloat(entry[HEADER_VALUE_ACHIEVED] || 0);

            if (activityTypes[activityType]) {
                metrics[activityTypes[activityType]]++;
            }
            metrics.totalValueAchieved += valueAchieved;

            metrics.activityTypeBreakdown[activityType] = (metrics.activityTypeBreakdown[activityType] || 0) + 1;
            if (leadSource) metrics.leadSources[leadSource] = (metrics.leadSources[leadSource] || 0) + 1;
            if (productInterest) metrics.productInterest[productInterest] = (metrics.productInterest[productInterest] || 0) + 1;
        });

        return metrics;
    }

    // Fixed targets (for demonstration, ideally from a config sheet)
    const TARGETS = {
        'Visit': 200,
        'Call': 150,
        'Product Sold': 50,
        'Value Achieved': 100000
    };

    function calculateAchievement(actual, target) {
        if (target === 0) return 'N/A';
        return ((actual / target) * 100).toFixed(2) + '%';
    }

    function getProgressBarColor(percentage) {
        if (percentage === 'N/A') return 'no-activity';
        const p = parseFloat(percentage);
        if (p >= 100) return 'success';
        if (p >= 75) return 'warning-high';
        if (p >= 50) return 'warning-medium';
        if (p >= 25) return 'warning-low';
        return 'danger';
    }

    function createProgressBarHTML(percentage) {
        const p = parseFloat(percentage);
        const width = isNaN(p) ? 0 : Math.min(p, 100); // Cap at 100% for display
        const colorClass = getProgressBarColor(percentage);
        const displayValue = percentage === 'N/A' ? 'N/A' : `${width.toFixed(0)}%`;

        return `
            <div class="progress-bar-container-small">
                <div class="progress-bar ${colorClass}" style="width: ${width}%;">
                    ${displayValue}
                </div>
            </div>
        `;
    }

    // --- NEW: Chart Functions ---

    function destroyChart(chartInstance) {
        if (chartInstance) {
            chartInstance.destroy();
            chartInstance = null; // Clear the reference
        }
    }

    function renderAllBranchSnapshotChart(branchesData) {
        destroyChart(allBranchSnapshotChart); // Destroy previous chart instance

        const canvas = document.createElement('canvas');
        canvas.id = 'allBranchSnapshotChart';
        canvas.style.maxHeight = '400px'; // Limit height for better display
        reportDisplay.appendChild(canvas);

        const labels = branchesData.map(b => b.branch);
        const visits = branchesData.map(b => b.totalVisits);
        const calls = branchesData.map(b => b.totalCalls);
        const productsSold = branchesData.map(b => b.totalProductsSold);
        const valueAchieved = branchesData.map(b => b.totalValueAchieved);

        allBranchSnapshotChart = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'Total Visits',
                        data: visits,
                        backgroundColor: 'rgba(75, 192, 192, 0.6)',
                        borderColor: 'rgba(75, 192, 192, 1)',
                        borderWidth: 1
                    },
                    {
                        label: 'Total Calls',
                        data: calls,
                        backgroundColor: 'rgba(153, 102, 255, 0.6)',
                        borderColor: 'rgba(153, 102, 255, 1)',
                        borderWidth: 1
                    },
                    {
                        label: 'Products Sold',
                        data: productsSold,
                        backgroundColor: 'rgba(255, 159, 64, 0.6)',
                        borderColor: 'rgba(255, 159, 64, 1)',
                        borderWidth: 1
                    },
                     {
                        label: 'Value Achieved',
                        data: valueAchieved,
                        backgroundColor: 'rgba(54, 162, 235, 0.6)',
                        borderColor: 'rgba(54, 162, 235, 1)',
                        borderWidth: 1
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: 'All Branch Performance Snapshot',
                        font: { size: 16 }
                    }
                },
                scales: {
                    x: {
                        title: {
                            display: true,
                            text: 'Branch'
                        }
                    },
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Count / Value'
                        }
                    }
                }
            }
        });
    }

    function renderOverallStaffPerformanceChart(employeesData) {
        destroyChart(overallStaffPerformanceChart); // Destroy previous chart instance

        const canvas = document.createElement('canvas');
        canvas.id = 'overallStaffPerformanceChart';
        canvas.style.maxHeight = '400px'; // Limit height for better display
        reportDisplay.appendChild(canvas);

        const labels = employeesData.map(e => e.employeeName);
        const visits = employeesData.map(e => e.visits);
        const calls = employeesData.map(e => e.calls);
        const productsSold = employeesData.map(e => e.productsSold);
        const valueAchieved = employeesData.map(e => e.valueAchieved);
        const visitAchievement = employeesData.map(e => parseFloat(e.visitAchievement.replace('%', '')) || 0);
        const callAchievement = employeesData.map(e => parseFloat(e.callAchievement.replace('%', '')) || 0);
        const productAchievement = employeesData.map(e => parseFloat(e.productAchievement.replace('%', '')) || 0);
        const valueAchievement = employeesData.map(e => parseFloat(e.valueAchievement.replace('%', '')) || 0);


        overallStaffPerformanceChart = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'Visits',
                        data: visits,
                        backgroundColor: 'rgba(75, 192, 192, 0.6)',
                        stack: 'performance',
                        borderColor: 'rgba(75, 192, 192, 1)',
                        borderWidth: 1
                    },
                    {
                        label: 'Calls',
                        data: calls,
                        backgroundColor: 'rgba(153, 102, 255, 0.6)',
                        stack: 'performance',
                        borderColor: 'rgba(153, 102, 255, 1)',
                        borderWidth: 1
                    },
                    {
                        label: 'Products Sold',
                        data: productsSold,
                        backgroundColor: 'rgba(255, 159, 64, 0.6)',
                        stack: 'performance',
                        borderColor: 'rgba(255, 159, 64, 1)',
                        borderWidth: 1
                    },
                    {
                        label: 'Value Achieved',
                        data: valueAchieved,
                        backgroundColor: 'rgba(54, 162, 235, 0.6)',
                        stack: 'performance',
                        borderColor: 'rgba(54, 162, 235, 1)',
                        borderWidth: 1
                    },
                    // Achievement percentages can be tricky to overlay on a bar chart with counts/values.
                    // A separate line chart or a different chart type might be better,
                    // but for simplicity, we'll keep it to counts/values in this stacked bar.
                    // If showing percentages, consider a combined chart with a different Y-axis.
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: 'Overall Staff Performance (Counts)',
                        font: { size: 16 }
                    }
                },
                scales: {
                    x: {
                        title: {
                            display: true,
                            text: 'Employee'
                        }
                    },
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Count / Value'
                        },
                        stacked: false // Set to true if you want bars stacked, false for side-by-side
                    }
                }
            }
        });

        // Add a second chart for achievement percentages if needed, perhaps a radar chart or another bar chart
        const achievementCanvas = document.createElement('canvas');
        achievementCanvas.id = 'overallStaffAchievementChart';
        achievementCanvas.style.maxHeight = '400px';
        reportDisplay.appendChild(achievementCanvas);

        overallStaffPerformanceChart = new Chart(achievementCanvas, { // Re-using the variable name is fine for a separate chart
            type: 'radar', // Radar chart is good for multiple percentage metrics
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'Visits Achievement (%)',
                        data: visitAchievement,
                        backgroundColor: 'rgba(75, 192, 192, 0.2)',
                        borderColor: 'rgba(75, 192, 192, 1)',
                        pointBackgroundColor: 'rgba(75, 192, 192, 1)',
                        pointBorderColor: '#fff',
                        pointHoverBackgroundColor: '#fff',
                        pointHoverBorderColor: 'rgba(75, 192, 192, 1)'
                    },
                    {
                        label: 'Calls Achievement (%)',
                        data: callAchievement,
                        backgroundColor: 'rgba(153, 102, 255, 0.2)',
                        borderColor: 'rgba(153, 102, 255, 1)',
                        pointBackgroundColor: 'rgba(153, 102, 255, 1)',
                        pointBorderColor: '#fff',
                        pointHoverBackgroundColor: '#fff',
                        pointHoverBorderColor: 'rgba(153, 102, 255, 1)'
                    },
                    {
                        label: 'Products Sold Achievement (%)',
                        data: productAchievement,
                        backgroundColor: 'rgba(255, 159, 64, 0.2)',
                        borderColor: 'rgba(255, 159, 64, 1)',
                        pointBackgroundColor: 'rgba(255, 159, 64, 1)',
                        pointBorderColor: '#fff',
                        pointHoverBackgroundColor: '#fff',
                        pointHoverBorderColor: 'rgba(255, 159, 64, 1)'
                    },
                     {
                        label: 'Value Achieved Achievement (%)',
                        data: valueAchievement,
                        backgroundColor: 'rgba(54, 162, 235, 0.2)',
                        borderColor: 'rgba(54, 162, 235, 1)',
                        pointBackgroundColor: 'rgba(54, 162, 235, 1)',
                        pointBorderColor: '#fff',
                        pointHoverBackgroundColor: '#fff',
                        pointHoverBorderColor: 'rgba(54, 162, 235, 1)'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: 'Overall Staff Achievement (Percentages)',
                        font: { size: 16 }
                    }
                },
                scales: {
                    r: {
                        angleLines: {
                            display: false
                        },
                        suggestedMin: 0,
                        suggestedMax: 100, // Max percentage
                        pointLabels: {
                            font: { size: 10 }
                        },
                        grid: {
                            circular: true
                        }
                    }
                }
            }
        });
    }

    function renderEmployeeSummaryCharts(employeeMetrics) {
        destroyChart(employeeActivityChart);
        destroyChart(employeeLeadSourceChart);
        destroyChart(employeeProductInterestChart);

        const chartContainer = document.createElement('div');
        chartContainer.className = 'employee-summary-charts-container'; // For styling charts together
        reportDisplay.appendChild(chartContainer);

        // Activity Type Breakdown Pie Chart
        const activityCanvas = document.createElement('canvas');
        activityCanvas.id = 'employeeActivityChart';
        chartContainer.appendChild(activityCanvas);

        employeeActivityChart = new Chart(activityCanvas, {
            type: 'pie',
            data: {
                labels: Object.keys(employeeMetrics.activityTypeBreakdown),
                datasets: [{
                    data: Object.values(employeeMetrics.activityTypeBreakdown),
                    backgroundColor: ['#4CAF50', '#2196F3', '#FFC107', '#E53935'], // Custom colors
                    hoverOffset: 4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: 'Activity Type Breakdown',
                        font: { size: 14 }
                    }
                }
            }
        });

        // Lead Sources Bar Chart
        const leadSourceCanvas = document.createElement('canvas');
        leadSourceCanvas.id = 'employeeLeadSourceChart';
        chartContainer.appendChild(leadSourceCanvas);

        employeeLeadSourceChart = new Chart(leadSourceCanvas, {
            type: 'bar',
            data: {
                labels: Object.keys(employeeMetrics.leadSources),
                datasets: [{
                    label: 'Number of Leads',
                    data: Object.values(employeeMetrics.leadSources),
                    backgroundColor: 'rgba(75, 192, 192, 0.6)',
                    borderColor: 'rgba(75, 192, 192, 1)',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: 'Lead Sources',
                        font: { size: 14 }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            precision: 0 // Ensure whole numbers for count
                        }
                    }
                }
            }
        });

        // Product Interest Bar Chart
        const productInterestCanvas = document.createElement('canvas');
        productInterestCanvas.id = 'employeeProductInterestChart';
        chartContainer.appendChild(productInterestCanvas);

        employeeProductInterestChart = new Chart(productInterestCanvas, {
            type: 'bar',
            data: {
                labels: Object.keys(employeeMetrics.productInterest),
                datasets: [{
                    label: 'Interest Count',
                    data: Object.values(employeeMetrics.productInterest),
                    backgroundColor: 'rgba(153, 102, 255, 0.6)',
                    borderColor: 'rgba(153, 102, 255, 1)',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: 'Product Interest',
                        font: { size: 14 }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            precision: 0 // Ensure whole numbers for count
                        }
                    }
                }
            }
        });
    }


    // --- Report Display Functions ---

    function displayReport(type, data) {
        reportDisplay.innerHTML = ''; // Clear previous report content
        // Destroy any active charts before rendering new content
        destroyChart(allBranchSnapshotChart);
        destroyChart(overallStaffPerformanceChart);
        destroyChart(employeeActivityChart);
        destroyChart(employeeLeadSourceChart);
        destroyChart(employeeProductInterestChart);

        if (type === 'allBranchSnapshot') {
            reportDisplay.innerHTML += '<h2>All Branch Snapshot</h2>';
            if (data.length === 0) {
                reportDisplay.innerHTML += '<p>No data available for All Branch Snapshot.</p>';
                return;
            }

            // Render Chart
            renderAllBranchSnapshotChart(data);

            // Render Table (below chart)
            let tableHTML = `<div class="data-table-container">`;
            tableHTML += `<table class="all-branch-snapshot-table">
                <thead>
                    <tr>
                        <th>Branch</th>
                        <th>Total Visits</th>
                        <th>Visit Achieved %</th>
                        <th>Total Calls</th>
                        <th>Call Achieved %</th>
                        <th>Total Products Sold</th>
                        <th>Product Achieved %</th>
                        <th>Total Value Achieved</th>
                        <th>Value Achieved %</th>
                    </tr>
                </thead>
                <tbody>`;

            data.forEach(branch => {
                const visitAchievement = calculateAchievement(branch.totalVisits, TARGETS['Visit']);
                const callAchievement = calculateAchievement(branch.totalCalls, TARGETS['Call']);
                const productAchievement = calculateAchievement(branch.totalProductsSold, TARGETS['Product Sold']);
                const valueAchievement = calculateAchievement(branch.totalValueAchieved, TARGETS['Value Achieved']);

                tableHTML += `<tr>
                    <td data-label="Branch">${branch.branch}</td>
                    <td data-label="Total Visits">${branch.totalVisits}</td>
                    <td data-label="Visit Achieved %">${createProgressBarHTML(visitAchievement)}</td>
                    <td data-label="Total Calls">${branch.totalCalls}</td>
                    <td data-label="Call Achieved %">${createProgressBarHTML(callAchievement)}</td>
                    <td data-label="Total Products Sold">${branch.totalProductsSold}</td>
                    <td data-label="Product Achieved %">${createProgressBarHTML(productAchievement)}</td>
                    <td data-label="Total Value Achieved">${branch.totalValueAchieved.toFixed(2)}</td>
                    <td data-label="Value Achieved %">${createProgressBarHTML(valueAchievement)}</td>
                </tr>`;
            });

            tableHTML += `</tbody></table></div>`;
            reportDisplay.innerHTML += tableHTML;

        } else if (type === 'overallStaffPerformance') {
            reportDisplay.innerHTML += '<h2>Overall Staff Performance</h2>';
            if (data.length === 0) {
                reportDisplay.innerHTML += '<p>No data available for Overall Staff Performance.</p>';
                return;
            }

            // Render Charts
            renderOverallStaffPerformanceChart(data);

            // Render Table (below charts)
            let tableHTML = `<div class="data-table-container">`;
            tableHTML += `<table class="performance-table">
                <thead>
                    <tr>
                        <th>Employee Name</th>
                        <th>Branch</th>
                        <th>Designation</th>
                        <th colspan="2">Visits</th>
                        <th colspan="2">Calls</th>
                        <th colspan="2">Products Sold</th>
                        <th colspan="2">Value Achieved</th>
                    </tr>
                    <tr>
                        <th></th>
                        <th></th>
                        <th></th>
                        <th>Count</th>
                        <th>% Achieved</th>
                        <th>Count</th>
                        <th>% Achieved</th>
                        <th>Count</th>
                        <th>% Achieved</th>
                        <th>Value</th>
                        <th>% Achieved</th>
                    </tr>
                </thead>
                <tbody>`;

            data.forEach(employee => {
                tableHTML += `<tr>
                    <td data-label="Employee Name">${employee.employeeName}</td>
                    <td data-label="Branch">${employee.branch}</td>
                    <td data-label="Designation">${employee.designation}</td>
                    <td data-label="Visits Count">${employee.visits}</td>
                    <td data-label="Visits Achieved %">${createProgressBarHTML(employee.visitAchievement)}</td>
                    <td data-label="Calls Count">${employee.calls}</td>
                    <td data-label="Calls Achieved %">${createProgressBarHTML(employee.callAchievement)}</td>
                    <td data-label="Products Sold Count">${employee.productsSold}</td>
                    <td data-label="Products Sold Achieved %">${createProgressBarHTML(employee.productAchievement)}</td>
                    <td data-label="Value Achieved">${employee.valueAchieved.toFixed(2)}</td>
                    <td data-label="Value Achieved %">${createProgressBarHTML(employee.valueAchievement)}</td>
                </tr>`;
            });

            tableHTML += `</tbody></table></div>`;
            reportDisplay.innerHTML += tableHTML;

        } else if (type === 'employeeSummary') {
            reportDisplay.innerHTML += `<h2>Summary for ${data.employeeName} (${data.employeeCode})</h2>`;
            reportDisplay.innerHTML += `<h3>Branch: ${data.branch} | Designation: ${data.designation}</h3>`;
            if (!data.metrics || (data.metrics.totalVisits === 0 && data.metrics.totalCalls === 0 && data.metrics.totalProductsSold === 0 && data.metrics.totalValueAchieved === 0)) {
                reportDisplay.innerHTML += `<p>No canvassing activity recorded for ${data.employeeName}.</p>`;
                return;
            }

            // Render Charts
            renderEmployeeSummaryCharts(data.metrics);


            // Render Summary Details (below charts)
            const summaryDetailsHTML = `
                <div class="summary-breakdown-card">
                    <div class="summary-details-container">
                        <h4>Overall Performance</h4>
                        <ul class="summary-list">
                            <li><strong>Total Visits:</strong> <span>${data.metrics.totalVisits}</span></li>
                            <li><strong>Visit Achieved %:</strong> <span>${createProgressBarHTML(calculateAchievement(data.metrics.totalVisits, TARGETS['Visit']))}</span></li>
                            <li><strong>Total Calls:</strong> <span>${data.metrics.totalCalls}</span></li>
                            <li><strong>Call Achieved %:</strong> <span>${createProgressBarHTML(calculateAchievement(data.metrics.totalCalls, TARGETS['Call']))}</span></li>
                            <li><strong>Total Products Sold:</strong> <span>${data.metrics.totalProductsSold}</span></li>
                            <li><strong>Product Achieved %:</strong> <span>${createProgressBarHTML(calculateAchievement(data.metrics.totalProductsSold, TARGETS['Product Sold']))}</span></li>
                            <li><strong>Total Value Achieved:</strong> <span>${data.metrics.totalValueAchieved.toFixed(2)}</span></li>
                            <li><strong>Value Achieved %:</strong> <span>${createProgressBarHTML(calculateAchievement(data.metrics.totalValueAchieved, TARGETS['Value Achieved']))}</span></li>
                        </ul>
                    </div>
                </div>
            `;
            reportDisplay.innerHTML += summaryDetailsHTML;

        } else if (type === 'allCanvassingEntries') {
            reportDisplay.innerHTML += `<h2>All Canvassing Entries for ${data.employeeName}</h2>`;
            if (data.entries.length === 0) {
                reportDisplay.innerHTML += `<p>No canvassing entries found for ${data.employeeName}.</p>`;
                return;
            }

            let tableHTML = `<div class="data-table-container">`;
            tableHTML += `<table class="all-entries-table">
                <thead>
                    <tr>
                        <th>Timestamp</th>
                        <th>Customer Name</th>
                        <th>Contact</th>
                        <th>Activity Type</th>
                        <th>Product Interest</th>
                        <th>Lead Source</th>
                        <th>Value Achieved</th>
                        <th>Remarks</th>
                    </tr>
                </thead>
                <tbody>`;

            data.entries.forEach(entry => {
                tableHTML += `<tr>
                    <td data-label="Timestamp">${entry[HEADER_TIMESTAMP]}</td>
                    <td data-label="Customer Name">${entry[HEADER_CUSTOMER_NAME]}</td>
                    <td data-label="Contact">${entry[HEADER_CUSTOMER_CONTACT]}</td>
                    <td data-label="Activity Type">${entry[HEADER_ACTIVITY_TYPE]}</td>
                    <td data-label="Product Interest">${entry[HEADER_PRODUCT_INTEREST]}</td>
                    <td data-label="Lead Source">${entry[HEADER_LEAD_SOURCE]}</td>
                    <td data-label="Value Achieved">${parseFloat(entry[HEADER_VALUE_ACHIEVED] || 0).toFixed(2)}</td>
                    <td data-label="Remarks">${entry[HEADER_REMARKS]}</td>
                </tr>`;
            });

            tableHTML += `</tbody></table></div>`;
            reportDisplay.innerHTML += tableHTML;
        } else if (type === 'nonParticipatingBranches') {
            reportDisplay.innerHTML = '<h2>Non-Participating Branches</h2>';
            if (data.length === 0) {
                reportDisplay.innerHTML += '<p>All branches have recorded canvassing activity!</p>';
            } else {
                reportDisplay.innerHTML += `<p class="no-participation-message">The following branches have not recorded any canvassing activity:</p>`;
                let listHTML = `<ul class="non-participating-branch-list">`;
                data.forEach(branch => {
                    listHTML += `<li>${branch}</li>`;
                });
                listHTML += `</ul>`;
                reportDisplay.innerHTML += listHTML;
            }
        }
    }


    // --- Event Listeners for Report Controls ---

    viewAllBranchSnapshotBtn.addEventListener('click', () => {
        const branchSnapshotData = [];
        allUniqueBranches.forEach(branchName => {
            const metrics = calculateBranchMetrics(branchName);
            branchSnapshotData.push({
                branch: branchName,
                ...metrics
            });
        });
        displayReport('allBranchSnapshot', branchSnapshotData);
    });

    viewOverallStaffPerformanceBtn.addEventListener('click', () => {
        const overallPerformanceData = [];
        allUniqueEmployees.forEach(employeeName => {
            const employeeMetrics = calculateEmployeeMetrics(employeeName);
            const employeeInfo = employeeDataMap[employeeName] || {};
            overallPerformanceData.push({
                employeeName: employeeName,
                employeeCode: employeeInfo[HEADER_EMPLOYEE_CODE] || 'N/A',
                branch: employeeInfo[HEADER_BRANCH] || 'N/A',
                designation: employeeInfo[HEADER_DESIGNATION] || 'N/A',
                visits: employeeMetrics.totalVisits,
                calls: employeeMetrics.totalCalls,
                productsSold: employeeMetrics.totalProductsSold,
                valueAchieved: employeeMetrics.totalValueAchieved,
                visitAchievement: calculateAchievement(employeeMetrics.totalVisits, TARGETS['Visit']),
                callAchievement: calculateAchievement(employeeMetrics.totalCalls, TARGETS['Call']),
                productAchievement: calculateAchievement(employeeMetrics.totalProductsSold, TARGETS['Product Sold']),
                valueAchievement: calculateAchievement(employeeMetrics.totalValueAchieved, TARGETS['Value Achieved'])
            });
        });
        displayReport('overallStaffPerformance', overallPerformanceData.sort((a, b) => a.employeeName.localeCompare(b.employeeName)));
    });


    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        const selectedEmployee = employeeSelect.value; // Keep current employee selection

        // Populate employee dropdown based on selected branch
        if (selectedBranch && branchEmployeeMap[selectedBranch]) {
            populateDropdown(employeeSelect, Array.from(branchEmployeeMap[selectedBranch]).sort(), true);
        } else {
            populateDropdown(employeeSelect, allUniqueEmployees, true); // Show all employees if no branch selected
        }
        // Restore previous employee selection if it's still valid for the new branch
        if (selectedEmployee && Array.from(employeeSelect.options).some(option => option.value === selectedEmployee)) {
             employeeSelect.value = selectedEmployee;
        } else {
            employeeSelect.value = ''; // Reset if not valid
        }

        // Show/hide employee filter panel and specific view buttons
        if (selectedBranch) {
            employeeFilterPanel.style.display = 'block';
            viewAllBranchSnapshotBtn.style.display = 'none';
            viewOverallStaffPerformanceBtn.style.display = 'none';
            viewEmployeeSummaryBtn.style.display = 'inline-block';
            viewAllCanvassingEntriesBtn.style.display = 'inline-block';
        } else {
            employeeFilterPanel.style.display = 'none';
            viewAllBranchSnapshotBtn.style.display = 'inline-block';
            viewOverallStaffPerformanceBtn.style.display = 'inline-block';
            viewEmployeeSummaryBtn.style.display = 'none';
            viewAllCanvassingEntriesBtn.style.display = 'none';
        }

        // Auto-trigger report if an employee is already selected
        if (selectedEmployee && selectedBranch) {
            employeeSelect.dispatchEvent(new Event('change')); // Trigger employee change to update report
        } else if (selectedBranch && !selectedEmployee) {
            // If only branch is selected, show branch performance summary
            const branchMetrics = calculateBranchMetrics(selectedBranch);
            displayReport('employeeSummary', { // Re-use employeeSummary format for branch overview
                employeeName: `${selectedBranch} Branch`,
                employeeCode: '',
                branch: selectedBranch,
                designation: 'Branch Performance Overview',
                metrics: {
                    totalVisits: branchMetrics.totalVisits,
                    totalCalls: branchMetrics.totalCalls,
                    totalProductsSold: branchMetrics.totalProductsSold,
                    totalValueAchieved: branchMetrics.totalValueAchieved,
                    activityTypeBreakdown: {}, // Not applicable at branch level
                    leadSources: {}, // Not applicable at branch level
                    productInterest: {} // Not applicable at branch level
                }
            });
        } else {
            reportDisplay.innerHTML = '<p>Select a branch or an employee to view reports.</p>';
            // If nothing selected, show default All Branch Snapshot
            viewAllBranchSnapshotBtn.click();
        }
    });

    employeeSelect.addEventListener('change', () => {
        const selectedEmployeeName = employeeSelect.value;
        if (selectedEmployeeName) {
            const employeeInfo = employeeDataMap[selectedEmployeeName] || {};
            const employeeMetrics = calculateEmployeeMetrics(selectedEmployeeName);
            displayReport('employeeSummary', {
                employeeName: selectedEmployeeName,
                employeeCode: employeeInfo[HEADER_EMPLOYEE_CODE],
                branch: employeeInfo[HEADER_BRANCH],
                designation: employeeInfo[HEADER_DESIGNATION],
                metrics: employeeMetrics
            });
        } else if (branchSelect.value) {
             // If employee selection cleared but branch is selected, show branch performance
            const branchMetrics = calculateBranchMetrics(branchSelect.value);
            displayReport('employeeSummary', { // Re-use employeeSummary format for branch overview
                employeeName: `${branchSelect.value} Branch`,
                employeeCode: '',
                branch: branchSelect.value,
                designation: 'Branch Performance Overview',
                metrics: {
                    totalVisits: branchMetrics.totalVisits,
                    totalCalls: branchMetrics.totalCalls,
                    totalProductsSold: branchMetrics.totalProductsSold,
                    totalValueAchieved: branchMetrics.totalValueAchieved,
                    activityTypeBreakdown: {},
                    leadSources: {},
                    productInterest: {}
                }
            });
        } else {
            reportDisplay.innerHTML = '<p>Select an employee to view their summary or select a branch for an overview.</p>';
        }
    });

    viewEmployeeSummaryBtn.addEventListener('click', () => {
        employeeSelect.dispatchEvent(new Event('change')); // Trigger summary display
    });

    viewAllCanvassingEntriesBtn.addEventListener('click', () => {
        const selectedEmployeeName = employeeSelect.value;
        if (selectedEmployeeName) {
            const employeeEntries = allCanvassingData.filter(entry => entry[HEADER_EMPLOYEE_NAME] === selectedEmployeeName);
            displayReport('allCanvassingEntries', {
                employeeName: selectedEmployeeName,
                entries: employeeEntries
            });
        } else {
            displayMessage('Please select an employee to view all their canvassing entries.', true);
        }
    });


    // --- Tab Navigation ---

    function showTab(activeTabId) {
        // Hide all sections
        reportsSection.style.display = 'none';
        detailedCustomerViewSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(btn => {
            btn.classList.remove('active');
        });

        // Activate selected tab and show corresponding section
        const activeTabButton = document.getElementById(activeTabId);
        if (activeTabButton) {
            activeTabButton.classList.add('active');
            if (activeTabId === 'allBranchSnapshotTabBtn' || activeTabId === 'allStaffOverallPerformanceTabBtn' || activeTabId === 'nonParticipatingBranchesTabBtn') {
                reportsSection.style.display = 'block';
                // Trigger initial report display for these tabs
                if (activeTabId === 'allBranchSnapshotTabBtn') {
                    branchSelect.value = ''; // Ensure no branch is selected
                    employeeSelect.value = ''; // Ensure no employee is selected
                    employeeFilterPanel.style.display = 'none';
                    viewAllBranchSnapshotBtn.style.display = 'inline-block';
                    viewOverallStaffPerformanceBtn.style.display = 'inline-block';
                    viewEmployeeSummaryBtn.style.display = 'none';
                    viewAllCanvassingEntriesBtn.style.display = 'none';
                    viewAllBranchSnapshotBtn.click(); // Automatically show snapshot
                } else if (activeTabId === 'allStaffOverallPerformanceTabBtn') {
                    branchSelect.value = ''; // Ensure no branch is selected
                    employeeSelect.value = ''; // Ensure no employee is selected
                    employeeFilterPanel.style.display = 'none';
                    viewAllBranchSnapshotBtn.style.display = 'inline-block';
                    viewOverallStaffPerformanceBtn.style.display = 'inline-block';
                    viewEmployeeSummaryBtn.style.display = 'none';
                    viewAllCanvassingEntriesBtn.style.display = 'none';
                    viewOverallStaffPerformanceBtn.click(); // Automatically show staff performance
                } else if (activeTabId === 'nonParticipatingBranchesTabBtn') {
                     // Calculate non-participating branches
                    const branchesWithActivity = new Set(allCanvassingData.map(entry => entry[HEADER_BRANCH]));
                    const nonParticipating = allUniqueBranches.filter(branch => !branchesWithActivity.has(branch));
                    displayReport('nonParticipatingBranches', nonParticipating);
                }
            } else if (activeTabId === 'detailedCustomerViewTabBtn') {
                detailedCustomerViewSection.style.display = 'flex'; // Use flex for this section
                // Clear customer details when tab is activated
                customerDetailsCard.innerHTML = '<p>Select a customer from the list to view their details.</p>';
                customerList.innerHTML = ''; // Clear customer list
                customerBranchSelect.value = ''; // Reset branch selection
                customerEmployeeSelect.value = ''; // Reset employee selection
            } else if (activeTabId === 'employeeManagementTabBtn') {
                employeeManagementSection.style.display = 'block';
                // Clear any previous management messages
                employeeManagementMessage.innerHTML = '';
            }
        }
    }

    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    detailedCustomerViewTabBtn.addEventListener('click', () => showTab('detailedCustomerViewTabBtn'));
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));


    // --- Detailed Customer View Logic ---
    let currentCustomerData = []; // To store filtered customer data

    customerBranchSelect.addEventListener('change', () => {
        const selectedBranch = customerBranchSelect.value;
        const selectedEmployee = customerEmployeeSelect.value; // Keep current employee selection

        // Filter employees for the customer employee dropdown based on selected branch
        if (selectedBranch && branchEmployeeMap[selectedBranch]) {
            populateDropdown(customerEmployeeSelect, Array.from(branchEmployeeMap[selectedBranch]).sort());
        } else {
            populateDropdown(customerEmployeeSelect, allUniqueEmployees); // Show all employees if no branch selected
        }

        // Restore previous employee selection if it's still valid for the new branch
        if (selectedEmployee && Array.from(customerEmployeeSelect.options).some(option => option.value === selectedEmployee)) {
             customerEmployeeSelect.value = selectedEmployee;
        } else {
            customerEmployeeSelect.value = ''; // Reset if not valid
        }

        filterAndDisplayCustomers();
    });

    customerEmployeeSelect.addEventListener('change', filterAndDisplayCustomers);

    function filterAndDisplayCustomers() {
        const selectedBranch = customerBranchSelect.value;
        const selectedEmployee = customerEmployeeSelect.value;

        currentCustomerData = allCanvassingData.filter(entry => {
            const matchesBranch = selectedBranch ? entry[HEADER_BRANCH] === selectedBranch : true;
            const matchesEmployee = selectedEmployee ? entry[HEADER_EMPLOYEE_NAME] === selectedEmployee : true;
            return matchesBranch && matchesEmployee;
        });

        // Group entries by customer to show unique customers
        const uniqueCustomersMap = new Map(); // Key: Customer Name + Contact, Value: Latest entry for that customer
        currentCustomerData.forEach(entry => {
            const customerKey = `${entry[HEADER_CUSTOMER_NAME]}-${entry[HEADER_CUSTOMER_CONTACT]}`;
            // For simplicity, we'll store the last entry for a customer if multiple exist,
            // or you could store an array of all entries for a customer.
            uniqueCustomersMap.set(customerKey, entry);
        });

        customerList.innerHTML = ''; // Clear previous list
        customerDetailsCard.innerHTML = '<p>Select a customer from the list to view their details.</p>';

        if (uniqueCustomersMap.size === 0) {
            customerList.innerHTML = '<li>No customers found for the selected criteria.</li>';
            return;
        }

        Array.from(uniqueCustomersMap.values()).forEach(entry => {
            const listItem = document.createElement('li');
            listItem.className = 'customer-list-item';
            listItem.textContent = entry[HEADER_CUSTOMER_NAME];
            listItem.dataset.customerName = entry[HEADER_CUSTOMER_NAME];
            listItem.dataset.customerContact = entry[HEADER_CUSTOMER_CONTACT];
            listItem.addEventListener('click', () => {
                // Remove active class from previous active item
                document.querySelectorAll('.customer-list-item').forEach(item => {
                    item.classList.remove('active');
                });
                // Add active class to clicked item
                listItem.classList.add('active');
                displayCustomerDetails(entry[HEADER_CUSTOMER_NAME], entry[HEADER_CUSTOMER_CONTACT]);
            });
            customerList.appendChild(listItem);
        });
    }

    function displayCustomerDetails(customerName, customerContact) {
        const relevantEntries = allCanvassingData.filter(entry =>
            entry[HEADER_CUSTOMER_NAME] === customerName && entry[HEADER_CUSTOMER_CONTACT] === customerContact
        ).sort((a, b) => new Date(b[HEADER_TIMESTAMP]) - new Date(a[HEADER_TIMESTAMP])); // Sort by timestamp, latest first

        if (relevantEntries.length === 0) {
            customerDetailsCard.innerHTML = '<p>No details found for this customer.</p>';
            return;
        }

        const latestEntry = relevantEntries[0]; // Get the most recent entry for main details

        customerDetailsCard.innerHTML = ''; // Clear previous content

        // Customer Profile Section
        const profileSection = document.createElement('div');
        profileSection.className = 'customer-info-section full-width-section';
        profileSection.innerHTML = `
            <h3>Customer Profile</h3>
            <div class="detail-row"><span class="detail-label">Name:</span> <span class="detail-value">${latestEntry[HEADER_CUSTOMER_NAME]}</span></div>
            <div class="detail-row"><span class="detail-label">Contact:</span> <span class="detail-value">${latestEntry[HEADER_CUSTOMER_CONTACT]}</span></div>
            <div class="detail-row"><span class="detail-label">Profession:</span> <span class="detail-value">${latestEntry[HEADER_CUSTOMER_PROFESSION] || 'N/A'}</span></div>
        `;
        customerDetailsCard.appendChild(profileSection);

        // Canvassing Activity Section
        const activitySection = document.createElement('div');
        activitySection.className = 'customer-info-section';
        let activityHTML = `
            <h3>Latest Canvassing Activity</h3>
            <div class="detail-row"><span class="detail-label">Timestamp:</span> <span class="detail-value">${latestEntry[HEADER_TIMESTAMP]}</span></div>
            <div class="detail-row"><span class="detail-label">Employee:</span> <span class="detail-value">${latestEntry[HEADER_EMPLOYEE_NAME]} (${latestEntry[HEADER_EMPLOYEE_CODE]})</span></div>
            <div class="detail-row"><span class="detail-label">Branch:</span> <span class="detail-value">${latestEntry[HEADER_BRANCH]}</span></div>
            <div class="detail-row"><span class="detail-label">Activity Type:</span> <span class="detail-value">${latestEntry[HEADER_ACTIVITY_TYPE]}</span></div>
            <div class="detail-row"><span class="detail-label">Call Outcome:</span> <span class="detail-value">${latestEntry[HEADER_CALL_OUTCOME] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Lead Source:</span> <span class="detail-value">${latestEntry[HEADER_LEAD_SOURCE] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Product Interest:</span> <span class="detail-value">${latestEntry[HEADER_PRODUCT_INTEREST] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Value Achieved:</span> <span class="detail-value">${parseFloat(latestEntry[HEADER_VALUE_ACHIEVED] || 0).toFixed(2)}</span></div>
            <div class="detail-row"><span class="detail-label">Status:</span> <span class="detail-value">${latestEntry[HEADER_STATUS] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Next Action:</span> <span class="detail-value">${latestEntry[HEADER_NEXT_ACTION] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Remarks:</span> <span class="detail-value remark-text">${latestEntry[HEADER_REMARKS] || 'No remarks.'}</span></div>
        `;
        activitySection.innerHTML = activityHTML;
        customerDetailsCard.appendChild(activitySection);

        // History Section
        if (relevantEntries.length > 1) {
            const historySection = document.createElement('div');
            historySection.className = 'customer-info-section full-width-section';
            historySection.innerHTML = `<h3>Activity History (${relevantEntries.length} entries)</h3>`;
            const historyList = document.createElement('ul');
            historyList.style.listStyle = 'none';
            historyList.style.padding = '0';
            historyList.style.maxHeight = '200px';
            historyList.style.overflowY = 'auto';
            historyList.style.border = '1px solid #eee';
            historyList.style.borderRadius = '5px';
            historyList.style.backgroundColor = '#f9f9f9';

            relevantEntries.forEach((entry, index) => {
                // Skip the latest entry as it's already displayed above
                if (index === 0) return;

                const li = document.createElement('li');
                li.style.padding = '10px';
                li.style.borderBottom = '1px dashed #ddd';
                if (index === relevantEntries.length -1 ) li.style.borderBottom = 'none';

                li.innerHTML = `
                    <p style="font-weight: bold; margin-bottom: 5px;">${entry[HEADER_TIMESTAMP]} by ${entry[HEADER_EMPLOYEE_NAME]}:</p>
                    <div class="detail-row"><span class="detail-label">Activity Type:</span> <span class="detail-value">${entry[HEADER_ACTIVITY_TYPE]}</span></div>
                    <div class="detail-row"><span class="detail-label">Product Interest:</span> <span class="detail-value">${entry[HEADER_PRODUCT_INTEREST] || 'N/A'}</span></div>
                    <div class="detail-row"><span class="detail-label">Value Achieved:</span> <span class="detail-value">${parseFloat(entry[HEADER_VALUE_ACHIEVED] || 0).toFixed(2)}</span></div>
                    <div class="detail-row"><span class="detail-label">Remarks:</span> <span class="detail-value remark-text">${entry[HEADER_REMARKS] || 'No remarks.'}</span></div>
                `;
                historyList.appendChild(li);
            });
            if (historyList.children.length > 0) { // Only append if there are actual history entries
                historySection.appendChild(historyList);
                customerDetailsCard.appendChild(historySection);
            }
        }
    }


    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: employeeNameInput.value.trim(),
                [HEADER_EMPLOYEE_CODE]: employeeCodeInput.value.trim(),
                [HEADER_BRANCH]: branchNameInput.value.trim(),
                [HEADER_DESIGNATION]: designationInput.value.trim()
            };

            // Basic validation
            if (!employeeData[HEADER_EMPLOYEE_NAME] || !employeeData[HEADER_EMPLOYEE_CODE] || !employeeData[HEADER_BRANCH] || !employeeData[HEADER_DESIGNATION]) {
                displayEmployeeManagementMessage('All fields are required for adding an employee.', true);
                return;
            }

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
            const bulkDetails = bulkEmployeeDetailsInput.value.trim();

            if (!branchName) {
                displayEmployeeManagementMessage('Branch Name is required for bulk entry.', true);
                return;
            }
            if (!bulkDetails) {
                displayEmployeeManagementMessage('Employee details are required for bulk entry.', true);
                return;
            }

            const employeesToAdd = [];
            const lines = bulkDetails.split('\n');
            for (const line of lines) {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length >= 2 && parts[0] && parts[1]) { // Name, Code (Designation optional)
                    const employeeData = {
                        [HEADER_EMPLOYEE_NAME]: parts[0],
                        [HEADER_EMPLOYEE_CODE]: parts[1],
                        [HEADER_BRANCH]: branchName,
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
    processData().then(() => {
        showTab('allBranchSnapshotTabBtn'); // Show default tab after data is processed
    });
});
