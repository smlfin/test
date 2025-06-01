document.addEventListener('DOMContentLoaded', () => {
    // IMPORTANT: Replace with the actual public URL of your canvassing_data.json file from Google Drive
    // To get this: Go to your Google Drive folder, right-click on the canvassing_data.json file,
    // select "Share", then "Copy link". It will look like "https://drive.google.com/file/d/FILE_ID/view?usp=sharing".
    // You'll need to change 'view?usp=sharing' to 'export?format=json' or 'uc?export=download'
    // A common pattern is: 'https://drive.google.com/uc?id=YOUR_FILE_ID&export=download'
    const DATA_URL = "https://drive.google.com/uc?id=YOUR_FILE_ID&export=download"; // **IMPORTANT: Update this!**

    fetch(DATA_URL)
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .then(data => {
            console.log("Fetched data:", data); // Log data to console for debugging

            // --- Calculate Summaries ---
            const totalSubmissions = data.length;
            let totalDeposits = 0;
            // Assuming you have a column named 'Deposits Canvassed' or similar
            // Adjust this column name to match your Google Sheet
            data.forEach(entry => {
                const deposits = parseFloat(entry['Deposits Canvassed']); // Assuming it's a number
                if (!isNaN(deposits)) {
                    totalDeposits += deposits;
                }
            });

            document.getElementById('totalSubmissions').textContent = totalSubmissions;
            document.getElementById('totalDeposits').textContent = totalDeposits.toFixed(2); // Format to 2 decimal places

            // --- Generate Branch Reports ---
            const branchAggregates = {};
            data.forEach(entry => {
                const branch = entry['Branch Name']; // **IMPORTANT: Adjust column name**
                const deposits = parseFloat(entry['Deposits Canvassed']); // **IMPORTANT: Adjust column name**
                if (branch && !isNaN(deposits)) {
                    if (!branchAggregates[branch]) {
                        branchAggregates[branch] = { submissions: 0, totalDeposits: 0 };
                    }
                    branchAggregates[branch].submissions++;
                    branchAggregates[branch].totalDeposits += deposits;
                }
            });

            const branchList = document.getElementById('branchList');
            for (const branch in branchAggregates) {
                const li = document.createElement('li');
                li.textContent = `${branch}: ${branchAggregates[branch].submissions} submissions, ${branchAggregates[branch].totalDeposits.toFixed(2)} deposits`;
                branchList.appendChild(li);
            }

            // --- Generate Staff Reports ---
            const staffAggregates = {};
            data.forEach(entry => {
                const staff = entry['Staff Name']; // **IMPORTANT: Adjust column name**
                const deposits = parseFloat(entry['Deposits Canvassed']); // **IMPORTANT: Adjust column name**
                if (staff && !isNaN(deposits)) {
                    if (!staffAggregates[staff]) {
                        staffAggregates[staff] = { submissions: 0, totalDeposits: 0 };
                    }
                    staffAggregates[staff].submissions++;
                    staffAggregates[staff].totalDeposits += deposits;
                }
            });

            const staffList = document.getElementById('staffList');
            for (const staff in staffAggregates) {
                const li = document.createElement('li');
                li.textContent = `${staff}: ${staffAggregates[staff].submissions} submissions, ${staffAggregates[staff].totalDeposits.toFixed(2)} deposits`;
                staffList.appendChild(li);
            }

            // --- Add more reports here based on 'How they do it' and 'Prospect details' ---
            // You'll need to parse and aggregate that data similarly.
            // Example for 'How they do it':
            const howTheyDoItAggregates = {};
            data.forEach(entry => {
                const method = entry['How They Do It']; // **IMPORTANT: Adjust column name**
                if (method) {
                    howTheyDoItAggregates[method] = (howTheyDoItAggregates[method] || 0) + 1;
                }
            });
            console.log("How They Do It:", howTheyDoItAggregates); // Use this to see the structure and display accordingly

        })
        .catch(error => {
            console.error("Error fetching data:", error);
            document.querySelector('main').innerHTML = "<p>Error loading data. Please try again later.</p>";
        });
});