document.addEventListener('DOMContentLoaded', () => {
    // IMPORTANT: Replace with the actual public URL of your canvassing_data.json file from Google Drive
    // To get this: Go to your Google Drive folder, right-click on the canvassing_data.json file,
    // select "Share", then "Copy link". It will look like "https://drive.google.com/file/d/FILE_ID/view?usp=sharing".
    // You'll need to change 'view?usp=sharing' to 'export?format=json' or 'uc?export=download'
    // A common pattern is: 'https://drive.google.com/uc?id=YOUR_FILE_ID&export=download'
    const DATA_URL = "https://drive.google.com/uc?id=1mKgU6hdDKP8UauGRC6VM7DTR2nM9sOzX&export=download"; // Your verified DATA_URL

    fetch(DATA_URL)
        .then(response => {
            if (!response.ok) {
                // If the network response is not OK (e.g., 404, 500)
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json(); // Parse the JSON data from the response
        })
        .then(data => {
            console.log("Fetched data:", data); // Log fetched data to console for debugging

            // Display total submissions count
            const totalSubmissions = data.length;
            document.getElementById('totalSubmissions').textContent = totalSubmissions;

            // --- Generate Branch Reports ---
            const branchAggregates = {};
            data.forEach(entry => {
                const branch = entry['Branch Name']; // Use 'Branch Name' from your JSON
                if (branch) { // Only aggregate if branch name exists
                    if (!branchAggregates[branch]) {
                        branchAggregates[branch] = { submissions: 0 }; // Initialize with submissions count
                    }
                    branchAggregates[branch].submissions++;
                }
            });

            const branchList = document.getElementById('branchList');
            // Clear any existing list items to prevent duplicates on re-render
            branchList.innerHTML = '';
            for (const branch in branchAggregates) {
                const li = document.createElement('li');
                li.textContent = `${branch}: ${branchAggregates[branch].submissions} submissions`;
                branchList.appendChild(li);
            }

            // --- Generate Staff Reports ---
            const staffAggregates = {};
            data.forEach(entry => {
                const staff = entry['Employee Name']; // Use 'Employee Name' from your JSON
                if (staff) { // Only aggregate if staff name exists
                    if (!staffAggregates[staff]) {
                        staffAggregates[staff] = { submissions: 0 }; // Initialize with submissions count
                    }
                    staffAggregates[staff].submissions++;
                }
            });

            const staffList = document.getElementById('staffList');
            // Clear any existing list items
            staffList.innerHTML = '';
            for (const staff in staffAggregates) {
                const li = document.createElement('li');
                li.textContent = `${staff}: ${staffAggregates[staff].submissions} submissions`;
                staffList.appendChild(li);
            }

            // --- Generate "How Contacted" Reports ---
            const howContactedAggregates = {};
            data.forEach(entry => {
                const method = entry['How Contacted']; // Use 'How Contacted' from your JSON
                if (method) {
                    howContactedAggregates[method] = (howContactedAggregates[method] || 0) + 1;
                }
            });

            const howContactedList = document.getElementById('howContactedList'); // Assuming you have a <ul> with this ID in index.html
            if (howContactedList) { // Check if the element exists in HTML
                howContactedList.innerHTML = ''; // Clear existing list items
                for (const method in howContactedAggregates) {
                    const li = document.createElement('li');
                    li.textContent = `${method}: ${howContactedAggregates[method]} entries`;
                    howContactedList.appendChild(li);
                }
            } else {
                console.warn("Element with ID 'howContactedList' not found. Please add it to your index.html.");
            }

            // You can add more report sections here following the same pattern
            // For example, for "Activity Type" or "Profession"
            // You would need corresponding <ul> elements in your index.html (e.g., <ul id="activityTypeList">)

        })
        .catch(error => {
            // This block runs if fetching or parsing the data fails
            console.error("Error fetching or processing data:", error);
            // Display a user-friendly error message on the page
            const mainContent = document.querySelector('main');
            if (mainContent) {
                mainContent.innerHTML = "<p>Error loading data. Please try again later.</p>";
            } else {
                document.body.innerHTML = "<p>Error loading data. Please try again later.</p>";
            }
        });
});