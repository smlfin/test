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
    const EMPLOYEE_MASTER_DATA_URL = "UNUSED"; // Marked as UNUSED for clarity, won't be fetched for reports

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
        'Seniors': { // Added Investment Staff with custom Visit target
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

    // Corrected PREDEFINED_EMPLOYEES from the user's uploaded "LIve staff may.xlsx - Sheet3.csv"
    // Sorted by employee name for consistent display in dropdowns
    const PREDEFINED_EMPLOYEES = [
        { employeeCode: "4566", employeeName: "ABIN K R", branchName: "Corporate Office", designation: "RECOVERY MANAGER" },
        { employeeCode: "3115", employeeName: "ABISH THOMAS", branchName: "Pathanamthitta", designation: "DEPUTY BRANCH MANAGER" },
        { employeeCode: "2055", employeeName: "ABHILASH K A", branchName: "Paravoor", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "4626", employeeName: "ABHAY B", branchName: "Kuzhalmannam", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "4101", employeeName: "AKHIL AMBADAN", branchName: "Perumbavoor", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "1185", employeeName: "AKHIL RAJ M", branchName: "Thodupuzha", designation: "SENIOR BRANCH MANAGER" },
        { employeeCode: "3969", employeeName: "AKSHAY A C", branchName: "Kottayam", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "4448", employeeName: "A AJMAL", branchName: "Corporate Office", designation: "ASSISTANT MANAGER" },
        { employeeCode: "4557", employeeName: "AJAY PB", branchName: "Corporate Office", designation: "SENIOR EXECUTIVE" },
        { employeeCode: "4529", employeeName: "AJI GEORGE", branchName: "Corporate Office", designation: "SENIOR EXECUTIVE" },
        { employeeCode: "1063", employeeName: "Ajitha Haridas", branchName: "HO KKM", designation: "Sr.Officer" },
        { employeeCode: "4609", employeeName: "AMAL K ANIL", branchName: "Paravoor", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "4573", employeeName: "AMBIKA KUMARY K B", branchName: "Nedumkandom", designation: "SENIOR BRANCH MANAGER" },
        { employeeCode: "4611", employeeName: "AMARJITH A S", branchName: "Kunnamkulam", designation: "ASSOCIATE VICE PRESIDENT" },
        { employeeCode: "1032", employeeName: "AMMINI JACOB", branchName: "Perumbavoor", designation: "Sr.EXECUTIVE" },
        { employeeCode: "4526", employeeName: "AMPILI P V", branchName: "Harippad", designation: "ABM" },
        { employeeCode: "4582", employeeName: "ANAND M B", branchName: "Corporate Office", designation: "RECOVERY MANAGER" },
        { employeeCode: "4487", employeeName: "ANANTHU RAJENDRAN", branchName: "Corporate Office", designation: "DRIVER" },
        { employeeCode: "4292", employeeName: "ANOL T ROY", branchName: "Kottayam", designation: "BRANCH MANAGER" },
        { employeeCode: "3705", employeeName: "ANISHA S NADHAN", branchName: "Pathanamthitta", designation: "OFFICER" },
        { employeeCode: "4550", employeeName: "ANITHA N S", branchName: "Kunnamkulam", designation: "BRANCH MANAGER" },
        { employeeCode: "4411", employeeName: "ANNIE JOSE", branchName: "Corporate Office", designation: "Accounts Officer" },
        { employeeCode: "1292", employeeName: "ANOOP M A", branchName: "Corporate Office", designation: "ASSISTANT MANAGER" },
        { employeeCode: "1411", employeeName: "ANOOP ANANTHAN NAIR", branchName: "Corporate Office", designation: "OFFICER" },
        { employeeCode: "4308", employeeName: "Anuprasad N", branchName: "Kuzhalmannam", designation: "ASSISTANT MANAGER" },
        { employeeCode: "4046", employeeName: "ARUNDEV T K", branchName: "Perumbavoor", designation: "SECURITY" },
        { employeeCode: "1414", employeeName: "ASHOK KUMAR K G", branchName: "Corporate Office", designation: "Sr.CREDIT MANAGER" },
        { employeeCode: "2949", employeeName: "ASHA RAJESH", branchName: "Thodupuzha", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "4459", employeeName: "ASHIN JENSON", branchName: "Corporate Office", designation: "ACCOUNTS EXECUTIVE" },
        { employeeCode: "4623", employeeName: "ASWIN KRISHNA K", branchName: "Harippad", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "4421", employeeName: "Aswathy Sasi", branchName: "Corporate Office", designation: "MIS OFFICER" },
        { employeeCode: "4425", employeeName: "ASWATHY K", branchName: "Harippad", designation: "ABM" },
        { employeeCode: "4631", employeeName: "ATHIRA K T", branchName: "Edappally", designation: "Brach Manager" },
        { employeeCode: "4394", employeeName: "Athul Augustine", branchName: "Corporate Office", designation: "DRIVER" },
        { employeeCode: "3944", employeeName: "BEENA C S", branchName: "Corporate Office", designation: "EXECUTIVE" },
        { employeeCode: "4325", employeeName: "Benedict Anto S M", branchName: "Corporate Office", designation: "MANAGER" },
        { employeeCode: "4589", employeeName: "BENCY JOSEPH", branchName: "Kattapana", designation: "ABM" },
        { employeeCode: "1218", employeeName: "BIJI REJI", branchName: "Perumbavoor", designation: "RELATIONSHIP EXECUTIVE" },
        { employeeCode: "4278", employeeName: "Bimi Micle P", branchName: "HO KKM", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "4495", employeeName: "BINOJDAS V H", branchName: "Corporate Office", designation: "VICE PRESIDENT" },
        { employeeCode: "1441", employeeName: "BINSY JENSON", branchName: "Corporate Office", designation: "Sr.EXECUTIVE" },
        { employeeCode: "4465", employeeName: "BHAGYALAKSHMI E S", branchName: "Corporate Office", designation: "ACCOUNTS EXECUTIVE" },
        { employeeCode: "4627", employeeCode: "CHINJU VAISHAK", branchName: "Nenmara", designation: "OFFICER" },
        { employeeCode: "1078", employeeName: "CHITHRA AJITH", branchName: "Corporate Office", designation: "Sr.EXECUTIVE" },
        { employeeCode: "4560", employeeName: "DHANYA K", branchName: "Alathur", designation: "LOAN EXECUTIVE" },
        { employeeCode: "4333", employeeName: "Dhanya L", branchName: "Corporate Office", designation: "Accounts manager" },
        { employeeCode: "4501", employeeName: "DHANYAMOL K A", branchName: "Kottayam", designation: "LOAN OFFICER" },
        { employeeCode: "4599", employeeName: "DILJITH P J", branchName: "HO KKM", designation: "DRIVER" },
        { employeeCode: "2778", employeeName: "DINIYA STEPHEN", branchName: "Corporate Office", designation: "TELECALLING EXECUTIVE" },
        { employeeCode: "4290", employeeName: "Dipin C G", branchName: "Corporate Office", designation: "DRIVER" },
        { employeeCode: "1098", employeeName: "EBY SCARIA", branchName: "Corporate Office", designation: "MANAGER - POST VERIFICATION" },
        { employeeCode: "4528", employeeName: "ELIZABETH JANCY P M", branchName: "Angamaly", designation: "SENIOR BRANCH MANAGER" },
        { employeeCode: "4426", employeeName: "FABIN ANTONY", branchName: "Corporate Office", designation: "ASSISTANT MANAGER" },
        { employeeCode: "4412", employeeName: "Fathima Ripsana A M", branchName: "Corporate Office", designation: "Accounts Executive" },
        { employeeCode: "4575", employeeName: "GANESH K S", branchName: "Corporate Office", designation: "Sr.VICE PRESIDENT" },
        { employeeCode: "4389", employeeName: "GEETHU GOPAN", branchName: "Corporate Office", designation: "EXECUTIVE" },
        { employeeCode: "4513", employeeName: "GINCE GEORGE", branchName: "Corporate Office", designation: "SENIOR MANAGER" },
        { employeeCode: "1918", employeeName: "GOWRI R", branchName: "Corporate Office", designation: "Sr.AVP" },
        { employeeCode: "4514", employeeName: "GREESHMA VIJAYAN", branchName: "Mavelikara", designation: "MANAGER" },
        { employeeCode: "4423", employeeName: "Harikrishnan R", branchName: "Nedumkandom", designation: "LOAN OFFICER" },
        { employeeCode: "1528", employeeName: "Hridya", branchName: "HO KKM", designation: "MANAGER" },
        { employeeCode: "1249", employeeName: "JISHA MANOJ", branchName: "Perumbavoor", designation: "CUSTOMER SERVICE EXECUTIVE" },
        { employeeCode: "4298", employeeName: "JISHA K J", branchName: "Kunnamkulam", designation: "SALES CO - ORDINATOR" },
        { employeeCode: "1299", employeeName: "JIBI BIJI", branchName: "Pathanamthitta", designation: "CUSTOMER SERVICE EXECUTIVE" },
        { employeeCode: "1173", employeeName: "Jiji Reni", branchName: "Kunnamkulam", designation: "OFFICER" },
        { employeeCode: "4554", employeeName: "JINCY K RAJU", branchName: "Kattapana", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "4614", employeeName: "JISHNU P L", branchName: "HO KKM", designation: "RECOVERY MANAGER" },
        { employeeCode: "4479", employeeName: "JITHU K JOSE", branchName: "Corporate Office", designation: "Sr.EXECUTIVE" },
        { employeeCode: "4628", employeeName: "JEFERSON M S", branchName: "Kunnamkulam", designation: "DRIVER" },
        { employeeCode: "1051", employeeName: "JAIMOL BENNY", branchName: "Perumbavoor", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "1018", employeeName: "Janus Samuel", branchName: "HO KKM", designation: "Sr.Officer" },
        { employeeCode: "1452", employeeName: "JESSY ROY", branchName: "Corporate Office", designation: "MANAGER" },
        { employeeCode: "2073", employeeName: "JOLLY JOY", branchName: "Corporate Office", designation: "SENIOR EXECUTIVE (DCR TRACKER)" },
        { employeeCode: "4080", employeeName: "KARTHIK K", branchName: "Thiruvalla", designation: "BRANCH MANAGER" },
        { employeeCode: "4595", employeeName: "KIRAN K BALAN", branchName: "HO KKM", designation: "Litigation Officer(Legal)" },
        { employeeCode: "1245", employeeName: "KUMARI K N", branchName: "Corporate Office", designation: "Sr.EXECUTIVE" },
        { employeeCode: "1141", employeeName: "LIJOY V R", branchName: "Corporate Office", designation: "ASSISTANT MANAGER (MIS & OPERATIONS)" },
        { employeeCode: "4420", employeeName: "Lishamol K Kudakkachira", branchName: "Kattapana", designation: "BRANCH MANAGER" },
        { employeeCode: "4527", employeeName: "MAHESH KUMAR M", branchName: "Mavelikara", designation: "COLLECTION MANAGER" },
        { employeeCode: "4561", employeeName: "MANJULA B", branchName: "Alathur", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "4532", employeeName: "MANJU B NAIR", branchName: "Kottayam", designation: "LOAN EXECUTIVE" },
        { employeeCode: "4365", employeeName: "Maneesh Purushothaman", branchName: "Thiruvalla", designation: "OFFICER" },
        { employeeCode: "4104", employeeName: "MANITHA MOHANAN", branchName: "Corporate Office", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "4583", employeeName: "MANU MOHAN", branchName: "Corporate Office", designation: "MANAGER" },
        { employeeCode: "4441", employeeName: "Manoj N Das", branchName: "Corporate Office", designation: "Chief Manager" },
        { employeeCode: "4632", employeeName: "MARIA SONY", branchName: "Corporate Office", designation: "HR Executive" },
        { employeeCode: "1564", employeeName: "MARY LITTLE FLOWER", branchName: "Corporate Office", designation: "Legal Officer" },
        { employeeCode: "4340", employeeName: "Mary V A", branchName: "Mattanchery", designation: "BRANCH MANAGER" },
        { employeeCode: "1175", employeeName: "MARIYAMMA GEORGE", branchName: "Pathanamthitta", designation: "SENIOR MANAGER" },
        { employeeCode: "4385", employeeName: "MINI JAYAN NAIR", branchName: "Koduvayur", designation: "ABM" },
        { employeeCode: "1659", employeeName: "MINI PAUL", branchName: "Muvattupuzha", designation: "ASSISTANT BRANCH MANAGER" },
        { employeeCode: "1125", employeeName: "NARAYANAN K", branchName: "Corporate Office", designation: "ASSOCIATE VICE PRESIDENT" },
        { employeeCode: "4361", employeeName: "NAVEEN P", branchName: "Thiruwillamala", designation: "OFFICER" },
        { employeeCode: "3121", employeeName: "NIMISHA RAJEEV", branchName: "Corporate Office", designation: "Legal Officer" },
        { employeeCode: "4621", employeeName: "NIMISHA K U", branchName: "Kunnamkulam", designation: "OFFICER" },
        { employeeCode: "2733", employeeName: "NITHIN T D", branchName: "Mattanchery", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "4242", employeeName: "Niveeth P V", branchName: "Corporate Office", designation: "IT SUPPORT" },
        { employeeCode: "4337", employeeName: "Prashob P P", branchName: "HO KKM", designation: "DRIVER" },
        { employeeCode: "4486", employeeName: "PRAYAG VINOD", branchName: "Corporate Office", designation: "ACCOUNTS EXECUTIVE" },
        { employeeCode: "1261", employeeName: "Priya Mohanan", branchName: "HO KKM", designation: "Receptionist" },
        { employeeCode: "4057", employeeName: "PRIYESH P P", branchName: "Perumbavoor", designation: "BRANCH HEAD" },
        { employeeCode: "1057", employeeName: "PROMOD M", branchName: "Corporate Office", designation: "MANAGER" },
        { employeeCode: "4512", employeeName: "PUSHYA U", branchName: "Kuzhalmannam", designation: "LOAN OFFICER" },
        { employeeCode: "1109", employeeName: "RADHIKA C S", branchName: "Paravoor", designation: "CUSTOMER SERVICE EXECUTIVE" },
        { employeeCode: "1024", employeeName: "Ramakrishnan", branchName: "HO KKM", designation: "Mechanic" },
        { employeeCode: "4444", employeeName: "Ranjith E N", branchName: "Corporate Office", designation: "Legal Co-ordinator" },
        { employeeCode: "3966", employeeName: "RATHEESH KUMAR N R", branchName: "Thiruvalla", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "4515", employeeName: "REMYA N", branchName: "Thiruwillamala", designation: "LOAN EXECUTIVE" },
        { employeeCode: "4416", employeeName: "Renju Unnikrishnan", branchName: "Chengannur", designation: "BRANCH MANAGER" },
        { employeeCode: "4450", employeeName: "Renjini G", branchName: "Chengannur", designation: "LOAN EXECUTIVE" },
        { employeeCode: "4496", employeeName: "RESHMA R", branchName: "Nedumkandom", designation: "LOAN OFFICER" },
        { employeeCode: "3982", employeeName: "RINU C W", branchName: "Mattanchery", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "4482", employeeName: "ROOPESH V P", branchName: "Thodupuzha", designation: "ASSISTANT MANAGER" },
        { employeeCode: "1187", employeeName: "SABITHA VENU", branchName: "Corporate Office", designation: "TEAM LEADER" },
        { employeeCode: "4498", employeeName: "SAGEESH KUMAR S", branchName: "Alathur", designation: "LOAN OFFICER" },
        { employeeCode: "4272", employeeName: "Sandhya S", branchName: "Harippad", designation: "BRANCH MANAGER" },
        { employeeCode: "4347", employeeName: "Sangeeth Das", branchName: "Angamaly", designation: "LOAN OFFICER" },
        { employeeCode: "4434", employeeName: "Sanjay Sajan", branchName: "Corporate Office", designation: "EXECUTIVE" },
        { employeeCode: "4547", employeeName: "SANTHI KRISHNA", branchName: "Paravoor", designation: "BRANCH MANAGER" },
        { employeeCode: "4588", employeeName: "SANJITH K", branchName: "Koduvayur", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "4256", employeeName: "Sareena M", branchName: "Alathur", designation: "BRANCH MANAGER" },
        { employeeCode: "1064", employeeName: "SARITHA K T", branchName: "Corporate Office", designation: "Asst.Manager" },
        { employeeCode: "1610", employeeName: "SARITHA M S", branchName: "Paravoor", designation: "Sr.EXECUTIVE" },
        { employeeCode: "4062", employeeName: "SARIGA N S", branchName: "Paravoor", designation: "OFFICER" },
        { employeeCode: "4581", employeeName: "SARATH M R", branchName: "Koduvayur", designation: "LOAN EXECUTIVE" },
        { employeeCode: "4620", employeeName: "SAFA ESSA V M", branchName: "Mattanchery", designation: "OFFICER" },
        { employeeCode: "4615", employeeName: "SETHU RAJ", branchName: "Kattapana", designation: "BUSINESS DEVELOPMENT MANAGER" },
        { employeeCode: "4630", employeeName: "SHABANAM K U", branchName: "Edappally", designation: "OFFICER" },
        { employeeCode: "1009", employeeName: "Shajan.A.D", branchName: "HO KKM", designation: "AVP" },
        { employeeCode: "1409", employeeName: "SHAIJO JOSE", branchName: "Corporate Office", designation: "OFFICER" },
        { employeeCode: "4344", employeeName: "Shainy Santhosh", branchName: "Thodupuzha", designation: "MANAGER" },
        { employeeCode: "4110", employeeName: "SHILPA JIBIN", branchName: "Corporate Office", designation: "MIS COORDINATOR" },
        { employeeCode: "4603", employeeName: "SIJO K CHERIAN", branchName: "Muvattupuzha", designation: "BUSINESS DEVELOPMENT MANAGER" },
        { employeeCode: "4468", employeeName: "SIJO M J", branchName: "Corporate Office", designation: "SENIOR EXECUTIVE" },
        { employeeCode: "3948", employeeName: "SINI", branchName: "Corporate Office", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "1049", employeeName: "SIVARAJ KUMAR K P", branchName: "Corporate Office", designation: "SENIOR MANAGER" },
        { employeeCode: "1222", employeeName: "SIVAKALA DILEEP", branchName: "Edappally", designation: "CUSTOMER SERVICE EXECUTIVE" },
        { employeeCode: "4503", employeeName: "SMITHA M", branchName: "Mavelikara", designation: "LOAN EXECUTIVE" },
        { employeeCode: "4392", employeeName: "RIYAS A", branchName: "Nenmara", designation: "OFFICER" },
        { employeeCode: "4378", employeeName: "SOBHI BABU", branchName: "Corporate Office", designation: "HOUSE KEEPER" },
        { employeeCode: "4362", employeeName: "SONIA S", branchName: "Thiruwillamala", designation: "OFFICER" },
        { employeeCode: "4531", employeeName: "SOWRAV KRISHNA", branchName: "Nenmara", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "1118", employeeName: "Sreeja Jaison", branchName: "HO KKM", designation: "MANAGER" },
        { employeeCode: "1530", employeeName: "Sreejitha K K", branchName: "Kunnamkulam", designation: "LOAN OFFICER" },
        { employeeCode: "4544", employeeName: "SREEKALA P R", branchName: "Mavelikara", designation: "BRANCH MANAGER" },
        { employeeCode: "4610", employeeName: "SREEJITH K C", branchName: "Kunnamkulam", designation: "ABM" },
        { employeeCode: "4612", employeeName: "SRUTHY N S", branchName: "Kunnamkulam", designation: "LOAN OFFICER" },
        { employeeCode: "2275", employeeName: "Shreekumar S", branchName: "HO KKM", designation: "Sr.AVP" },
        { employeeCode: "2518", employeeName: "Sunny", branchName: "Thodupuzha", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "4366", employeeName: "Supriya P", branchName: "Kuzhalmannam", designation: "BRANCH MANAGER" },
        { employeeCode: "4592", employeeName: "SYAM KUMAR M S", branchName: "Muvattupuzha", designation: "BRANCH MANAGER" },
        { employeeCode: "4516", employeeName: "SYAMILY M NAIR", branchName: "Corporate Office", designation: "ASSISTANT MANAGER" },
        { employeeCode: "1130", employeeName: "SWAPNA PRASANTH", branchName: "Paravoor", designation: "LOAN OFFICER" },
        { employeeCode: "4602", employeeName: "SWATHY P P", branchName: "Kottayam", designation: "LOAN EXECUTIVE" },
        { employeeCode: "1025", employeeName: "Tony K F", branchName: "HO KKM", designation: "Sr.AVP" },
        { employeeCode: "4309", employeeName: "Tony George", branchName: "Corporate Office", designation: "MAINTENANCE EXECUTIVE" },
        { employeeCode: "1055", employeeName: "ULLAS A N", branchName: "Thodupuzha", designation: "MANAGER" },
        { employeeCode: "4431", employeeName: "VEENA KRISHNAN", branchName: "Chengannur", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "4391", employeeName: "VIDYA V", branchName: "Nenmara", designation: "OFFICER" },
        { employeeCode: "4235", employeeName: "Vijesh K V", branchName: "Corporate Office", designation: "MANAGER" },
        { employeeCode: "4451", employeeName: "Vindhya K R", branchName: "Pathanamthitta", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "4336", employeeName: "Vineesh T", branchName: "HO KKM", designation: "Sr.AVP" },
        { employeeCode: "4086", employeeName: "VISHNU T R", branchName: "Thiruvalla", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "4497", employeeName: "VISHNU S SANTHOSH", branchName: "Corporate Office", designation: "ACCOUNTS EXECUTIVE" },
        { employeeCode: "4128", employeeName: "SHIBEENA V M", branchName: "Corporate Office", designation: "LEGAL MANAGER" },
        { employeeCode: "4622", employeeName: "SHYAM KRISHNAN", branchName: "Kattapana", designation: "Collection Executive" },
        { employeeCode: "4369", employeeName: "Aju M D", branchName: "Pathanamthitta", designation: "Collection Executive" }
    ].sort((a, b) => a.employeeName.localeCompare(b.employeeName));


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
    const HEADER_PHONE_NUMBER_WHATSAPP = 'Phone Number(Whatsapp)'; // Corrected typo: "Numebr" to "Number"
    const HEADER_ADDRESS = 'Address';
    const HEADER_PROFESSION = 'Profession';
    const HEADER_DOB_WD = 'DOB/WD';
    const HEADER_PRODUCT_INTERESTED = 'Product Interested'; // Corrected typo: "Prodcut" to "Product"
    const HEADER_REMARKS = 'Remarks';
    const HEADER_NEXT_FOLLOW_UP_DATE = 'Next Follow-up Date';
    const HEADER_RELATION_WITH_STAFF = 'Relation With Staff';
    // NEW: Customer Detail Headers as provided by user
    const HEADER_FAMILY_DETAILS_1 = 'Family Deatils -1 Name of wife/Husband';
    const HEADER_FAMILY_DETAILS_2 = 'Family Deatils -2 Job of wife/Husband';
    const HEADER_FAMILY_DETAILS_3 = 'Family Deatils -3 Names of Children';
    const HEADER_FAMILY_DETAILS_4 = 'Family Deatils -4 Deatils of Children';
    const HEADER_PROFILE_OF_CUSTOMER = 'Profile of Customer';


    // *** DOM Elements ***
    const branchSelect = document.getElementById('branchSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const employeeSelect = document.getElementById('employeeSelect');
    const viewOptions = document.getElementById('viewOptions');
    const viewBranchPerformanceReportBtn = document.getElementById('viewBranchPerformanceReportBtn');
    const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
    const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
    const viewPerformanceReportBtn = document.getElementById('viewPerformanceReportBtn');

    // Main Report Display Area
    const reportDisplay = document.getElementById('reportDisplay');
    // Dedicated message area element
    const statusMessageDiv = document.getElementById('statusMessage');


    // Tab buttons for main navigation
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const nonParticipatingEmployeesTabBtn = document.getElementById('nonParticipatingEmployeesTabBtn'); // NEW: Non-Participating Employees tab
    const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn'); // NEW
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    // Main Content Sections to toggle
    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection'); // NEW
    const employeeManagementSection = document.getElementById('employeeManagementSection');

    // NEW: Detailed Customer View Elements
    const customerViewBranchSelect = document.getElementById('customerViewBranchSelect');
    const customerViewEmployeeSelect = document.getElementById('customerViewEmployeeSelect');
    const customerCanvassedList = document.getElementById('customerCanvassedList');
    const customerDetailsContent = document.getElementById('customerDetailsContent');


    // Employee Management Form Elements
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const newEmployeeNameInput = document.getElementById('newEmployeeName');
    const newEmployeeCodeInput = document.getElementById('newEmployeeCode');
    const newBranchNameInput = document.getElementById('newBranchName');
    const newDesignationInput = document.getElementById('newDesignation');
    const employeeManagementMessage = document.getElementById('employeeManagementMessage');

    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsTextarea = document.getElementById('bulkEmployeeDetails');

    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');


    // Global variables to store fetched data
    let allCanvassingData = []; // Raw activity data from Form Responses 2
    let allUniqueBranches = []; // Will be populated dynamically from PREDEFINED_EMPLOYEES
    let allUniqueEmployees = []; // Employee codes from PREDEFINED_EMPLOYEES
    let employeeCodeToNameMap = {}; // {code: name} from PREDEFINED_EMPLOYEES
    let employeeCodeToDesignationMap = {}; // {code: designation} from PREDEFINED_EMPLOYEES
    let selectedBranchEntries = []; // Activity entries filtered by branch (for main reports section)
    let selectedEmployeeCodeEntries = []; // Activity entries filtered by employee code (for main reports section)


    // Utility to format date to ISO-MM-DD
    const formatDate = (dateString) => {
        if (!dateString) return '';
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString;
        return date.toISOString().split('T')[0];
    };

    // Helper to clear and display messages in a specific div (now targets statusMessageDiv)
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
            displayMessage("Activity data fetched successfully!", 'success');
        } catch (error) {
            console.error("Error fetching canvassing data:", error);
            displayMessage(`Error fetching activity data: ${error.message}`, 'error');
            allCanvassingData = []; // Ensure data is cleared on error
        }
    }

    // Function to parse CSV text into an array of objects
    function parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        if (lines.length === 0) return [];

        const headers = lines[0].split(',').map(header => header.trim().replace(/"/g, ''));
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const currentLine = lines[i];
            // Handle commas within quoted fields (simple approach: split by comma, then merge quoted parts)
            // A more robust CSV parser would be needed for complex cases with escaped quotes.
            const values = [];
            let inQuote = false;
            let start = 0;
            for (let j = 0; j < currentLine.length; j++) {
                if (currentLine[j] === '"') {
                    inQuote = !inQuote;
                } else if (currentLine[j] === ',' && !inQuote) {
                    values.push(currentLine.substring(start, j).trim().replace(/"/g, ''));
                    start = j + 1;
                }
            }
            values.push(currentLine.substring(start).trim().replace(/"/g, '')); // Add last value

            if (values.length === headers.length) {
                let row = {};
                headers.forEach((header, index) => {
                    row[header] = values[index];
                });
                data.push(row);
            } else {
                console.warn("Skipping malformed row:", currentLine);
            }
        }
        return data;
    }

    // Function to initialize branch and employee dropdowns dynamically from PREDEFINED_EMPLOYEES
    function initializeFilters() {
        // Clear previous options
        if (branchSelect) branchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        if (customerViewBranchSelect) customerViewBranchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        if (employeeSelect) employeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';
        if (customerViewEmployeeSelect) customerViewEmployeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';

        // Dynamically get unique branches from PREDEFINED_EMPLOYEES
        const uniqueBranchesSet = new Set();
        PREDEFINED_EMPLOYEES.forEach(employee => {
            uniqueBranchesSet.add(employee.branchName);
        });
        allUniqueBranches = Array.from(uniqueBranchesSet).sort(); // Convert to array and sort

        allUniqueBranches.forEach(branch => {
            const option = document.createElement('option');
            option.value = branch;
            option.textContent = branch;
            if (branchSelect) branchSelect.appendChild(option);
            if (customerViewBranchSelect) customerViewBranchSelect.appendChild(option.cloneNode(true)); // For Detailed Customer View
        });

        // Populate employee dropdown based on PREDEFINED_EMPLOYEES
        allUniqueEmployees = []; // Reset
        employeeCodeToNameMap = {}; // Reset
        employeeCodeToDesignationMap = {}; // Reset

        PREDEFINED_EMPLOYEES.forEach(employee => {
            allUniqueEmployees.push(employee.employeeCode);
            employeeCodeToNameMap[employee.employeeCode] = employee.employeeName;
            employeeCodeToDesignationMap[employee.employeeCode] = employee.designation;

            const option = document.createElement('option');
            option.value = employee.employeeCode;
            option.textContent = `${employee.employeeName} (${employee.employeeCode}) - ${employee.branchName}`;
            if (employeeSelect) employeeSelect.appendChild(option);
            if (customerViewEmployeeSelect) customerViewEmployeeSelect.appendChild(option.cloneNode(true)); // For Detailed Customer View
        });
    }


    // Filter data based on selected branch and employee for main reports section
    function filterDataForReports() {
        const selectedBranch = branchSelect ? branchSelect.value : '';
        const selectedEmployeeCode = employeeSelect ? employeeSelect.value : '';

        selectedBranchEntries = [];
        selectedEmployeeCodeEntries = [];

        if (selectedBranch) {
            selectedBranchEntries = allCanvassingData.filter(entry =>
                entry[HEADER_BRANCH_NAME] === selectedBranch
            );
        } else {
            selectedBranchEntries = allCanvassingData; // If no branch selected, consider all for branch-level stats
        }

        if (selectedEmployeeCode) {
            selectedEmployeeCodeEntries = allCanvassingData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode
            );
            employeeFilterPanel.style.display = 'block'; // Show employee filter if an employee is selected
        } else {
            selectedEmployeeCodeEntries = selectedBranchEntries; // If no employee selected, use branch filtered data
            employeeFilterPanel.style.display = 'none'; // Hide if no employee selected
        }

        // Show/hide relevant view options
        if (selectedBranch && selectedEmployeeCode) {
            viewBranchPerformanceReportBtn.style.display = 'none';
            viewEmployeeSummaryBtn.style.display = 'block';
            viewAllEntriesBtn.style.display = 'block';
        } else if (selectedBranch && !selectedEmployeeCode) {
            viewBranchPerformanceReportBtn.style.display = 'block';
            viewEmployeeSummaryBtn.style.display = 'none';
            viewAllEntriesBtn.style.display = 'block';
        } else {
            viewBranchPerformanceReportBtn.style.display = 'none';
            viewEmployeeSummaryBtn.style.display = 'none';
            viewAllEntriesBtn.style.display = 'block';
        }
        viewPerformanceReportBtn.style.display = 'none'; // This button seems redundant now
    }


    // --- Report Generation Functions ---

    function generateAllBranchSnapshotReport() {
        reportDisplay.innerHTML = '<h2>All Branch Snapshot Report</h2>';

        const branchActivity = {};
        allUniqueBranches.forEach(branch => { // Use dynamically generated branches
            branchActivity[branch] = {
                'Visit': 0,
                'Call': 0,
                'Reference': 0,
                'New Customer Leads': 0
            };
        });

        allCanvassingData.forEach(entry => {
            const branch = entry[HEADER_BRANCH_NAME];
            const activity = entry[HEADER_ACTIVITY_TYPE];
            const typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER];

            if (branchActivity[branch]) {
                if (activity === 'Visit') {
                    branchActivity[branch]['Visit']++;
                } else if (activity === 'Call') {
                    branchActivity[branch]['Call']++;
                } else if (activity === 'Reference') {
                    branchActivity[branch]['Reference']++;
                }

                if (typeOfCustomer === 'New Customer Lead') {
                    branchActivity[branch]['New Customer Leads']++;
                }
            }
        });

        let tableHtml = '<table class="report-table all-branch-snapshot-table"><thead><tr>';
        tableHtml += '<th>Branch Name</th><th>Total Visits</th><th>Total Calls</th><th>Total References</th><th>Total New Customer Leads</th>';
        tableHtml += '</tr></thead><tbody>';

        // Sort branches alphabetically
        const sortedBranches = Object.keys(branchActivity).sort();

        sortedBranches.forEach(branch => {
            const data = branchActivity[branch];
            tableHtml += `<tr>
                <td data-label="Branch Name">${branch}</td>
                <td data-label="Visits">${data['Visit']}</td>
                <td data-label="Calls">${data['Call']}</td>
                <td data-label="References">${data['Reference']}</td>
                <td data-label="New Customer Leads">${data['New Customer Leads']}</td>
            </tr>`;
        });

        tableHtml += '</tbody></table>';
        reportDisplay.innerHTML += tableHtml;
    }


    function generateBranchPerformanceReport() {
        const selectedBranch = branchSelect.value;
        if (!selectedBranch) {
            reportDisplay.innerHTML = '<p>Please select a branch to view its performance report.</p>';
            return;
        }

        reportDisplay.innerHTML = `<h2>Performance Report for ${selectedBranch}</h2>`;

        const employeesInBranch = PREDEFINED_EMPLOYEES.filter(emp => emp.branchName === selectedBranch);

        const employeePerformance = {};

        employeesInBranch.forEach(emp => {
            employeePerformance[emp.employeeCode] = {
                name: emp.employeeName,
                designation: emp.designation,
                'Visit': 0,
                'Call': 0,
                'Reference': 0,
                'New Customer Leads': 0,
                targets: TARGETS[emp.designation] || TARGETS['Default']
            };
        });


        selectedBranchEntries.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const activity = entry[HEADER_ACTIVITY_TYPE];
            const typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER];

            if (employeePerformance[employeeCode]) {
                if (activity === 'Visit') {
                    employeePerformance[employeeCode]['Visit']++;
                } else if (activity === 'Call') {
                    employeePerformance[employeeCode]['Call']++;
                } else if (activity === 'Reference') {
                    employeePerformance[employeeCode]['Reference']++;
                }

                if (typeOfCustomer === 'New Customer Lead') {
                    employeePerformance[employeeCode]['New Customer Leads']++;
                }
            }
        });

        let tableHtml = '<table class="report-table"><thead><tr>';
        tableHtml += '<th>Employee Name</th><th>Designation</th><th>Visits</th><th>Calls</th><th>References</th><th>New Customer Leads</th><th>Visit Target</th><th>Call Target</th><th>Reference Target</th><th>New Customer Leads Target</th><th>Visit Ach (%)</th><th>Call Ach (%)</th><th>Reference Ach (%)</th><th>New Customer Leads Ach (%)</th>';
        tableHtml += '</tr></thead><tbody>';

        // Sort employees by name
        const sortedEmployees = Object.values(employeePerformance).sort((a, b) => a.name.localeCompare(b.name));

        sortedEmployees.forEach(emp => {
            const visitAch = emp.targets['Visit'] > 0 ? ((emp['Visit'] / emp.targets['Visit']) * 100).toFixed(1) : (emp['Visit'] > 0 ? '100+' : '0.0');
            const callAch = emp.targets['Call'] > 0 ? ((emp['Call'] / emp.targets['Call']) * 100).toFixed(1) : (emp['Call'] > 0 ? '100+' : '0.0');
            const refAch = emp.targets['Reference'] > 0 ? ((emp['Reference'] / emp.targets['Reference']) * 100).toFixed(1) : (emp['Reference'] > 0 ? '100+' : '0.0');
            const newLeadAch = emp.targets['New Customer Leads'] > 0 ? ((emp['New Customer Leads'] / emp.targets['New Customer Leads']) * 100).toFixed(1) : (emp['New Customer Leads'] > 0 ? '100+' : '0.0');


            tableHtml += `<tr>
                <td data-label="Employee Name">${emp.name}</td>
                <td data-label="Designation">${emp.designation}</td>
                <td data-label="Visits">${emp['Visit']}</td>
                <td data-label="Calls">${emp['Call']}</td>
                <td data-label="References">${emp['Reference']}</td>
                <td data-label="New Customer Leads">${emp['New Customer Leads']}</td>
                <td data-label="Visit Target">${emp.targets['Visit']}</td>
                <td data-label="Call Target">${emp.targets['Call']}</td>
                <td data-label="Reference Target">${emp.targets['Reference']}</td>
                <td data-label="New Customer Leads Target">${emp.targets['New Customer Leads']}</td>
                <td data-label="Visit Ach (%)">${visitAch}%</td>
                <td data-label="Call Ach (%)">${callAch}%</td>
                <td data-label="Reference Ach (%)">${refAch}%</td>
                <td data-label="New Customer Leads Ach (%)">${newLeadAch}%</td>
            </tr>`;
        });

        tableHtml += '</tbody></table>';
        reportDisplay.innerHTML += tableHtml;
    }


    function generateEmployeeSummaryReport() {
        const selectedEmployeeCode = employeeSelect.value;
        const selectedEmployeeName = employeeCodeToNameMap[selectedEmployeeCode];
        const selectedEmployeeDesignation = employeeCodeToDesignationMap[selectedEmployeeCode];

        if (!selectedEmployeeCode) {
            reportDisplay.innerHTML = '<p>Please select an employee to view their summary report.</p>';
            return;
        }

        reportDisplay.innerHTML = `<h2>Summary Report for ${selectedEmployeeName} (${selectedEmployeeCode})</h2>`;
        reportDisplay.innerHTML += `<p><strong>Designation:</strong> ${selectedEmployeeDesignation}</p>`;

        const employeeActivities = {
            'Visit': 0,
            'Call': 0,
            'Reference': 0,
            'New Customer Leads': 0
        };

        selectedEmployeeCodeEntries.forEach(entry => {
            const activity = entry[HEADER_ACTIVITY_TYPE];
            const typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER];

            if (employeeActivities[activity] !== undefined) {
                employeeActivities[activity]++;
            }
            if (typeOfCustomer === 'New Customer Lead') {
                employeeActivities['New Customer Leads']++;
            }
        });

        const targets = TARGETS[selectedEmployeeDesignation] || TARGETS['Default'];

        let summaryHtml = '<div class="summary-breakdown-card">';
        for (const activityType in employeeActivities) {
            const achieved = employeeActivities[activityType];
            const target = targets[activityType] || 0;
            const percentage = target > 0 ? ((achieved / target) * 100).toFixed(1) : (achieved > 0 ? '100+' : '0.0'); // Handle division by zero

            summaryHtml += `
                <div class="summary-item">
                    <h3>${activityType}</h3>
                    <p>Achieved: <span class="achieved-value">${achieved}</span></p>
                    <p>Target: <span class="target-value">${target}</span></p>
                    <p>Achievement: <span class="percentage-value ${parseFloat(percentage) >= 100 ? 'positive' : 'negative'}">${percentage}%</span></p>
                </div>
            `;
        }
        summaryHtml += '</div>';
        reportDisplay.innerHTML += summaryHtml;
    }

    function generateAllEntriesReport() {
        const selectedBranch = branchSelect.value;
        const selectedEmployeeCode = employeeSelect.value;

        reportDisplay.innerHTML = '<h2>All Entries</h2>';

        let filteredEntries = allCanvassingData;

        if (selectedBranch) {
            filteredEntries = filteredEntries.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch);
            reportDisplay.innerHTML = `<h2>All Entries for ${selectedBranch}</h2>`;
        }

        if (selectedEmployeeCode) {
            filteredEntries = filteredEntries.filter(entry => entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode);
            reportDisplay.innerHTML = `<h2>All Entries for ${employeeCodeToNameMap[selectedEmployeeCode]} (${selectedEmployeeCode}) in ${selectedBranch}</h2>`;
        }

        if (filteredEntries.length === 0) {
            reportDisplay.innerHTML += '<p>No entries found for the selected criteria.</p>';
            return;
        }

        let tableHtml = '<table class="report-table all-entries-table"><thead><tr>';
        // Dynamically create headers from the first entry's keys,
        // or a predefined list if specific order is needed.
        const displayHeaders = [
            HEADER_TIMESTAMP,
            HEADER_DATE,
            HEADER_BRANCH_NAME,
            HEADER_EMPLOYEE_NAME,
            HEADER_EMPLOYEE_CODE,
            HEADER_DESIGNATION,
            HEADER_ACTIVITY_TYPE,
            HEADER_TYPE_OF_CUSTOMER,
            HEADER_R_LEAD_SOURCE,
            HEADER_HOW_CONTACTED,
            HEADER_PROSPECT_NAME,
            HEADER_PHONE_NUMBER_WHATSAPP,
            HEADER_ADDRESS,
            HEADER_PROFESSION,
            HEADER_DOB_WD,
            HEADER_PRODUCT_INTERESTED,
            HEADER_REMARKS,
            HEADER_NEXT_FOLLOW_UP_DATE,
            HEADER_RELATION_WITH_STAFF,
            HEADER_FAMILY_DETAILS_1, // NEW
            HEADER_FAMILY_DETAILS_2, // NEW
            HEADER_FAMILY_DETAILS_3, // NEW
            HEADER_FAMILY_DETAILS_4, // NEW
            HEADER_PROFILE_OF_CUSTOMER // NEW
        ];

        displayHeaders.forEach(header => {
            tableHtml += `<th>${header}</th>`;
        });
        tableHtml += '</tr></thead><tbody>';

        filteredEntries.forEach(entry => {
            tableHtml += '<tr>';
            displayHeaders.forEach(header => {
                let cellValue = entry[header] || ''; // Use empty string for undefined/null
                if (header === HEADER_DATE || header === HEADER_NEXT_FOLLOW_UP_DATE) {
                    cellValue = formatDate(cellValue);
                }
                tableHtml += `<td data-label="${header}">${cellValue}</td>`;
            });
            tableHtml += '</tr>';
        });

        tableHtml += '</tbody></table>';
        reportDisplay.innerHTML += tableHtml;
    }

    function generateAllStaffOverallPerformanceReport() {
        reportDisplay.innerHTML = '<h2>All Staff Overall Performance Report</h2>';

        const employeeOverallPerformance = {};

        // Initialize all predefined employees, including those with no activities
        PREDEFINED_EMPLOYEES.forEach(emp => {
            employeeOverallPerformance[emp.employeeCode] = {
                name: emp.employeeName,
                branch: emp.branchName,
                designation: emp.designation,
                'Visit': 0,
                'Call': 0,
                'Reference': 0,
                'New Customer Leads': 0,
                targets: TARGETS[emp.designation] || TARGETS['Default']
            };
        });

        // Populate actual activity counts from canvassing data
        allCanvassingData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const activity = entry[HEADER_ACTIVITY_TYPE];
            const typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER];

            if (employeeOverallPerformance[employeeCode]) {
                if (activity === 'Visit') {
                    employeeOverallPerformance[employeeCode]['Visit']++;
                } else if (activity === 'Call') {
                    employeeOverallPerformance[employeeCode]['Call']++;
                } else if (activity === 'Reference') {
                    employeeOverallPerformance[employeeCode]['Reference']++;
                }

                if (typeOfCustomer === 'New Customer Lead') {
                    employeeOverallPerformance[employeeCode]['New Customer Leads']++;
                }
            }
        });

        let tableHtml = '<table class="report-table"><thead><tr>';
        tableHtml += '<th>Employee Name</th><th>Employee Code</th><th>Branch</th><th>Designation</th><th>Visits</th><th>Calls</th><th>References</th><th>New Customer Leads</th><th>Visit Target</th><th>Call Target</th><th>Reference Target</th><th>New Customer Leads Target</th><th>Visit Ach (%)</th><th>Call Ach (%)</th><th>Reference Ach (%)</th><th>New Customer Leads Ach (%)</th>';
        tableHtml += '</tr></thead><tbody>';

        const sortedEmployees = Object.values(employeeOverallPerformance).sort((a, b) => a.name.localeCompare(b.name));

        sortedEmployees.forEach(emp => {
            const visitAch = emp.targets['Visit'] > 0 ? ((emp['Visit'] / emp.targets['Visit']) * 100).toFixed(1) : (emp['Visit'] > 0 ? '100+' : '0.0');
            const callAch = emp.targets['Call'] > 0 ? ((emp['Call'] / emp.targets['Call']) * 100).toFixed(1) : (emp['Call'] > 0 ? '100+' : '0.0');
            const refAch = emp.targets['Reference'] > 0 ? ((emp['Reference'] / emp.targets['Reference']) * 100).toFixed(1) : (emp['Reference'] > 0 ? '100+' : '0.0');
            const newLeadAch = emp.targets['New Customer Leads'] > 0 ? ((emp['New Customer Leads'] / emp.targets['New Customer Leads']) * 100).toFixed(1) : (emp['New Customer Leads'] > 0 ? '100+' : '0.0');

            tableHtml += `<tr>
                <td data-label="Employee Name">${emp.name}</td>
                <td data-label="Employee Code">${emp.employeeCode}</td>
                <td data-label="Branch">${emp.branch}</td>
                <td data-label="Designation">${emp.designation}</td>
                <td data-label="Visits">${emp['Visit']}</td>
                <td data-label="Calls">${emp['Call']}</td>
                <td data-label="References">${emp['Reference']}</td>
                <td data-label="New Customer Leads">${emp['New Customer Leads']}</td>
                <td data-label="Visit Target">${emp.targets['Visit']}</td>
                <td data-label="Call Target">${emp.targets['Call']}</td>
                <td data-label="Reference Target">${emp.targets['Reference']}</td>
                <td data-label="New Customer Leads Target">${emp.targets['New Customer Leads']}</td>
                <td data-label="Visit Ach (%)">${visitAch}%</td>
                <td data-label="Call Ach (%)">${callAch}%</td>
                <td data-label="Reference Ach (%)">${refAch}%</td>
                <td data-label="New Customer Leads Ach (%)">${newLeadAch}%</td>
            </tr>`;
        });

        tableHtml += '</tbody></table>';
        reportDisplay.innerHTML += tableHtml;
    }


    function generateNonParticipatingBranchesReport() {
        reportDisplay.innerHTML = '<h2>Non-Participating Branches</h2>';

        const participatingBranches = new Set();
        allCanvassingData.forEach(entry => {
            participatingBranches.add(entry[HEADER_BRANCH_NAME]);
        });

        // Use allUniqueBranches (dynamically generated) for comparison
        const nonParticipatingBranches = allUniqueBranches.filter(
            branch => !participatingBranches.has(branch)
        );

        if (nonParticipatingBranches.length > 0) {
            let listHtml = '<ul>';
            nonParticipatingBranches.forEach(branch => {
                listHtml += `<li>${branch}</li>`;
            });
            listHtml += '</ul>';
            reportDisplay.innerHTML += listHtml;
        } else {
            reportDisplay.innerHTML += '<p class="no-participation-message">All predefined branches have participated!</p>';
        }
    }

    function generateNonParticipatingEmployeesReport() {
        reportDisplay.innerHTML = '<h2>Non-Participating Employees</h2>';

        const participatingEmployeeCodes = new Set();
        allCanvassingData.forEach(entry => {
            participatingEmployeeCodes.add(entry[HEADER_EMPLOYEE_CODE]);
        });

        const nonParticipatingEmployees = PREDEFINED_EMPLOYEES.filter(
            employee => !participatingEmployeeCodes.has(employee.employeeCode)
        );

        if (nonParticipatingEmployees.length > 0) {
            let tableHtml = '<table class="report-table"><thead><tr><th>Employee Name</th><th>Employee Code</th><th>Branch</th><th>Designation</th></tr></thead><tbody>';
            nonParticipatingEmployees.forEach(emp => {
                tableHtml += `<tr>
                    <td data-label="Employee Name">${emp.employeeName}</td>
                    <td data-label="Employee Code">${emp.employeeCode}</td>
                    <td data-label="Branch">${emp.branchName}</td>
                    <td data-label="Designation">${emp.designation}</td>
                </tr>`;
            });
            tableHtml += '</tbody></table>';
            reportDisplay.innerHTML += tableHtml;
        } else {
            reportDisplay.innerHTML += '<p class="no-participation-message">All predefined employees have participated!</p>';
        }
    }

    // NEW: Detailed Customer View Functions
    function populateCustomerCanvassedList() {
        const selectedBranch = customerViewBranchSelect.value;
        const selectedEmployeeCode = customerViewEmployeeSelect.value;

        let filteredCustomers = allCanvassingData;

        if (selectedBranch) {
            filteredCustomers = filteredCustomers.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch);
        }
        if (selectedEmployeeCode) {
            filteredCustomers = filteredCustomers.filter(entry => entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode);
        }

        // Filter for unique prospect names/phone numbers to avoid duplicates in the list
        // Assuming Phone Number (Whatsapp) is a good unique identifier
        const uniqueCustomers = new Map();
        filteredCustomers.forEach(entry => {
            const phoneNumber = entry[HEADER_PHONE_NUMBER_WHATSAPP];
            if (phoneNumber && !uniqueCustomers.has(phoneNumber)) {
                uniqueCustomers.set(phoneNumber, entry);
            }
        });


        customerCanvassedList.innerHTML = ''; // Clear previous list
        customerDetailsContent.innerHTML = '<p>Select a customer from the list to view details.</p>'; // Clear details

        if (uniqueCustomers.size === 0) {
            customerCanvassedList.innerHTML = '<li>No customers found for the selected filters.</li>';
            return;
        }

        uniqueCustomers.forEach(customer => {
            const listItem = document.createElement('li');
            listItem.textContent = `${customer[HEADER_PROSPECT_NAME] || 'Unknown'} - ${customer[HEADER_PHONE_NUMBER_WHATSAPP] || 'N/A'}`;
            listItem.classList.add('customer-list-item');
            listItem.dataset.phoneNumber = customer[HEADER_PHONE_NUMBER_WHATSAPP]; // Store phone number for lookup
            listItem.addEventListener('click', () => displayCustomerDetails(customer));
            customerCanvassedList.appendChild(listItem);
        });
    }

    function displayCustomerDetails(customerData) {
        customerDetailsContent.innerHTML = ''; // Clear previous details

        // Remove active class from all list items
        document.querySelectorAll('.customer-list-item').forEach(item => {
            item.classList.remove('active');
        });

        // Add active class to the clicked item
        const clickedItem = document.querySelector(`.customer-list-item[data-phone-number="${customerData[HEADER_PHONE_NUMBER_WHATSAPP]}"]`);
        if (clickedItem) {
            clickedItem.classList.add('active');
        }

        const detailsHtml = `
            <h3>Customer Details: ${customerData[HEADER_PROSPECT_NAME] || 'N/A'}</h3>
            <div class="detail-grid">
                <div class="detail-row"><span class="detail-label">Prospect Name:</span><span class="detail-value">${customerData[HEADER_PROSPECT_NAME] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Phone Number:</span><span class="detail-value">${customerData[HEADER_PHONE_NUMBER_WHATSAPP] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Address:</span><span class="detail-value">${customerData[HEADER_ADDRESS] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Profession:</span><span class="detail-value">${customerData[HEADER_PROFESSION] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">DOB/WD:</span><span class="detail-value">${customerData[HEADER_DOB_WD] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Product Interested:</span><span class="detail-value">${customerData[HEADER_PRODUCT_INTERESTED] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Relation With Staff:</span><span class="detail-value">${customerData[HEADER_RELATION_WITH_STAFF] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Family Details (Wife/Husband Name):</span><span class="detail-value">${customerData[HEADER_FAMILY_DETAILS_1] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Family Details (Wife/Husband Job):</span><span class="detail-value">${customerData[HEADER_FAMILY_DETAILS_2] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Family Details (Children Names):</span><span class="detail-value">${customerData[HEADER_FAMILY_DETAILS_3] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Family Details (Children Details):</span><span class="detail-value">${customerData[HEADER_FAMILY_DETAILS_4] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Profile of Customer:</span><span class="detail-value">${customerData[HEADER_PROFILE_OF_CUSTOMER] || 'N/A'}</span></div>
            </div>
            <h4>Canvassing History:</h4>
            <ul class="canvassing-history-list">
        `;

        // Get all entries for this specific customer (based on phone number)
        const customerHistory = allCanvassingData.filter(entry =>
            entry[HEADER_PHONE_NUMBER_WHATSAPP] === customerData[HEADER_PHONE_NUMBER_WHATSAPP]
        ).sort((a, b) => new Date(b[HEADER_TIMESTAMP]) - new Date(a[HEADER_TIMESTAMP])); // Sort by newest first

        if (customerHistory.length > 0) {
            customerHistory.forEach(entry => {
                const employeeName = employeeCodeToNameMap[entry[HEADER_EMPLOYEE_CODE]] || entry[HEADER_EMPLOYEE_NAME] || 'Unknown Employee';
                detailsHtml += `
                    <li>
                        <strong>Date:</strong> ${formatDate(entry[HEADER_DATE]) || 'N/A'} <br>
                        <strong>Employee:</strong> ${employeeName} (${entry[HEADER_EMPLOYEE_CODE] || 'N/A'})<br>
                        <strong>Activity:</strong> ${entry[HEADER_ACTIVITY_TYPE] || 'N/A'}<br>
                        <strong>Type of Customer:</strong> ${entry[HEADER_TYPE_OF_CUSTOMER] || 'N/A'}<br>
                        <strong>Remarks:</strong> ${entry[HEADER_REMARKS] || 'N/A'}<br>
                        <strong>Next Follow-up:</strong> ${formatDate(entry[HEADER_NEXT_FOLLOW_UP_DATE]) || 'N/A'}
                    </li>
                `;
            });
        } else {
            detailsHtml += '<li>No canvassing history found for this customer.</li>';
        }
        detailsHtml += '</ul>';

        customerDetailsContent.innerHTML = detailsHtml;
    }


    // Function to send data to Google Apps Script Web App
    async function sendDataToGoogleAppsScript(action, data) {
        displayEmployeeManagementMessage('Sending data to server...', false);
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Important for cross-origin requests
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ action: action, data: data }),
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || `HTTP error! Status: ${response.status}`);
            }

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(result.message || 'Operation successful!', false);
                await processData(); // Re-fetch data and refresh reports after successful operation
                return true;
            } else {
                displayEmployeeManagementMessage(result.message || 'Operation failed!', true);
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Apps Script:', error);
            displayEmployeeManagementMessage(`Error: ${error.message}`, true);
            return false;
        }
    }


    // Main function to orchestrate data fetching and initial report generation
    async function processData() {
        displayMessage('Loading dashboard data...', 'info');
        await fetchCanvassingData();
        initializeFilters(); // Initialize dropdowns with predefined data

        // Set initial filter to show all entries
        if (branchSelect) branchSelect.value = '';
        if (employeeSelect) employeeSelect.value = '';
        filterDataForReports(); // Apply initial filters to setup report buttons
        generateAllBranchSnapshotReport(); // Always show initial report
        displayMessage('Dashboard data loaded.', 'success');
    }

    // --- Event Listeners ---

    // Tab Navigation
    if (allBranchSnapshotTabBtn) {
        allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    }
    if (allStaffOverallPerformanceTabBtn) {
        allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    }
    if (nonParticipatingBranchesTabBtn) {
        nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    }
    if (nonParticipatingEmployeesTabBtn) { // NEW
        nonParticipatingEmployeesTabBtn.addEventListener('click', () => showTab('nonParticipatingEmployeesTabBtn'));
    }
    if (detailedCustomerViewTabBtn) { // NEW
        detailedCustomerViewTabBtn.addEventListener('click', () => showTab('detailedCustomerViewTabBtn'));
    }
    if (employeeManagementTabBtn) {
        employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    }


    function showTab(tabId) {
        // Remove 'active' class from all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Add 'active' class to the clicked tab button
        document.getElementById(tabId).classList.add('active');

        // Hide all main sections
        if (reportsSection) reportsSection.style.display = 'none';
        if (detailedCustomerViewSection) detailedCustomerViewSection.style.display = 'none'; // NEW
        if (employeeManagementSection) employeeManagementSection.style.display = 'none';

        // Show the relevant section and generate report
        if (tabId === 'allBranchSnapshotTabBtn') {
            if (reportsSection) reportsSection.style.display = 'block';
            filterDataForReports(); // Re-apply filters for button visibility
            generateAllBranchSnapshotReport();
            // Reset dropdowns for this view
            if (branchSelect) branchSelect.value = '';
            if (employeeSelect) employeeSelect.value = '';
            if (employeeFilterPanel) employeeFilterPanel.style.display = 'none';
            if (viewOptions) viewOptions.style.display = 'none'; // Hide view options
        } else if (tabId === 'allStaffOverallPerformanceTabBtn') {
            if (reportsSection) reportsSection.style.display = 'block';
            generateAllStaffOverallPerformanceReport();
            if (viewOptions) viewOptions.style.display = 'none'; // Hide view options
            // Reset dropdowns for this view
            if (branchSelect) branchSelect.value = '';
            if (employeeSelect) employeeSelect.value = '';
            if (employeeFilterPanel) employeeFilterPanel.style.display = 'none';
        } else if (tabId === 'nonParticipatingBranchesTabBtn') {
            if (reportsSection) reportsSection.style.display = 'block';
            generateNonParticipatingBranchesReport();
            if (viewOptions) viewOptions.style.display = 'none'; // Hide view options
            // Reset dropdowns for this view
            if (branchSelect) branchSelect.value = '';
            if (employeeSelect) employeeSelect.value = '';
            if (employeeFilterPanel) employeeFilterPanel.style.display = 'none';
        } else if (tabId === 'nonParticipatingEmployeesTabBtn') { // NEW
            if (reportsSection) reportsSection.style.display = 'block';
            generateNonParticipatingEmployeesReport();
            if (viewOptions) viewOptions.style.display = 'none';
            if (branchSelect) branchSelect.value = '';
            if (employeeSelect) employeeSelect.value = '';
            if (employeeFilterPanel) employeeFilterPanel.style.display = 'none';
        } else if (tabId === 'detailedCustomerViewTabBtn') { // NEW
            if (detailedCustomerViewSection) detailedCustomerViewSection.style.display = 'block';
            // Ensure customer view dropdowns are populated
            if (customerViewBranchSelect) customerViewBranchSelect.value = '';
            if (customerViewEmployeeSelect) customerViewEmployeeSelect.value = '';
            populateCustomerCanvassedList();
        } else if (tabId === 'employeeManagementTabBtn') {
            if (employeeManagementSection) employeeManagementSection.style.display = 'block';
        }
    }


    if (branchSelect) {
        branchSelect.addEventListener('change', () => {
            filterDataForReports();
            if (branchSelect.value) {
                // If a branch is selected, show branch and employee filter panels
                employeeFilterPanel.style.display = 'block';
                viewOptions.style.display = 'block'; // Show general view options
                // Populate employee dropdown based on selected branch
                const filteredEmployees = PREDEFINED_EMPLOYEES.filter(emp => emp.branchName === branchSelect.value);
                employeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>'; // Clear existing
                filteredEmployees.sort((a, b) => a.employeeName.localeCompare(b.employeeName));
                filteredEmployees.forEach(emp => {
                    const option = document.createElement('option');
                    option.value = emp.employeeCode;
                    option.textContent = `${emp.employeeName} (${emp.employeeCode})`;
                    employeeSelect.appendChild(option);
                });
                employeeSelect.value = ''; // Reset employee selection
                generateBranchPerformanceReport(); // Default report for branch
            } else {
                // If no branch selected, hide employee filter and show all branch snapshot
                employeeFilterPanel.style.display = 'none';
                viewOptions.style.display = 'none';
                generateAllBranchSnapshotReport();
            }
        });
    }

    if (employeeSelect) {
        employeeSelect.addEventListener('change', () => {
            filterDataForReports();
            if (employeeSelect.value) {
                viewOptions.style.display = 'block'; // Ensure view options are visible
                generateEmployeeSummaryReport(); // Default report for employee
            } else {
                // If no employee selected, show branch performance if branch is selected, else all branch snapshot
                if (branchSelect.value) {
                    generateBranchPerformanceReport();
                } else {
                    generateAllBranchSnapshotReport();
                    viewOptions.style.display = 'none'; // Hide if neither branch nor employee is selected
                }
            }
        });
    }

    if (viewBranchPerformanceReportBtn) {
        viewBranchPerformanceReportBtn.addEventListener('click', generateBranchPerformanceReport);
    }
    if (viewEmployeeSummaryBtn) {
        viewEmployeeSummaryBtn.addEventListener('click', generateEmployeeSummaryReport);
    }
    if (viewAllEntriesBtn) {
        viewAllEntriesBtn.addEventListener('click', generateAllEntriesReport);
    }

    // NEW: Customer View Filter Event Listeners
    if (customerViewBranchSelect) {
        customerViewBranchSelect.addEventListener('change', () => {
            populateCustomerCanvassedList();
            const selectedBranch = customerViewBranchSelect.value;
            // Populate employee dropdown based on selected branch
            const filteredEmployees = PREDEFINED_EMPLOYEES.filter(emp => emp.branchName === selectedBranch);
            customerViewEmployeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>'; // Clear existing
            filteredEmployees.sort((a, b) => a.employeeName.localeCompare(b.employeeName));
            filteredEmployees.forEach(emp => {
                const option = document.createElement('option');
                option.value = emp.employeeCode;
                option.textContent = `${emp.employeeName} (${emp.employeeCode})`;
                customerViewEmployeeSelect.appendChild(option);
            });
            customerViewEmployeeSelect.value = ''; // Reset employee selection
        });
    }
    if (customerViewEmployeeSelect) {
        customerViewEmployeeSelect.addEventListener('change', populateCustomerCanvassedList);
    }


    // Employee Management Forms Event Listeners
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const newEmployeeData = {
                [HEADER_EMPLOYEE_NAME]: newEmployeeNameInput.value.trim(),
                [HEADER_EMPLOYEE_CODE]: newEmployeeCodeInput.value.trim(),
                [HEADER_BRANCH_NAME]: newBranchNameInput.value.trim(),
                [HEADER_DESIGNATION]: newDesignationInput.value.trim()
            };

            if (!newEmployeeData[HEADER_EMPLOYEE_NAME] || !newEmployeeData[HEADER_EMPLOYEE_CODE] || !newEmployeeData[HEADER_BRANCH_NAME] || !newEmployeeData[HEADER_DESIGNATION]) {
                displayEmployeeManagementMessage('All fields are required for adding an employee.', true);
                return;
            }

            // Check if employee code already exists in PREDEFINED_EMPLOYEES
            const employeeExists = PREDEFINED_EMPLOYEES.some(emp => emp.employeeCode === newEmployeeData[HEADER_EMPLOYEE_CODE]);
            if (employeeExists) {
                displayEmployeeManagementMessage(`Employee with code ${newEmployeeData[HEADER_EMPLOYEE_CODE]} already exists. Please use a unique code.`, true);
                return;
            }

            const success = await sendDataToGoogleAppsScript('add_employee', newEmployeeData);
            if (success) {
                addEmployeeForm.reset();
                // Add to PREDEFINED_EMPLOYEES and re-initialize filters locally for immediate reflection
                PREDEFINED_EMPLOYEES.push(newEmployeeData);
                // Re-sort PREDEFINED_EMPLOYEES after adding new employee
                PREDEFINED_EMPLOYEES.sort((a, b) => a.employeeName.localeCompare(b.employeeName));
                initializeFilters(); // Re-populate dropdowns
            }
        });
    }

    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const detailsText = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName || !detailsText) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk addition.', true);
                return;
            }

            const lines = detailsText.split('\n');
            const employeesToAdd = [];
            const existingCodes = new Set(PREDEFINED_EMPLOYEES.map(emp => emp.employeeCode));
            const codesInCurrentBulk = new Set();

            for (const line of lines) {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length >= 2) { // Expecting at least Name and Code
                    const employeeCode = parts[1];
                    if (existingCodes.has(employeeCode) || codesInCurrentBulk.has(employeeCode)) {
                        displayEmployeeManagementMessage(`Skipping duplicate employee code in bulk upload: ${employeeCode}`, true);
                        continue;
                    }
                    codesInCurrentBulk.add(employeeCode);
                    const employeeData = {
                        [HEADER_EMPLOYEE_NAME]: parts[0],
                        [HEADER_EMPLOYEE_CODE]: parts[1],
                        [HEADER_BRANCH_NAME]: branchName,
                        [HEADER_DESIGNATION]: parts[2] || '' // Designation is optional
                    };
                    employeesToAdd.push(employeeData);
                }
            }

            if (employeesToAdd.length > 0) {
                const success = await sendDataToGoogleAppsScript('add_bulk_employees', employeesToAdd);
                if (success) {
                    bulkAddEmployeeForm.reset();
                    // Update PREDEFINED_EMPLOYEES and re-initialize filters locally
                    PREDEFINED_EMPLOYEES.push(...employeesToAdd);
                    // Re-sort PREDEFINED_EMPLOYEES after adding new employees
                    PREDEFINED_EMPLOYEES.sort((a, b) => a.employeeName.localeCompare(b.employeeName));
                    initializeFilters();
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

            // Check if employee exists in PREDEFINED_EMPLOYEES
            const employeeIndex = PREDEFINED_EMPLOYEES.findIndex(emp => emp.employeeCode === employeeCodeToDelete);
            if (employeeIndex === -1) {
                displayEmployeeManagementMessage(`Employee with code ${employeeCodeToDelete} not found.`, true);
                return;
            }

            const deleteData = { [HEADER_EMPLOYEE_CODE]: employeeCodeToDelete };
            const success = await sendDataToGoogleAppsScript('delete_employee', deleteData);

            if (success) {
                deleteEmployeeForm.reset();
                // Remove from PREDEFINED_EMPLOYEES and re-initialize filters locally
                PREDEFINED_EMPLOYEES.splice(employeeIndex, 1);
                initializeFilters(); // Re-populate dropdowns
            }
        });
    }

    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn');
});
