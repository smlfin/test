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
    // Predefined list of branches for the dropdown and "no participation" check
    const PREDEFINED_BRANCHES = [
        "Angamaly", "Corporate Office", "Edappally", "Harippad", "Koduvayur", "Kuzhalmannam",
        "Mattanchery", "Mavelikara", "Nedumkandom", "Nenmara", "Paravoor", "Perumbavoor",
        "Thiruwillamala", "Thodupuzha", "Chengannur", "Alathur", "Kottayam", "Kattapana",
        "Muvattupuzha", "Thiruvalla", "Pathanamthitta", "HO KKM","Kunnamkulam" // Corrected "Pathanamthitta" typo if it existed previously
    ].sort();

    const PREDEFINED_EMPLOYEES = [
        { employeeCode: "4256", employeeName: "Sareena M", branchName: "Alathur", designation: "BRANCH MANAGER" },
        { employeeCode: "4498", employeeName: "SAGEESH KUMAR S", branchName: "Alathur", designation: "LOAN OFFICER" },
        { employeeCode: "4560", employeeName: "DHANYA K", branchName: "Alathur", designation: "LOAN EXECUTIVE" },
        { employeeCode: "4561", employeeName: "MANJULA B", branchName: "Alathur", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "4528", employeeName: "ELIZABETH JANCY P M", branchName: "Angamaly", designation: "SENIOR BRANCH MANAGER" },
        { employeeCode: "4416", employeeName: "Renju Unnikrishnan", branchName: "Chengannur", designation: "BRANCH MANAGER" },
        { employeeCode: "4431", employeeName: "VEENA KRISHNAN", branchName: "Chengannur", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "4450", employeeName: "Renjini G", branchName: "Chengannur", designation: "LOAN EXECUTIVE" },
        { employeeCode: "1292", employeeName: "ANOOP M A", branchName: "Corporate Office", designation: "ASSISTANT MANAGER" },
        { employeeCode: "1414", employeeName: "ASHOK KUMAR K G", branchName: "Corporate Office", designation: "Sr.CREDIT MANAGER" },
        { employeeCode: "1098", employeeName: "EBY SCARIA", branchName: "Corporate Office", designation: "MANAGER - POST VERIFICATION" },
        { employeeCode: "1245", employeeName: "KUMARI K N", branchName: "Corporate Office", designation: "Sr.EXECUTIVE" },
        { employeeCode: "1141", employeeName: "LIJOY V R", branchName: "Corporate Office", designation: "ASSISTANT MANAGER (MIS & OPERATIONS)" },
        { employeeCode: "1564", employeeName: "MARY LITTLE FLOWER", branchName: "Corporate Office", designation: "Legal Officer" },
        { employeeCode: "1057", employeeName: "PROMOD M", branchName: "Corporate Office", designation: "MANAGER" },
        { employeeCode: "1409", employeeName: "SHAIJO JOSE", branchName: "Corporate Office", designation: "OFFICER" },
        { employeeCode: "1049", employeeName: "SIVARAJ KUMAR K P", branchName: "Corporate Office", designation: "SENIOR MANAGER" },
        { employeeCode: "1125", employeeName: "NARAYANAN K", branchName: "Corporate Office", designation: "ASSOCIATE VICE PRESIDENT" },
        { employeeCode: "1064", employeeName: "SARITHA K T", branchName: "Corporate Office", designation: "Asst.Manager" },
        { employeeCode: "1078", employeeName: "CHITHRA AJITH", branchName: "Corporate Office", designation: "Sr.EXECUTIVE" },
        { employeeCode: "1411", employeeName: "ANOOP ANANTHAN NAIR", branchName: "Corporate Office", designation: "OFFICER" },
        { employeeCode: "1607", employeeName: "SHIJO PAUL", branchName: "Corporate Office", designation: "DRIVER" },
        { employeeCode: "1441", employeeName: "BINSY JENSON", branchName: "Corporate Office", designation: "Sr.EXECUTIVE" },
        { employeeCode: "1452", employeeName: "JESSY ROY", branchName: "Corporate Office", designation: "MANAGER" },
        { employeeCode: "1918", employeeName: "GOWRI R", branchName: "Corporate Office", designation: "Sr.AVP" }, // Combined from sources
        { employeeCode: "3121", employeeName: "NIMISHA RAJEEV", branchName: "Corporate Office", designation: "Legal Officer" },
        { employeeCode: "3948", employeeName: "SINI", branchName: "Corporate Office", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "4102", employeeName: "ALVIN THOMAS", branchName: "Corporate Office", designation: "SENIOR EXECUTIVE" },
        { employeeCode: "4104", employeeName: "MANITHA MOHANAN", branchName: "Corporate Office", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "4110", employeeName: "SHILPA JIBIN", branchName: "Corporate Office", designation: "MIS COORDINATOR" },
        { employeeCode: "4128", employeeName: "SHIBEENA V M", branchName: "Corporate Office", designation: "LEGAL MANAGER" },
        { employeeCode: "4235", employeeName: "Vijesh K V", branchName: "Corporate Office", designation: "MANAGER" },
        { employeeCode: "4242", employeeName: "Niveeth P V", branchName: "Corporate Office", designation: "IT SUPPORT" },
        { employeeCode: "4290", employeeName: "Dipin C G", branchName: "Corporate Office", designation: "DRIVER" },
        { employeeCode: "4309", employeeName: "Tony George", branchName: "Corporate Office", designation: "MAINTENANCE EXECUTIVE" },
        { employeeCode: "4325", employeeName: "Benedict Anto S M", branchName: "Corporate Office", designation: "MANAGER" },
        { employeeCode: "4333", employeeName: "Dhanya L", branchName: "Corporate Office", designation: "Accounts manager" },
        { employeeCode: "4378", employeeName: "SOBHI BABU", branchName: "Corporate Office", designation: "HOUSE KEEPER" },
        { employeeCode: "4389", employeeName: "GEETHU GOPAN", branchName: "Corporate Office", designation: "EXECUTIVE" },
        { employeeCode: "4412", employeeName: "Fathima Ripsana A M", branchName: "Corporate Office", designation: "Accounts Executive" },
        { employeeCode: "4411", employeeName: "ANNIE JOSE", branchName: "Corporate Office", designation: "Accounts Officer" },
        { employeeCode: "4394", employeeName: "Athul Augustine", branchName: "Corporate Office", designation: "DRIVER" },
        { employeeCode: "4421", employeeName: "Aswathy Sasi", branchName: "Corporate Office", designation: "MIS OFFICER" },
        { employeeCode: "4426", employeeName: "FABIN ANTONY", branchName: "Corporate Office", designation: "ASSISTANT MANAGER" },
        { employeeCode: "4434", employeeName: "Sanjay Sajan", branchName: "Corporate Office", designation: "EXECUTIVE" },
        { employeeCode: "4441", employeeName: "Manoj N Das", branchName: "Corporate Office", designation: "Chief Manager" },
        { employeeCode: "4448", employeeName: "A AJMAL", branchName: "Corporate Office", designation: "ASSISTANT MANAGER" },
        { employeeCode: "4444", employeeName: "Ranjith E N", branchName: "Corporate Office", designation: "Legal Co-ordinator" },
        { employeeCode: "4465", employeeName: "BHAGYALAKSHMI E S", branchName: "Corporate Office", designation: "ACCOUNTS EXECUTIVE" },
        { employeeCode: "4459", employeeName: "ASHIN JENSON", branchName: "Corporate Office", designation: "ACCOUNTS EXECUTIVE" },
        { employeeCode: "4468", employeeName: "SIJO M J", branchName: "Corporate Office", designation: "SENIOR EXECUTIVE" },
        { employeeCode: "4479", employeeName: "JITHU K JOSE", branchName: "Corporate Office", designation: "Sr.EXECUTIVE" },
        { employeeCode: "4495", employeeName: "BINOJDAS V H", branchName: "Corporate Office", designation: "VICE PRESIDENT" },
        { employeeCode: "4487", employeeName: "ANANTHU RAJENDRAN", branchName: "Corporate Office", designation: "DRIVER" },
        { employeeCode: "4486", employeeName: "PRAYAG VINOD", branchName: "Corporate Office", designation: "ACCOUNTS EXECUTIVE" },
        { employeeCode: "4497", employeeName: "VISHNU S SANTHOSH", branchName: "Corporate Office", designation: "ACCOUNTS EXECUTIVE" },
        { employeeCode: "4513", employeeName: "GINCE GEORGE", branchName: "Corporate Office", designation: "SENIOR MANAGER" },
        { employeeCode: "4516", employeeName: "SYAMILY M NAIR", branchName: "Corporate Office", designation: "ASSISTANT MANAGER" },
        { employeeCode: "4529", employeeName: "AJI GEORGE", branchName: "Corporate Office", designation: "SENIOR EXECUTIVE" },
        { employeeCode: "4566", employeeName: "ABIN K R", branchName: "Corporate Office", designation: "RECOVERY MANAGER" },
        { employeeCode: "4557", employeeName: "AJAY PB", branchName: "Corporate Office", designation: "SENIOR EXECUTIVE" },
        { employeeCode: "4575", employeeName: "GANESH K S", branchName: "Corporate Office", designation: "Sr.VICE PRESIDENT" },
        { employeeCode: "4576", employeeName: "VASUDEVAN V", branchName: "Corporate Office", designation: "Sr.AVP (RESOURCE MOBILIZATION & OPERATIONAL ADMIN)" },
        { employeeCode: "4582", employeeName: "ANAND M B", branchName: "Corporate Office", designation: "RECOVERY MANAGER" },
        { employeeCode: "4583", employeeName: "MANU MOHAN", branchName: "Corporate Office", designation: "MANAGER" },
        { employeeCode: "4632", employeeName: "MARIA SONY", branchName: "Corporate Office", designation: "HR Executive" },
        { employeeCode: "3944", employeeName: "BEENA C S", branchName: "Corporate Office", designation: "EXECUTIVE" },
        { employeeCode: "1222", employeeName: "SIVAKALA DILEEP", branchName: "Edappally", designation: "CUSTOMER SERVICE EXECUTIVE" },
        { employeeCode: "4631", employeeName: "ATHIRA K T", branchName: "Edappally", designation: "Brach Manager" },
        { employeeCode: "4630", employeeName: "SHABANAM K U", branchName: "Edappally", designation: "OFFICER" },
        { employeeCode: "4272", employeeName: "Sandhya S", branchName: "Harippad", designation: "BRANCH MANAGER" },
        { employeeCode: "4425", employeeName: "ASWATHY K", branchName: "Harippad", designation: "ABM" },
        { employeeCode: "4526", employeeName: "AMPILI P V", branchName: "Harippad", designation: "ABM" },
        { employeeCode: "4623", employeeName: "ASWIN KRISHNA K", branchName: "Harippad", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "1025", employeeName: "Tony K F", branchName: "HO KKM", designation: "Sr.AVP" },
        { employeeCode: "1024", employeeName: "Ramakrishnan", branchName: "HO KKM", designation: "Mechanic" },
        { employeeCode: "1528", employeeName: "Hridya", branchName: "HO KKM", designation: "MANAGER" },
        { employeeCode: "2275", employeeName: "Shreekumar S", branchName: "HO KKM", designation: "Sr.AVP" },
        { employeeCode: "1261", employeeName: "Priya Mohanan", branchName: "HO KKM", designation: "Receptionist" },
        { employeeCode: "1018", employeeName: "Janus Samuel", branchName: "HO KKM", designation: "Sr.Officer" },
        { employeeCode: "1063", employeeName: "Ajitha Haridas", branchName: "HO KKM", designation: "Sr.Officer" },
        { employeeCode: "1118", employeeName: "Sreeja Jaison", branchName: "HO KKM", designation: "MANAGER" },
        { employeeCode: "1555", employeeName: "Deepa Vijayan", branchName: "HO KKM", designation: "Office Executive" },
        { employeeCode: "1009", employeeName: "Shajan.A.D", branchName: "HO KKM", designation: "AVP" },
        { employeeCode: "4278", employeeName: "Bimi Micle P", branchName: "HO KKM", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "4336", employeeName: "Vineesh T", branchName: "HO KKM", designation: "Sr.AVP" },
        { employeeCode: "4337", employeeName: "Prashob P P", branchName: "HO KKM", designation: "DRIVER" },
        { employeeCode: "4595", employeeName: "KIRAN K BALAN", branchName: "HO KKM", designation: "Litigation Officer(Legal)" },
        { employeeCode: "4599", employeeName: "DILJITH P J", branchName: "HO KKM", designation: "DRIVER" },
        { employeeCode: "4614", employeeName: "JISHNU P L", branchName: "HO KKM", designation: "RECOVERY MANAGER" },
        { employeeCode: "4420", employeeName: "Lishamol K Kudakkachira", branchName: "Kattapana", designation: "BRANCH MANAGER" },
        { employeeCode: "4554", employeeName: "JINCY K RAJU", branchName: "Kattapana", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "4589", employeeName: "BENCY JOSEPH", branchName: "Kattapana", designation: "ABM" },
        { employeeCode: "4615", employeeName: "SETHU RAJ", branchName: "Kattapana", designation: "BUSINESS DEVELOPMENT MANAGER" },
        { employeeCode: "4622", employeeName: "SHYAM KRISHNAN", branchName: "Kattapana", designation: "Collection Executive" },
        { employeeCode: "4385", employeeName: "MINI JAYAN NAIR", branchName: "Koduvayur", designation: "ABM" },
        { employeeCode: "4581", employeeName: "SARATH M R", branchName: "Koduvayur", designation: "LOAN EXECUTIVE" },
        { employeeCode: "4588", employeeName: "SANJITH K", branchName: "Koduvayur", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "3969", employeeName: "AKSHAY A C", branchName: "Kottayam", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "4292", employeeName: "ANOL T ROY", branchName: "Kottayam", designation: "BRANCH MANAGER" },
        { employeeCode: "4501", employeeName: "DHANYAMOL K A", branchName: "Kottayam", designation: "LOAN OFFICER" },
        { employeeCode: "4532", employeeName: "MANJU B NAIR", branchName: "Kottayam", designation: "LOAN EXECUTIVE" },
        { employeeCode: "4602", employeeName: "SWATHY P P", branchName: "Kottayam", designation: "LOAN EXECUTIVE" },
        { employeeCode: "1166", employeeName: "Jancy Antony", branchName: "Kunnamkulam", designation: "Office Executive" },
        { employeeCode: "1530", employeeName: "Sreejitha K K", branchName: "Kunnamkulam", designation: "LOAN OFFICER" },
        { employeeCode: "1173", employeeName: "Jiji Reni", branchName: "Kunnamkulam", designation: "OFFICER" },
        { employeeCode: "4298", employeeName: "JISHA K J", branchName: "Kunnamkulam", designation: "SALES CO - ORDINATOR" },
        { employeeCode: "4550", employeeName: "ANITHA N S", branchName: "Kunnamkulam", designation: "BRANCH MANAGER" },
        { employeeCode: "4608", employeeName: "FAYAZ K A", branchName: "Kunnamkulam", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "4610", employeeName: "SREEJITH K C", branchName: "Kunnamkulam", designation: "ABM" },
        { employeeCode: "4611", employeeName: "AMARJITH A S", branchName: "Kunnamkulam", designation: "ASSOCIATE VICE PRESIDENT" },
        { employeeCode: "4612", employeeName: "SRUTHY N S", branchName: "Kunnamkulam", designation: "LOAN OFFICER" },
        { employeeCode: "4628", employeeName: "JEFERSON M S", branchName: "Kunnamkulam", designation: "DRIVER" },
        { employeeCode: "4621", employeeName: "NIMISHA K U", branchName: "Kunnamkulam", designation: "OFFICER" },
        { employeeCode: "4308", employeeName: "Anuprasad N", branchName: "Kuzhalmannam", designation: "ASSISTANT MANAGER" },
        { employeeCode: "4366", employeeName: "Supriya P", branchName: "Kuzhalmannam", designation: "BRANCH MANAGER" },
        { employeeCode: "4512", employeeName: "PUSHYA U", branchName: "Kuzhalmannam", designation: "LOAN OFFICER" },
        { employeeCode: "4626", employeeName: "ABHAY B", branchName: "Kuzhalmannam", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "4620", employeeName: "SAFA ESSA V M", branchName: "Mattanchery", designation: "OFFICER" },
        { employeeCode: "2733", employeeName: "NITHIN T D", branchName: "Mattanchery", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "3982", employeeName: "RINU C W", branchName: "Mattanchery", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "4340", employeeName: "Mary V A", branchName: "Mattanchery", designation: "BRANCH MANAGER" },
        { employeeCode: "4347", employeeName: "Sangeeth Das", branchName: "Angamaly", designation: "LOAN OFFICER" },
        { employeeCode: "4503", employeeName: "SMITHA M", branchName: "Mavelikara", designation: "LOAN EXECUTIVE" },
        { employeeCode: "4514", employeeName: "GREESHMA VIJAYAN", branchName: "Mavelikara", designation: "MANAGER" },
        { employeeCode: "4511", employeeName: "RENJITH V", branchName: "Mavelikara", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "4527", employeeName: "MAHESH KUMAR M", branchName: "Mavelikara", designation: "COLLECTION MANAGER" },
        { employeeCode: "4544", employeeName: "SREEKALA P R", branchName: "Mavelikara", designation: "BRANCH MANAGER" },
        { employeeCode: "1659", employeeName: "MINI PAUL", branchName: "Muvattupuzha", designation: "ASSISTANT BRANCH MANAGER" },
        { employeeCode: "1838", employeeName: "PRAJESH P P", branchName: "Muvattupuzha", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "4443", employeeName: "K K Vishnumaya", branchName: "Muvattupuzha", designation: "OFFICER" },
        { employeeCode: "4592", employeeName: "SYAM KUMAR M S", branchName: "Muvattupuzha", designation: "BRANCH MANAGER" },
        { employeeCode: "4603", employeeName: "SIJO K CHERIAN", branchName: "Muvattupuzha", designation: "BUSINESS DEVELOPMENT MANAGER" },
        { employeeCode: "4423", employeeName: "Harikrishnan R", branchName: "Nedumkandom", designation: "LOAN OFFICER" },
        { employeeCode: "4496", employeeName: "RESHMA R", branchName: "Nedumkandom", designation: "LOAN OFFICER" },
        { employeeCode: "4573", employeeName: "AMBIKA KUMARY K B", branchName: "Nedumkandom", designation: "SENIOR BRANCH MANAGER" },
        { employeeCode: "4391", employeeName: "VIDYA V", branchName: "Nenmara", designation: "OFFICER" },
        { employeeCode: "4392", employeeName: "RIYAS A", branchName: "Nenmara", designation: "OFFICER" },
        { employeeCode: "4531", employeeName: "SOWRAV KRISHNA", branchName: "Nenmara", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "4627", employeeName: "CHINJU VAISHAK", branchName: "Nenmara", designation: "OFFICER" },
        { employeeCode: "1109", employeeName: "RADHIKA C S", branchName: "Paravoor", designation: "CUSTOMER SERVICE EXECUTIVE" },
        { employeeCode: "1610", employeeName: "SARITHA M S", branchName: "Paravoor", designation: "Sr.EXECUTIVE" },
        { employeeCode: "1130", employeeName: "SWAPNA PRASANTH", branchName: "Paravoor", designation: "LOAN OFFICER" },
        { employeeCode: "2055", employeeName: "ABHILASH K A", branchName: "Paravoor", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "4062", employeeName: "SARIGA N S", branchName: "Paravoor", designation: "OFFICER" },
        { employeeCode: "4547", employeeName: "SANTHI KRISHNA", branchName: "Paravoor", designation: "BRANCH MANAGER" },
        { employeeCode: "4609", employeeName: "AMAL K ANIL", branchName: "Paravoor", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "1299", employeeName: "JIBI BIJI", branchName: "Pathanamthitta", designation: "CUSTOMER SERVICE EXECUTIVE" },
        { employeeCode: "1175", employeeName: "MARIYAMMA GEORGE", branchName: "Pathanamthitta", designation: "SENIOR MANAGER" },
        { employeeCode: "3115", employeeName: "ABISH THOMAS", branchName: "Pathanamthitta", designation: "DEPUTY BRANCH MANAGER" },
        { employeeCode: "3705", employeeName: "ANISHA S NADHAN", branchName: "Pathanamthitta", designation: "OFFICER" },
        { employeeCode: "4369", employeeName: "Aju M D", branchName: "Pathanamthitta", designation: "Collection Executive" },
        { employeeCode: "4451", employeeName: "Vindhya K R", branchName: "Pathanamthitta", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "1032", employeeName: "AMMINI JACOB", branchName: "Perumbavoor", designation: "Sr.EXECUTIVE" },
        { employeeCode: "1218", employeeName: "BIJI REJI", branchName: "Perumbavoor", designation: "RELATIONSHIP EXECUTIVE" },
        { employeeCode: "1051", employeeName: "JAIMOL BENNY", branchName: "Perumbavoor", designation: "OPERATIONS EXECUTIVE" },
        { employeeCode: "1249", employeeName: "JISHA MANOJ", branchName: "Perumbavoor", designation: "CUSTOMER SERVICE EXECUTIVE" },
        { employeeCode: "4046", employeeName: "ARUNDEV T K", branchName: "Perumbavoor", designation: "SECURITY" },
        { employeeCode: "4057", employeeName: "PRIYESH P P", branchName: "Perumbavoor", designation: "BRANCH HEAD" },
        { employeeCode: "4101", employeeName: "AKHIL AMBADAN", branchName: "Perumbavoor", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "1187", employeeName: "SABITHA VENU", branchName: "Corporate Office", designation: "TEAM LEADER" },
        { employeeCode: "2073", employeeName: "JOLLY JOY", branchName: "Corporate Office", designation: "SENIOR EXECUTIVE (DCR TRACKER)" },
        { employeeCode: "2778", employeeName: "DINIYA STEPHEN", branchName: "Corporate Office", designation: "TELECALLING EXECUTIVE" },
        { employeeCode: "3966", employeeName: "RATHEESH KUMAR N R", branchName: "Thiruvalla", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "4080", employeeName: "KARTHIK K", branchName: "Thiruvalla", designation: "BRANCH MANAGER" },
        { employeeCode: "4086", employeeName: "VISHNU T R", branchName: "Thiruvalla", designation: "COLLECTION EXECUTIVE" },
        { employeeCode: "4365", employeeName: "Maneesh Purushothaman", branchName: "Thiruvalla", designation: "OFFICER" },
        { employeeCode: "4361", employeeName: "NAVEEN P", branchName: "Thiruwillamala", designation: "OFFICER" },
        { employeeCode: "4362", employeeName: "SONIA S", branchName: "Thiruwillamala", designation: "OFFICER" },
        { employeeCode: "4515", employeeName: "REMYA N", branchName: "Thiruwillamala", designation: "LOAN EXECUTIVE" },
        { employeeCode: "1185", employeeName: "AKHIL RAJ M", branchName: "Thodupuzha", designation: "SENIOR BRANCH MANAGER" },
        { employeeCode: "1055", employeeName: "ULLAS A N", branchName: "Thodupuzha", designation: "MANAGER" },
        { employeeCode: "2949", employeeName: "ASHA RAJESH", branchName: "Thodupuzha", designation: "BUSINESS EXECUTIVE" },
        { employeeCode: "4344", employeeName: "Shainy Santhosh", branchName: "Thodupuzha", designation: "MANAGER" },
        { employeeCode: "4482", employeeName: "ROOPESH V P", branchName: "Thodupuzha", designation: "ASSISTANT MANAGER" },
        { employeeCode: "2518", employeeName: "Sunny", branchName: "Thodupuzha", designation: "COLLECTION EXECUTIVE" }
    ];

    // --- Column Headers Mapping (IMPORTANT: These must EXACTLY match the column names in your "Form Responses 2" Google Sheet) ---
    const HEADER_TIMESTAMP = 'Timestamp';
    const HEADER_DATE = 'Date';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_ACTIVITY_TYPE = 'Activity Type';
    const HEADER_TYPE_OF_CUSTOMER = 'Type of Customer'; // !!! CORRECTED TYPO HERE !!!
    const HEADER_R_LEAD_SOURCE = 'rLead Source';      // Keeping user's provided interpretation of split header
    const HEADER_HOW_CONTACTED = 'How Contacted'; // This is not in the list provided by user, but is in the original script. Keeping it.
    const HEADER_PROSPECT_NAME = 'Prospect Name';
    const HEADER_PHONE_NUMBER_WHATSAPP = 'Phone Numebr(Whatsapp)'; // Keeping user's provided typo
    const HEADER_ADDRESS = 'Address';
    const HEADER_PROFESSION = 'Profession';
    const HEADER_DOB_WD = 'DOB/WD';
    const HEADER_PRODUCT_INTERESTED = 'Prodcut Interested'; // Keeping user's provided typo
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
    let allUniqueBranches = []; // Will be populated from PREDEFINED_BRANCHES
    let allUniqueEmployees = []; // Employee codes from Canvassing Data
    let employeeCodeToNameMap = {}; // {code: name} from Canvassing Data
    let employeeCodeToDesignationMap = {}; // {code: designation} from Canvassing Data
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
        // Only fetch canvassing data, ignoring MasterEmployees for front-end reports
        await fetchCanvassingData();

        // Re-initialize allUniqueBranches from the predefined list
        allUniqueBranches = [...PREDEFINED_BRANCHES].sort(); // Use the hardcoded list

        // Populate employeeCodeToNameMap and employeeCodeToDesignationMap ONLY from Canvassing Data
        employeeCodeToNameMap = {}; // Reset map before populating
        employeeCodeToDesignationMap = {}; // Reset map before populating
        allCanvassingData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const designation = entry[HEADER_DESIGNATION];

            if (employeeCode) {
                // If an employee code exists in canvassing data, use its name/designation
                employeeCodeToNameMap[employeeCode] = employeeName || employeeCode;
                employeeCodeToDesignationMap[employeeCode] = designation || 'Default';
            }
        });
        
        // Add predefined employees to the maps as well, prioritizing existing canvassing data
        PREDEFINED_EMPLOYEES.forEach(employee => {
            const employeeCode = employee.employeeCode;
            // Only add if not already present from canvassing data, or if you want to overwrite
            if (!employeeCodeToNameMap[employeeCode]) {
                employeeCodeToNameMap[employeeCode] = employee.employeeName;
            }
            if (!employeeCodeToDesignationMap[employeeCode]) {
                employeeCodeToDesignationMap[employeeCode] = employee.designation;
            }
        });


        // Re-populate allUniqueEmployees based ONLY on canvassing data
        allUniqueEmployees = [...new Set([
            ...allCanvassingData.map(entry => entry[HEADER_EMPLOYEE_CODE]),
            ...PREDEFINED_EMPLOYEES.map(employee => employee.employeeCode) // Add predefined employee codes
        ])].sort((codeA, codeB) => {
            // Use the name from the map if available, otherwise use the code for sorting and display
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        });

        populateDropdown(branchSelect, allUniqueBranches); // Populate branch dropdown with predefined branches
        populateDropdown(customerViewBranchSelect, allUniqueBranches); // Populate branch dropdown for detailed customer view
        console.log('Final All Unique Branches (Predefined):', allUniqueBranches);
        console.log('Final Employee Code To Name Map (from Canvassing Data):', employeeCodeToNameMap);
        console.log('Final Employee Code To Designation Map (from Canvassing Data):', employeeCodeToDesignationMap);
        console.log('Final All Unique Employees (Codes from Canvassing Data):', allUniqueEmployees);

        // After data is loaded and maps are populated, render the initial report
        renderAllBranchSnapshot(); // Render the default "All Branch Snapshot" report
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

    // Filter employees based on selected branch
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            employeeFilterPanel.style.display = 'block';

            // Get employee codes ONLY from Canvassing Data for the selected branch
            const employeeCodesInBranchFromCanvassing = allCanvassingData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]);

            // Get employee codes from PREDEFINED_EMPLOYEES for the selected branch
            const employeeCodesInBranchFromPredefined = PREDEFINED_EMPLOYEES
                .filter(employee => employee.branchName === selectedBranch)
                .map(employee => employee.employeeCode);

            // Combine and unique all employee codes for the selected branch
            const combinedEmployeeCodes = new Set([
                ...employeeCodesInBranchFromCanvassing,
                ...employeeCodesInBranchFromPredefined
            ]);

            // Convert Set back to array and sort
            const sortedEmployeeCodesInBranch = [...combinedEmployeeCodes].sort((codeA, codeB) => {
                // Use the name from the map if available, otherwise use the code for sorting and display
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });

            populateDropdown(employeeSelect, sortedEmployeeCodesInBranch, true);
            viewOptions.style.display = 'flex'; // Show view options
            // Reset employee selection and employee-specific display when branch changes
            employeeSelect.value = "";
            selectedEmployeeCodeEntries = []; // Clear previous activity filter
            reportDisplay.innerHTML = '<p>Select an employee or choose a report option.</p>';

            // Deactivate all buttons in viewOptions and then reactivate the appropriate ones
            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));

        } else {
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none'; // Hide view options
            reportDisplay.innerHTML = '<p>Please select a branch from the dropdown above to view reports.</p>';
            selectedBranchEntries = []; // Clear previous activity filter
            selectedEmployeeCodeEntries = []; // Clear previous activity filter
        }
    });

    // Handle employee selection (now based on employee CODE)
    employeeSelect.addEventListener('change', () => {
        const selectedEmployeeCode = employeeSelect.value;
        if (selectedEmployeeCode) {
            // Filter activity data by employee code (from allCanvassingData)
            selectedEmployeeCodeEntries = allCanvassingData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode &&
                entry[HEADER_BRANCH_NAME] === branchSelect.value // Filter by selected branch as well
            );
            const employeeDisplayName = employeeCodeToNameMap[selectedEmployeeCode] || selectedEmployeeCode;
            reportDisplay.innerHTML = `<p>Ready to view reports for ${employeeDisplayName}.</p>`;
            
            // Automatically trigger the Employee Summary (d4.PNG style)
            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
            viewEmployeeSummaryBtn.classList.add('active'); // Set Employee Summary as active
            renderEmployeeSummary(selectedEmployeeCodeEntries); // Render the Employee Summary
            
        } else {
            selectedEmployeeCodeEntries = []; // Clear previous activity filter
            reportDisplay.innerHTML = '<p>Select an employee or choose a report option.S</p>';
            // Clear active button if employee selection is cleared
            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        }
    });

    // Helper to calculate total activity from a set of activity entries based on Activity Type
    function calculateTotalActivity(entries) {
        const totalActivity = { 'Visit': 0, 'Call': 0, 'Reference': 0, 'New Customer Leads': 0 }; // Initialize counters
        const productInterests = new Set(); // To collect unique product interests
        
        console.log('Calculating total activity for entries:', entries.length); // Log entries being processed
        entries.forEach((entry, index) => {
            let activityType = entry[HEADER_ACTIVITY_TYPE];
            let typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER];
            let productInterested = entry[HEADER_PRODUCT_INTERESTED]; // Get product interested

            // Trim and convert to lowercase for robust comparison
            const trimmedActivityType = activityType ? activityType.trim().toLowerCase() : '';
            const trimmedTypeOfCustomer = typeOfCustomer ? typeOfCustomer.trim().toLowerCase() : '';
            const trimmedProductInterested = productInterested ? productInterested.trim() : ''; // Don't lowercase products unless explicitly asked

            console.log(`--- Entry ${index + 1} Debug ---`);
            console.log(`  Processed Activity Type (trimmed, lowercase): '${trimmedActivityType}'`);
            console.log(`  Processed Type of Customer (trimmed, lowercase): '${trimmedTypeOfCustomer}'`);
            console.log(`  Processed Product Interested (trimmed): '${trimmedProductInterested}'`);


            // Direct matching to user's provided sheet values (now lowercase)
            if (trimmedActivityType === 'visit') {
                totalActivity['Visit']++;
            } else if (trimmedActivityType === 'calls') { // Matches "Calls" from sheet, now lowercase
                totalActivity['Call']++;
            } else if (trimmedActivityType === 'referance') { // Matches "Referance" (with typo) from sheet, now lowercase
                totalActivity['Reference']++;
            } else {
                // If it's not one of the direct activity types, log for debugging
                console.warn(`  Unknown or unhandled Activity Type encountered (trimmed, lowercase): '${trimmedActivityType}'.`);
            }
            
            // --- UPDATED LOGIC FOR 'New Customer Leads' ---
            // Based on the user's previously working script, New Customer Leads are counted
            // if the 'Type of Customer' (now correctly spelled) is simply 'new', regardless of 'Activity Type'.
            if (trimmedTypeOfCustomer === 'new') {
                totalActivity['New Customer Leads']++;
                console.log(`  New Customer Lead INCREMENTED based on Type of Customer === 'new'.`);
            } else {
                console.log(`  New Customer Lead NOT INCREMENTED: Type of Customer is not 'new'.`);
            }
            // --- END UPDATED LOGIC ---

            // Collect unique product interests
            if (trimmedProductInterested) {
                productInterests.add(trimmedProductInterested);
            }
            console.log(`--- End Entry ${index + 1} Debug ---`);
        });
        console.log('Calculated Total Activity Final:', totalActivity);
        
        // Return both total activities and product interests
        return { totalActivity, productInterests: [...productInterests] };
    }

    // Render All Branch Snapshot (now uses PREDEFINED_BRANCHES and checks for participation)
    function renderAllBranchSnapshot() {
        reportDisplay.innerHTML = '<h2>All Branch Snapshot</h2>';
        
        const table = document.createElement('table');
        table.className = 'all-branch-snapshot-table';
        
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const headers = ['Branch Name', 'Employees with Activity', 'Total Visits', 'Total Calls', 'Total References', 'Total New Customer Leads'];
        headers.forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();

        PREDEFINED_BRANCHES.forEach(branch => {
            const branchActivityEntries = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
            const { totalActivity } = calculateTotalActivity(branchActivityEntries); // Destructure to get totalActivity
            const employeeCodesInBranch = [...new Set(branchActivityEntries.map(entry => entry[HEADER_EMPLOYEE_CODE]))];
            const displayEmployeeCount = employeeCodesInBranch.length;

            const row = tbody.insertRow();
            // Assign data-label for mobile view
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

        reportDisplay.appendChild(table);
    }

    // NEW: Render Non-Participating Branches Report
    function renderNonParticipatingBranches() {
        reportDisplay.innerHTML = '<h2>Non-Participating Branches</h2>';
        const nonParticipatingBranches = [];
        PREDEFINED_BRANCHES.forEach(branch => {
            const hasActivity = allCanvassingData.some(entry => entry[HEADER_BRANCH_NAME] === branch);
            if (!hasActivity) {
                nonParticipatingBranches.push(branch);
            }
        });

        if (nonParticipatingBranches.length > 0) {
            const ul = document.createElement('ul');
            ul.className = 'non-participating-branch-list';
            nonParticipatingBranches.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                ul.appendChild(li);
            });
            reportDisplay.appendChild(ul);
        } else {
            reportDisplay.innerHTML += '<p class="no-participation-message">All predefined branches have recorded activity!</p>';
        }
    }

    // Render All Staff Overall Performance Report (for d1.PNG)
    function renderOverallStaffPerformanceReport() {
        reportDisplay.innerHTML = '<h2>Overall Staff Performance Report (This Month)</h2>';
        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container'; // For horizontal scrolling
        const table = document.createElement('table');
        table.className = 'performance-table';
        const thead = table.createTHead();
        let headerRow = thead.insertRow();
        // Main Headers
        headerRow.insertCell().textContent = 'Employee Name';
        headerRow.insertCell().textContent = 'Branch Name';
        headerRow.insertCell().textContent = 'Designation';
        // Define metrics for the performance table
        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
        metrics.forEach(metric => {
            const th = document.createElement('th');
            th.colSpan = 3; // 'Actual', 'Target', '%'
            th.textContent = metric;
            headerRow.appendChild(th);
        });

        // Sub-headers
        headerRow = thead.insertRow(); // New row for sub-headers
        headerRow.insertCell(); // Empty cell for Employee Name
        headerRow.insertCell(); // Empty cell for Branch Name
        headerRow.insertCell(); // Empty cell for Designation
        metrics.forEach(() => {
            ['Act', 'Tgt', '%'].forEach(subHeader => {
                const th = document.createElement('th');
                th.textContent = subHeader;
                headerRow.appendChild(th);
            });
        });

        const tbody = table.createTBody();

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        // Get unique employees who have made at least one entry this month
        const employeesWithActivityThisMonth = [...new Set(allCanvassingData
            .filter(entry => {
                const entryDate = new Date(entry[HEADER_TIMESTAMP]);
                return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
            })
            .map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        });

        if (employeesWithActivityThisMonth.length === 0) {
            reportDisplay.innerHTML += '<p>No employee activity found for the current month.</p>';
            return;
        }

        employeesWithActivityThisMonth.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const branchName = allCanvassingData.find(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode)?.[HEADER_BRANCH_NAME] || 'N/A';
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const employeeActivities = allCanvassingData.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode && new Date(entry[HEADER_TIMESTAMP]).getMonth() === currentMonth && new Date(entry[HEADER_TIMESTAMP]).getFullYear() === currentYear);
            const { totalActivity } = calculateTotalActivity(employeeActivities);

            const row = tbody.insertRow();
            row.insertCell().textContent = employeeName;
            row.insertCell().textContent = branchName;
            row.insertCell().textContent = designation;

            const targets = TARGETS[designation] || TARGETS['Default'];

            metrics.forEach(metric => {
                const actual = totalActivity[metric];
                const target = targets[metric];
                const percentage = target > 0 ? ((actual / target) * 100).toFixed(0) : 'N/A';

                row.insertCell().textContent = actual;
                row.insertCell().textContent = target;
                row.insertCell().textContent = percentage !== 'N/A' ? `${percentage}%` : 'N/A';
            });
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Render Employee Summary (for d4.PNG)
    function renderEmployeeSummary(entries) {
        if (!entries || entries.length === 0) {
            reportDisplay.innerHTML = '<p>No activity data for this employee in the selected branch.</p>';
            return;
        }

        // Get employee details from the first entry (assuming consistency)
        const employeeCode = entries[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const branchName = entries[0][HEADER_BRANCH_NAME];
        const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

        reportDisplay.innerHTML = `<h2>Employee Summary: ${employeeName} (${employeeCode})</h2>`;
        reportDisplay.innerHTML += `<p><strong>Branch:</strong> ${branchName}</p>`;
        reportDisplay.innerHTML += `<p><strong>Designation:</strong> ${designation}</p>`;

        const { totalActivity, productInterests } = calculateTotalActivity(entries);

        // Display Activity Summary
        const activitySummaryDiv = document.createElement('div');
        activitySummaryDiv.className = 'activity-summary';
        activitySummaryDiv.innerHTML = '<h3>Activity Summary</h3>';
        activitySummaryDiv.innerHTML += `<ul>
            <li><strong>Total Visits:</strong> ${totalActivity['Visit']}</li>
            <li><strong>Total Calls:</strong> ${totalActivity['Call']}</li>
            <li><strong>Total References:</strong> ${totalActivity['Reference']}</li>
            <li><strong>Total New Customer Leads:</strong> ${totalActivity['New Customer Leads']}</li>
        </ul>`;
        reportDisplay.appendChild(activitySummaryDiv);

        // Display Product Interests
        if (productInterests.length > 0) {
            const productInterestDiv = document.createElement('div');
            productInterestDiv.className = 'product-interests';
            productInterestDiv.innerHTML = '<h3>Products Interested</h3>';
            productInterestDiv.innerHTML += `<p>${productInterests.join(', ')}</p>`;
            reportDisplay.appendChild(productInterestDiv);
        } else {
            reportDisplay.innerHTML += '<p>No specific product interests recorded.</p>';
        }

        // Display Recent Activity (last 5 entries)
        const recentActivityDiv = document.createElement('div');
        recentActivityDiv.className = 'recent-activity';
        recentActivityDiv.innerHTML = '<h3>Recent Activities</h3>';

        const sortedEntries = [...entries].sort((a, b) => new Date(b[HEADER_TIMESTAMP]) - new Date(a[HEADER_TIMESTAMP]));
        const recentFiveEntries = sortedEntries.slice(0, 5);

        if (recentFiveEntries.length > 0) {
            const ul = document.createElement('ul');
            recentFiveEntries.forEach(entry => {
                const li = document.createElement('li');
                li.innerHTML = `<strong>Date:</strong> ${formatDate(entry[HEADER_DATE]) || formatDate(entry[HEADER_TIMESTAMP])},
                                <strong>Activity:</strong> ${entry[HEADER_ACTIVITY_TYPE] || 'N/A'},
                                <strong>Customer:</strong> ${entry[HEADER_PROSPECT_NAME] || 'N/A'}
                                (${entry[HEADER_TYPE_OF_CUSTOMER] || 'N/A'})`;
                ul.appendChild(li);
            });
            recentActivityDiv.appendChild(ul);
        } else {
            recentActivityDiv.innerHTML += '<p>No recent activities found.</p>';
        }
        reportDisplay.appendChild(recentActivityDiv);
    }


    // Render Branch Performance Report (for d3.PNG)
    function renderBranchPerformanceReport() {
        const selectedBranch = branchSelect.value;
        if (!selectedBranch) {
            reportDisplay.innerHTML = '<p>Please select a branch to view the Branch Performance Report.</p>';
            return;
        }

        reportDisplay.innerHTML = `<h2>Branch Performance Report: ${selectedBranch}</h2>`;

        // Filter activities for the selected branch
        const branchActivityEntries = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch);

        // Get unique employees for the selected branch (from combined canvassing and predefined data)
        const employeesInBranch = [...new Set(
            branchActivityEntries.map(entry => entry[HEADER_EMPLOYEE_CODE]).concat(
                PREDEFINED_EMPLOYEES.filter(emp => emp.branchName === selectedBranch).map(emp => emp.employeeCode)
            )
        )].sort((codeA, codeB) => {
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        });

        if (employeesInBranch.length === 0) {
            reportDisplay.innerHTML += '<p>No employees found for this branch or no activity recorded.</p>';
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';
        const table = document.createElement('table');
        table.className = 'performance-table';
        const thead = table.createTHead();
        let headerRow = thead.insertRow();

        headerRow.insertCell().textContent = 'Employee Name';
        headerRow.insertCell().textContent = 'Designation';

        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
        metrics.forEach(metric => {
            const th = document.createElement('th');
            th.colSpan = 3; // Actual, Target, %
            th.textContent = metric;
            headerRow.appendChild(th);
        });

        // Sub-headers
        headerRow = thead.insertRow(); // New row for sub-headers
        headerRow.insertCell(); // Empty for Employee Name
        headerRow.insertCell(); // Empty for Designation
        metrics.forEach(() => {
            ['Act', 'Tgt', '%'].forEach(subHeader => {
                const th = document.createElement('th');
                th.textContent = subHeader;
                headerRow.appendChild(th);
            });
        });

        const tbody = table.createTBody();

        employeesInBranch.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const employeeActivitiesInBranch = branchActivityEntries.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
            const { totalActivity } = calculateTotalActivity(employeeActivitiesInBranch);

            const row = tbody.insertRow();
            row.insertCell().textContent = employeeName;
            row.insertCell().textContent = designation;

            const targets = TARGETS[designation] || TARGETS['Default'];

            metrics.forEach(metric => {
                const actual = totalActivity[metric];
                const target = targets[metric];
                const percentage = target > 0 ? ((actual / target) * 100).toFixed(0) : 'N/A';

                row.insertCell().textContent = actual;
                row.insertCell().textContent = target;
                row.insertCell().textContent = percentage !== 'N/A' ? `${percentage}%` : 'N/A';
            });
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }


    // Event Listeners for main report view buttons
    viewBranchPerformanceReportBtn.addEventListener('click', () => {
        // Deactivate all buttons first
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        // Activate the clicked button
        viewBranchPerformanceReportBtn.classList.add('active');
        renderBranchPerformanceReport();
    });

    viewEmployeeSummaryBtn.addEventListener('click', () => {
        const selectedEmployeeCode = employeeSelect.value;
        if (!selectedEmployeeCode) {
            reportDisplay.innerHTML = '<p>Please select an employee to view the Employee Summary.</p>';
            return;
        }
        // Deactivate all buttons first
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        // Activate the clicked button
        viewEmployeeSummaryBtn.classList.add('active');
        // Use the already filtered selectedEmployeeCodeEntries
        renderEmployeeSummary(selectedEmployeeCodeEntries);
    });

    viewAllEntriesBtn.addEventListener('click', () => {
        const selectedBranch = branchSelect.value;
        const selectedEmployeeCode = employeeSelect.value;
        if (!selectedBranch && !selectedEmployeeCode) {
            reportDisplay.innerHTML = '<p>Please select a branch or an employee to view all entries.</p>';
            return;
        }

        // Deactivate all buttons first
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        // Activate the clicked button
        viewAllEntriesBtn.classList.add('active');
        renderAllEntries(selectedBranch, selectedEmployeeCode);
    });

    // Function to render all entries
    function renderAllEntries(branchName, employeeCode) {
        reportDisplay.innerHTML = '<h2>All Entries</h2>';
        let filteredEntries = allCanvassingData;

        if (branchName) {
            filteredEntries = filteredEntries.filter(entry => entry[HEADER_BRANCH_NAME] === branchName);
            reportDisplay.innerHTML = `<h2>All Entries for ${branchName}</h2>`;
        }
        if (employeeCode) {
            filteredEntries = filteredEntries.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
            reportDisplay.innerHTML = `<h2>All Entries for ${employeeCodeToNameMap[employeeCode] || employeeCode} in ${branchName}</h2>`;
        }

        if (filteredEntries.length === 0) {
            reportDisplay.innerHTML += '<p>No entries found matching the criteria.</p>';
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';
        const table = document.createElement('table');
        table.className = 'all-entries-table';

        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        // Dynamically create headers from the first entry's keys, excluding internal ones
        const headersToDisplay = Object.keys(filteredEntries[0]).filter(header =>
            ![HEADER_TIMESTAMP].includes(header)
        );

        headersToDisplay.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        filteredEntries.forEach(entry => {
            const row = tbody.insertRow();
            headersToDisplay.forEach(header => {
                const cell = row.insertCell();
                let cellValue = entry[header];
                // Format Date and DOB/WD columns
                if (header === HEADER_DATE || header === HEADER_DOB_WD || header === HEADER_NEXT_FOLLOW_UP_DATE) {
                    cellValue = formatDate(cellValue);
                }
                cell.textContent = cellValue || 'N/A';
            });
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }


    viewPerformanceReportBtn.addEventListener('click', () => {
        // Deactivate all buttons first
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        // Activate the clicked button
        viewPerformanceReportBtn.classList.add('active');
        renderOverallStaffPerformanceReport();
    });



    // Event Listeners for Main Tab Navigation (Image: d2.PNG - tabs at top)
    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    detailedCustomerViewTabBtn.addEventListener('click', () => showTab('detailedCustomerViewTabBtn')); // NEW
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));


    // Function to show/hide sections and render reports based on active tab
    function showTab(activeTabId) {
        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
        // Activate the clicked tab button
        document.getElementById(activeTabId).classList.add('active');

        // Hide all main content sections
        reportsSection.style.display = 'none';
        detailedCustomerViewSection.style.display = 'none'; // NEW
        employeeManagementSection.style.display = 'none';

        // Reset filter and report displays
        branchSelect.value = "";
        employeeSelect.innerHTML = '<option value="">-- Select --</option>';
        employeeFilterPanel.style.display = 'none';
        viewOptions.style.display = 'none';
        reportDisplay.innerHTML = ''; // Clear report display area

        // Show the relevant section and render the appropriate content
        switch (activeTabId) {
            case 'allBranchSnapshotTabBtn':
                reportsSection.style.display = 'block';
                renderAllBranchSnapshot();
                break;
            case 'allStaffOverallPerformanceTabBtn':
                reportsSection.style.display = 'block';
                renderOverallStaffPerformanceReport();
                break;
            case 'nonParticipatingBranchesTabBtn':
                reportsSection.style.display = 'block';
                renderNonParticipatingBranches();
                break;
            case 'detailedCustomerViewTabBtn': // NEW
                detailedCustomerViewSection.style.display = 'block';
                // Initialize customer view dropdowns
                populateDropdown(customerViewBranchSelect, allUniqueBranches);
                customerCanvassedList.innerHTML = '<p>Select a branch and employee to view canvassed customers.</p>';
                customerDetailsContent.innerHTML = ''; // Clear customer details
                break;
            case 'employeeManagementTabBtn':
                employeeManagementSection.style.display = 'block';
                // Populate branch dropdowns in employee management forms
                populateDropdown(newBranchNameInput, allUniqueBranches);
                populateDropdown(bulkEmployeeBranchNameInput, allUniqueBranches);
                break;
        }
    }


    // NEW: Detailed Customer View Logic
    customerViewBranchSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        if (selectedBranch) {
            // Filter employees for the selected branch (from combined canvassing and predefined)
            const employeeCodesInBranch = [...new Set(
                allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch).map(entry => entry[HEADER_EMPLOYEE_CODE]).concat(
                PREDEFINED_EMPLOYEES.filter(emp => emp.branchName === selectedBranch).map(emp => emp.employeeCode)
                )
            )].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });
            populateDropdown(customerViewEmployeeSelect, employeeCodesInBranch, true);
            customerCanvassedList.innerHTML = '<p>Select an employee to view their canvassed customers.</p>';
            customerDetailsContent.innerHTML = '';
        } else {
            customerViewEmployeeSelect.innerHTML = '<option value="">-- Select --</option>';
            customerCanvassedList.innerHTML = '<p>Select a branch and employee to view canvassed customers.</p>';
            customerDetailsContent.innerHTML = '';
        }
    });

    customerViewEmployeeSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        const selectedEmployeeCode = customerViewEmployeeSelect.value;
        if (selectedBranch && selectedEmployeeCode) {
            renderCanvassedCustomers(selectedBranch, selectedEmployeeCode);
        } else {
            customerCanvassedList.innerHTML = '<p>Select an employee to view their canvassed customers.</p>';
            customerDetailsContent.innerHTML = '';
        }
    });

    function renderCanvassedCustomers(branchName, employeeCode) {
        customerCanvassedList.innerHTML = ''; // Clear previous list
        customerDetailsContent.innerHTML = ''; // Clear previous details

        const customers = allCanvassingData.filter(entry =>
            entry[HEADER_BRANCH_NAME] === branchName &&
            entry[HEADER_EMPLOYEE_CODE] === employeeCode
        );

        if (customers.length === 0) {
            customerCanvassedList.innerHTML = '<p>No canvassed customers found for this employee in this branch.</p>';
            return;
        }

        const ul = document.createElement('ul');
        ul.className = 'customer-list';
        customers.forEach((customer, index) => {
            const li = document.createElement('li');
            li.textContent = `${customer[HEADER_PROSPECT_NAME] || 'Unknown Customer'} (${customer[HEADER_ACTIVITY_TYPE] || 'N/A'}) - ${formatDate(customer[HEADER_DATE] || customer[HEADER_TIMESTAMP])}`;
            li.addEventListener('click', () => displayCustomerDetails(customer));
            ul.appendChild(li);
        });
        customerCanvassedList.appendChild(ul);
    }

    function displayCustomerDetails(customer) {
        customerDetailsContent.innerHTML = '<h3>Customer Details</h3>';
        const detailsHtml = `
            <p><strong>Prospect Name:</strong> ${customer[HEADER_PROSPECT_NAME] || 'N/A'}</p>
            <p><strong>Phone Number (Whatsapp):</strong> ${customer[HEADER_PHONE_NUMBER_WHATSAPP] || 'N/A'}</p>
            <p><strong>Address:</strong> ${customer[HEADER_ADDRESS] || 'N/A'}</p>
            <p><strong>Profession:</strong> ${customer[HEADER_PROFESSION] || 'N/A'}</p>
            <p><strong>DOB/WD:</strong> ${formatDate(customer[HEADER_DOB_WD]) || 'N/A'}</p>
            <p><strong>Product Interested:</strong> ${customer[HEADER_PRODUCT_INTERESTED] || 'N/A'}</p>
            <p><strong>Remarks:</strong> ${customer[HEADER_REMARKS] || 'N/A'}</p>
            <p><strong>Next Follow-up Date:</strong> ${formatDate(customer[HEADER_NEXT_FOLLOW_UP_DATE]) || 'N/A'}</p>
            <p><strong>Activity Type:</strong> ${customer[HEADER_ACTIVITY_TYPE] || 'N/A'}</p>
            <p><strong>Type of Customer:</strong> ${customer[HEADER_TYPE_OF_CUSTOMER] || 'N/A'}</p>
            <p><strong>Lead Source:</strong> ${customer[HEADER_R_LEAD_SOURCE] || 'N/A'}</p>
            <p><strong>How Contacted:</strong> ${customer[HEADER_HOW_CONTACTED] || 'N/A'}</p>
            <p><strong>Relation With Staff:</strong> ${customer[HEADER_RELATION_WITH_STAFF] || 'N/A'}</p>
            <h4>Family Details:</h4>
            <ul>
                <li><strong>Name of Wife/Husband:</strong> ${customer[HEADER_FAMILY_DETAILS_1] || 'N/A'}</li>
                <li><strong>Job of Wife/Husband:</strong> ${customer[HEADER_FAMILY_DETAILS_2] || 'N/A'}</li>
                <li><strong>Names of Children:</strong> ${customer[HEADER_FAMILY_DETAILS_3] || 'N/A'}</li>
                <li><strong>Details of Children:</strong> ${customer[HEADER_FAMILY_DETAILS_4] || 'N/A'}</li>
            </ul>
            <p><strong>Profile of Customer:</strong> ${customer[HEADER_PROFILE_OF_CUSTOMER] || 'N/A'}</p>
        `;
        customerDetailsContent.innerHTML += detailsHtml;
    }


    // Function to send data to Google Apps Script Web App
    async function sendDataToGoogleAppsScript(action, data) {
        displayEmployeeManagementMessage(`Sending data to server for ${action}...`, false);
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Crucial for cross-origin requests
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ action, data }),
            });

            if (!response.ok) {
                const errorText = await response.text();
                console.error(`Google Apps Script Web App Error for ${action}! Status: ${response.status}. Details: ${errorText}`);
                displayEmployeeManagementMessage(`Error processing ${action}: ${errorText}`, true);
                return false;
            }

            const result = await response.json();
            console.log(`Google Apps Script ${action} response:`, result);

            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(`${result.message || 'Operation successful!'}`);
                // Re-fetch and process data to update UI after successful operation
                await processData();
                return true;
            } else {
                displayEmployeeManagementMessage(`${result.message || 'Operation failed!'}: ${result.details || ''}`, true);
                return false;
            }
        } catch (error) {
            console.error(`Network or unexpected error during ${action}:`, error);
            displayEmployeeManagementMessage(`Network error during ${action}: ${error.message}`, true);
            return false;
        }
    }


    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault(); // Prevent default form submission

            const employeeName = newEmployeeNameInput.value.trim();
            const employeeCode = newEmployeeCodeInput.value.trim();
            const branchName = newBranchNameInput.value.trim();
            const designation = newDesignationInput.value.trim();

            if (!employeeName || !employeeCode || !branchName || !designation) {
                displayEmployeeManagementMessage('All fields are required for adding an employee.', true);
                return;
            }

            const employeeData = {
                [HEADER_EMPLOYEE_CODE]: employeeCode,
                [HEADER_EMPLOYEE_NAME]: employeeName,
                [HEADER_BRANCH_NAME]: branchName,
                [HEADER_DESIGNATION]: designation
            };

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);

            if (success) {
                addEmployeeForm.reset(); // Clear form on success
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
                displayEmployeeManagementMessage('Branch Name and Bulk Details are required for bulk addition.', true);
                return;
            }

            const employeesToAdd = [];
            const lines = bulkDetails.split('\n').filter(line => line.trim() !== '');

            for (const line of lines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length >= 2) { // Expecting at least employeeCode, employeeName
                    const employeeData = {
                        [HEADER_EMPLOYEE_CODE]: parts[0],
                        [HEADER_EMPLOYEE_NAME]: parts[1],
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
