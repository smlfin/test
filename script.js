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
        "Muvattupuzha", "Thiruvalla", "Pathanamthitta", "HO KKM", "Kunnamkulam" // Corrected "Pathanamthitta" typo if it existed previously
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
        displayMessage("Activity data fetched successfully!", 'success');
    } catch (error) {
        console.error("Error fetching canvassing data:", error);
        displayMessage(`Error fetching activity data: ${error.message}`, 'error');
        allCanvassingData = []; // Ensure data is cleared on error
    }
}
 // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn');
});
