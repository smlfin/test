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

        // Re-populate allUniqueEmployees based ONLY on canvassing data
        allUniqueEmployees = [...new Set(allCanvassingData.map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
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

            // Combine and unique all employee codes for the selected branch
            const combinedEmployeeCodes = new Set([
                ...employeeCodesInBranchFromCanvassing
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

    // NEW: Render Non-Participating Employees Report
    function renderNonParticipatingEmployees() {
        reportDisplay.innerHTML = '<h2>Non-Participating Employees (Current Month)</h2>';
        const nonParticipatingEmployees = [];
        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        // Get employee codes who have recorded *any* activity in the current month
        const employeesWithActivityThisMonth = new Set(
            allCanvassingData
                .filter(entry => {
                    const entryDate = new Date(entry[HEADER_TIMESTAMP]);
                    return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
                })
                .map(entry => entry[HEADER_EMPLOYEE_CODE])
        );

        // allUniqueEmployees contains all employees who have *ever* submitted data
        allUniqueEmployees.forEach(employeeCode => {
            if (!employeesWithActivityThisMonth.has(employeeCode)) {
                // Get display name, default to code if name is not available
                const employeeDisplayName = employeeCodeToNameMap[employeeCode] || employeeCode;
                nonParticipatingEmployees.push(employeeDisplayName);
            }
        });

        if (nonParticipatingEmployees.length > 0) {
            const ul = document.createElement('ul');
            ul.className = 'non-participating-employee-list'; // Reusing style from branch list or adding a new one
            nonParticipatingEmployees.sort().forEach(employeeName => { // Sort by name for better readability
                const li = document.createElement('li');
                li.textContent = employeeName;
                ul.appendChild(li);
            });
            reportDisplay.appendChild(ul);
        } else {
            reportDisplay.innerHTML += '<p class="no-participation-message">All employees with historical data have recorded activity this month!</p>';
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
        let headerRow = thead.insertRow(); // Main Headers
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
            reportDisplay.innerHTML += '<p class="no-participation-message">No staff recorded activities this month.</p>';
            return;
        }

        employeesWithActivityThisMonth.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';
            
            // Find the branch name for the employee from the first relevant entry
            const employeeBranchEntry = allCanvassingData.find(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
            const branchName = employeeBranchEntry ? employeeBranchEntry[HEADER_BRANCH_NAME] : 'N/A';

            // Filter entries for the current month for this specific employee
            const employeeMonthlyActivities = allCanvassingData.filter(entry => {
                const entryDate = new Date(entry[HEADER_TIMESTAMP]);
                return entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
                       entryDate.getMonth() === currentMonth &&
                       entryDate.getFullYear() === currentYear;
            });

            const { totalActivity } = calculateTotalActivity(employeeMonthlyActivities);
            const targets = TARGETS[designation] || TARGETS['Default'];

            const row = tbody.insertRow();
            row.insertCell().textContent = employeeName;
            row.insertCell().textContent = branchName;
            row.insertCell().textContent = designation;

            metrics.forEach(metric => {
                const actual = totalActivity[metric];
                const target = targets[metric];
                const percentage = target > 0 ? ((actual / target) * 100).toFixed(0) : (actual > 0 ? 100 : 0);

                row.insertCell().textContent = actual;
                row.insertCell().textContent = target;
                
                const percentCell = row.insertCell();
                const progressBarContainer = document.createElement('div');
                progressBarContainer.className = 'progress-bar-container-small';
                
                const progressBar = document.createElement('div');
                progressBar.className = 'progress-bar';
                progressBar.style.width = `${Math.min(100, percentage)}%`; // Cap at 100% for display
                progressBar.textContent = `${percentage}%`;

                if (percentage >= 100) {
                    progressBar.classList.add('success');
                } else if (percentage >= 75) {
                    progressBar.classList.add('warning-high');
                } else if (percentage >= 50) {
                    progressBar.classList.add('warning-medium');
                } else if (percentage > 0) {
                    progressBar.classList.add('warning-low');
                } else {
                    progressBar.classList.add('no-activity');
                    progressBar.textContent = '0%'; // Explicitly show 0% if no activity
                }
                progressBarContainer.appendChild(progressBar);
                percentCell.appendChild(progressBarContainer);
            });
        });
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Render Branch Performance Report (for d3.PNG)
    function renderBranchPerformanceReport() {
        reportDisplay.innerHTML = `<h2>${branchSelect.value} Branch Performance Report (Current Month)</h2>`;
        const selectedBranch = branchSelect.value;
        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        const branchMonthlyActivities = allCanvassingData.filter(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            return entry[HEADER_BRANCH_NAME] === selectedBranch &&
                   entryDate.getMonth() === currentMonth &&
                   entryDate.getFullYear() === currentYear;
        });

        if (branchMonthlyActivities.length === 0) {
            reportDisplay.innerHTML += `<p class="no-participation-message">No activities recorded for ${selectedBranch} branch this month.</p>`;
            return;
        }

        const employeesInBranchThisMonth = [...new Set(branchMonthlyActivities.map(entry => entry[HEADER_EMPLOYEE_CODE]))];

        const gridContainer = document.createElement('div');
        gridContainer.className = 'branch-performance-grid';

        employeesInBranchThisMonth.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';
            const targets = TARGETS[designation] || TARGETS['Default'];

            const employeeActivities = branchMonthlyActivities.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
            const { totalActivity, productInterests } = calculateTotalActivity(employeeActivities);

            const card = document.createElement('div');
            card.className = 'employee-performance-card';
            card.innerHTML = `
                <h3>${employeeName}</h3>
                <p><strong>Designation:</strong> ${designation}</p>
                <div class="performance-metrics">
                    <div class="metric-item">
                        <span>Visits:</span>
                        <div class="progress-bar-container">
                            <div class="progress-bar ${getProgressColorClass(totalActivity['Visit'], targets['Visit'])}" style="width: ${getProgressBarWidth(totalActivity['Visit'], targets['Visit'])}%;">
                                ${totalActivity['Visit']} / ${targets['Visit']}
                            </div>
                        </div>
                    </div>
                    <div class="metric-item">
                        <span>Calls:</span>
                        <div class="progress-bar-container">
                            <div class="progress-bar ${getProgressColorClass(totalActivity['Call'], targets['Call'])}" style="width: ${getProgressBarWidth(totalActivity['Call'], targets['Call'])}%;">
                                ${totalActivity['Call']} / ${targets['Call']}
                            </div>
                        </div>
                    </div>
                    <div class="metric-item">
                        <span>References:</span>
                        <div class="progress-bar-container">
                            <div class="progress-bar ${getProgressColorClass(totalActivity['Reference'], targets['Reference'])}" style="width: ${getProgressBarWidth(totalActivity['Reference'], targets['Reference'])}%;">
                                ${totalActivity['Reference']} / ${targets['Reference']}
                            </div>
                        </div>
                    </div>
                    <div class="metric-item">
                        <span>New Leads:</span>
                        <div class="progress-bar-container">
                            <div class="progress-bar ${getProgressColorClass(totalActivity['New Customer Leads'], targets['New Customer Leads'])}" style="width: ${getProgressBarWidth(totalActivity['New Customer Leads'], targets['New Customer Leads'])}%;">
                                ${totalActivity['New Customer Leads']} / ${targets['New Customer Leads']}
                            </div>
                        </div>
                    </div>
                </div>
                <div class="product-interest-section">
                    <h4>Products Interested:</h4>
                    <p>${productInterests.length > 0 ? productInterests.join(', ') : 'N/A'}</p>
                </div>
            `;
            gridContainer.appendChild(card);
        });
        reportDisplay.appendChild(gridContainer);
    }

    // Helper for progress bar color
    function getProgressColorClass(actual, target) {
        if (target === 0) return 'no-activity'; // Or a default color if target is zero but there's activity
        const percentage = (actual / target) * 100;
        if (percentage >= 100) return 'success';
        if (percentage >= 75) return 'warning-high';
        if (percentage >= 50) return 'warning-medium';
        if (percentage > 0) return 'warning-low';
        return 'no-activity';
    }

    // Helper for progress bar width
    function getProgressBarWidth(actual, target) {
        if (target === 0) return actual > 0 ? 100 : 0; // If target is 0, 100% if any activity, else 0%
        return Math.min(100, (actual / target) * 100);
    }

    // Render Employee Summary (Current Month) for d4.PNG
    function renderEmployeeSummary(entries) {
        const selectedEmployeeCode = employeeSelect.value;
        const employeeName = employeeCodeToNameMap[selectedEmployeeCode] || selectedEmployeeCode;
        const designation = employeeCodeToDesignationMap[selectedEmployeeCode] || 'Default';
        const targets = TARGETS[designation] || TARGETS['Default'];

        reportDisplay.innerHTML = `<h2>${employeeName}'s Performance Summary (Current Month)</h2>`;

        if (entries.length === 0) {
            reportDisplay.innerHTML += `<p class="no-participation-message">No activities recorded for ${employeeName} this month.</p>`;
            return;
        }

        const { totalActivity, productInterests } = calculateTotalActivity(entries);

        const summaryCard = document.createElement('div');
        summaryCard.className = 'summary-breakdown-card'; // Main card for layout

        // Section 1: Overall Activity Summary (Top Left)
        const overallSummarySection = document.createElement('div');
        overallSummarySection.className = 'summary-section overall-activity-summary';
        overallSummarySection.innerHTML = `
            <h3>Overall Activity</h3>
            <p><strong>Designation:</strong> ${designation}</p>
            <div class="overall-metrics">
                <div><span>Total Visits:</span> <span>${totalActivity['Visit']}</span></div>
                <div><span>Total Calls:</span> <span>${totalActivity['Call']}</span></div>
                <div><span>Total References:</span> <span>${totalActivity['Reference']}</span></div>
                <div><span>Total New Customer Leads:</span> <span>${totalActivity['New Customer Leads']}</span></div>
            </div>
        `;
        summaryCard.appendChild(overallSummarySection);

        // Section 2: Progress Against Targets (Top Right)
        const progressSection = document.createElement('div');
        progressSection.className = 'summary-section progress-targets';
        progressSection.innerHTML = `<h3>Progress Against Targets</h3>`;
        const progressList = document.createElement('div');
        progressList.className = 'progress-list';

        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
        metrics.forEach(metric => {
            const actual = totalActivity[metric];
            const target = targets[metric];
            const percentage = target > 0 ? ((actual / target) * 100).toFixed(0) : (actual > 0 ? 100 : 0);

            const progressItem = document.createElement('div');
            progressItem.className = 'progress-item';
            progressItem.innerHTML = `
                <span>${metric}:</span>
                <div class="progress-bar-container-small">
                    <div class="progress-bar ${getProgressColorClass(actual, target)}" style="width: ${Math.min(100, percentage)}%;">
                        ${percentage}%
                    </div>
                </div>
                <span class="actual-target-text">(${actual}/${target})</span>
            `;
            progressList.appendChild(progressItem);
        });
        progressSection.appendChild(progressList);
        summaryCard.appendChild(progressSection);

        // Section 3: Product Interest Breakdown (Bottom)
        const productInterestSection = document.createElement('div');
        productInterestSection.className = 'summary-section product-interest-breakdown';
        productInterestSection.innerHTML = `
            <h3>Products of Interest</h3>
            <p>${productInterests.length > 0 ? productInterests.join(', ') : 'No specific product interests recorded this month.'}</p>
        `;
        summaryCard.appendChild(productInterestSection); // This will naturally go to the next row in the grid

        reportDisplay.appendChild(summaryCard);
    }

    // Render All Entries (Employee)
    function renderAllEntriesReport() {
        const selectedEmployeeCode = employeeSelect.value;
        const employeeName = employeeCodeToNameMap[selectedEmployeeCode] || selectedEmployeeCode;
        reportDisplay.innerHTML = `<h2>All Entries for ${employeeName}</h2>`;

        if (selectedEmployeeCodeEntries.length === 0) {
            reportDisplay.innerHTML += `<p class="no-participation-message">No entries found for ${employeeName}.</p>`;
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';
        const table = document.createElement('table');
        table.className = 'all-entries-table';

        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const headers = [
            HEADER_DATE, HEADER_BRANCH_NAME, HEADER_ACTIVITY_TYPE, HEADER_TYPE_OF_CUSTOMER,
            HEADER_PROSPECT_NAME, HEADER_PHONE_NUMBER_WHATSAPP, HEADER_PRODUCT_INTERESTED,
            HEADER_NEXT_FOLLOW_UP_DATE, HEADER_REMARKS
        ]; // Simplified headers for clarity in this view

        headers.forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        selectedEmployeeCodeEntries.forEach(entry => {
            const row = tbody.insertRow();
            row.insertCell().setAttribute('data-label', HEADER_DATE);
            row.lastChild.textContent = formatDate(entry[HEADER_DATE]);
            row.insertCell().setAttribute('data-label', HEADER_BRANCH_NAME);
            row.lastChild.textContent = entry[HEADER_BRANCH_NAME];
            row.insertCell().setAttribute('data-label', HEADER_ACTIVITY_TYPE);
            row.lastChild.textContent = entry[HEADER_ACTIVITY_TYPE];
            row.insertCell().setAttribute('data-label', HEADER_TYPE_OF_CUSTOMER);
            row.lastChild.textContent = entry[HEADER_TYPE_OF_CUSTOMER];
            row.insertCell().setAttribute('data-label', HEADER_PROSPECT_NAME);
            row.lastChild.textContent = entry[HEADER_PROSPECT_NAME];
            row.insertCell().setAttribute('data-label', HEADER_PHONE_NUMBER_WHATSAPP);
            row.lastChild.textContent = entry[HEADER_PHONE_NUMBER_WHATSAPP];
            row.insertCell().setAttribute('data-label', HEADER_PRODUCT_INTERESTED);
            row.lastChild.textContent = entry[HEADER_PRODUCT_INTERESTED];
            row.insertCell().setAttribute('data-label', HEADER_NEXT_FOLLOW_UP_DATE);
            row.lastChild.textContent = formatDate(entry[HEADER_NEXT_FOLLOW_UP_DATE]);
            row.insertCell().setAttribute('data-label', HEADER_REMARKS);
            row.lastChild.textContent = entry[HEADER_REMARKS];
        });
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Render Employee Performance Report (Monthly Breakdown)
    function renderEmployeePerformanceReport() {
        const selectedEmployeeCode = employeeSelect.value;
        const employeeName = employeeCodeToNameMap[selectedEmployeeCode] || selectedEmployeeCode;
        reportDisplay.innerHTML = `<h2>${employeeName}'s Monthly Performance Breakdown</h2>`;

        // Group entries by month
        const monthlyData = {};
        selectedEmployeeCodeEntries.forEach(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            const monthYear = `${entryDate.getFullYear()}-${String(entryDate.getMonth() + 1).padStart(2, '0')}`;
            if (!monthlyData[monthYear]) {
                monthlyData[monthYear] = [];
            }
            monthlyData[monthYear].push(entry);
        });

        const sortedMonths = Object.keys(monthlyData).sort((a, b) => new Date(a).getTime() - new Date(b).getTime());

        if (sortedMonths.length === 0) {
            reportDisplay.innerHTML += `<p class="no-participation-message">No monthly performance data available for ${employeeName}.</p>`;
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';
        const table = document.createElement('table');
        table.className = 'performance-table';

        const thead = table.createTHead();
        let headerRow = thead.insertRow(); // Main Headers
        headerRow.insertCell().textContent = 'Month';
        headerRow.insertCell().textContent = 'Designation';

        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
        metrics.forEach(metric => {
            const th = document.createElement('th');
            th.colSpan = 3; // 'Actual', 'Target', '%'
            th.textContent = metric;
            headerRow.appendChild(th);
        });

        headerRow = thead.insertRow(); // Sub-headers
        headerRow.insertCell(); // Empty for Month
        headerRow.insertCell(); // Empty for Designation
        metrics.forEach(() => {
            ['Act', 'Tgt', '%'].forEach(subHeader => {
                const th = document.createElement('th');
                th.textContent = subHeader;
                headerRow.appendChild(th);
            });
        });

        const tbody = table.createTBody();
        sortedMonths.forEach(monthYear => {
            const monthEntries = monthlyData[monthYear];
            const { totalActivity } = calculateTotalActivity(monthEntries);
            const designation = employeeCodeToDesignationMap[selectedEmployeeCode] || 'Default';
            const targets = TARGETS[designation] || TARGETS['Default'];

            const row = tbody.insertRow();
            row.insertCell().textContent = monthYear;
            row.insertCell().textContent = designation;

            metrics.forEach(metric => {
                const actual = totalActivity[metric];
                const target = targets[metric];
                const percentage = target > 0 ? ((actual / target) * 100).toFixed(0) : (actual > 0 ? 100 : 0);

                row.insertCell().textContent = actual;
                row.insertCell().textContent = target;
                
                const percentCell = row.insertCell();
                const progressBarContainer = document.createElement('div');
                progressBarContainer.className = 'progress-bar-container-small';
                
                const progressBar = document.createElement('div');
                progressBar.className = 'progress-bar';
                progressBar.style.width = `${Math.min(100, percentage)}%`;
                progressBar.textContent = `${percentage}%`;

                if (percentage >= 100) {
                    progressBar.classList.add('success');
                } else if (percentage >= 75) {
                    progressBar.classList.add('warning-high');
                } else if (percentage >= 50) {
                    progressBar.classList.add('warning-medium');
                } else if (percentage > 0) {
                    progressBar.classList.add('warning-low');
                } else {
                    progressBar.classList.add('no-activity');
                    progressBar.textContent = '0%';
                }
                progressBarContainer.appendChild(progressBar);
                percentCell.appendChild(progressBarContainer);
            });
        });
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }


    // Function to show/hide sections based on tab clicked
    function showTab(tabButtonId) {
        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Hide all main content sections
        reportsSection.style.display = 'none';
        detailedCustomerViewSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        // Clear previous report display content
        reportDisplay.innerHTML = '';
        branchSelect.value = ''; // Reset branch selection
        employeeSelect.value = ''; // Reset employee selection
        employeeFilterPanel.style.display = 'none'; // Hide employee filter
        viewOptions.style.display = 'none'; // Hide view options
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active')); // Deactivate report buttons

        // Activate the clicked tab button and show relevant section
        const clickedButton = document.getElementById(tabButtonId);
        if (clickedButton) {
            clickedButton.classList.add('active');
            if (tabButtonId === 'allBranchSnapshotTabBtn' || tabButtonId === 'allStaffOverallPerformanceTabBtn' || tabButtonId === 'nonParticipatingBranchesTabBtn' || tabButtonId === 'nonParticipatingEmployeesTabBtn') {
                reportsSection.style.display = 'block';
                // Render initial report for the tab
                if (tabButtonId === 'allBranchSnapshotTabBtn') {
                    renderAllBranchSnapshot();
                } else if (tabButtonId === 'allStaffOverallPerformanceTabBtn') {
                    renderOverallStaffPerformanceReport();
                } else if (tabButtonId === 'nonParticipatingBranchesTabBtn') {
                    renderNonParticipatingBranches();
                } else if (tabButtonId === 'nonParticipatingEmployeesTabBtn') {
                    renderNonParticipatingEmployees();
                }
            } else if (tabButtonId === 'detailedCustomerViewTabBtn') {
                detailedCustomerViewSection.style.display = 'block';
                // Reset customer view dropdowns
                customerViewBranchSelect.value = "";
                customerViewEmployeeSelect.value = "";
                customerCanvassedList.innerHTML = '<p>Select a branch and employee to see customers.</p>';
                customerDetailsContent.querySelectorAll('[data-field]').forEach(el => el.textContent = '');
                customerDetailsContent.querySelector('.profile-text').textContent = 'Select a customer to view their detailed profile.';
                customerDetailsContent.querySelector('.remark-text').textContent = 'Select a customer to view their remarks.';
            } else if (tabButtonId === 'employeeManagementTabBtn') {
                employeeManagementSection.style.display = 'block';
                displayEmployeeManagementMessage('Use the forms below to manage employees.', false); // Info message
            }
        }
    }


    // Event Listeners for Tab Buttons
    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    nonParticipatingEmployeesTabBtn.addEventListener('click', () => showTab('nonParticipatingEmployeesTabBtn')); // NEW: Event listener for new tab
    detailedCustomerViewTabBtn.addEventListener('click', () => showTab('detailedCustomerViewTabBtn'));
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));


    // Event Listeners for Report View Options
    viewBranchPerformanceReportBtn.addEventListener('click', (event) => {
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        event.target.classList.add('active');
        renderBranchPerformanceReport();
    });

    viewEmployeeSummaryBtn.addEventListener('click', (event) => {
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        event.target.classList.add('active');
        renderEmployeeSummary(selectedEmployeeCodeEntries);
    });

    viewAllEntriesBtn.addEventListener('click', (event) => {
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        event.target.classList.add('active');
        renderAllEntriesReport();
    });

    viewPerformanceReportBtn.addEventListener('click', (event) => {
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        event.target.classList.add('active');
        renderEmployeePerformanceReport();
    });

    // NEW: Detailed Customer View Logic
    customerViewBranchSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        customerCanvassedList.innerHTML = '<p>Select an employee to see customers.</p>';
        customerDetailsContent.querySelectorAll('[data-field]').forEach(el => el.textContent = '');
        customerDetailsContent.querySelector('.profile-text').textContent = 'Select a customer to view their detailed profile.';
        customerDetailsContent.querySelector('.remark-text').textContent = 'Select a customer to view their remarks.';

        if (selectedBranch) {
            const employeeCodesInBranchFromCanvassing = allCanvassingData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]);
            
            const sortedEmployeeCodesInBranch = [...new Set(employeeCodesInBranchFromCanvassing)].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });
            populateDropdown(customerViewEmployeeSelect, sortedEmployeeCodesInBranch, true);
            customerViewEmployeeSelect.value = ""; // Reset employee selection
        } else {
            customerViewEmployeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';
        }
    });

    customerViewEmployeeSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        const selectedEmployeeCode = customerViewEmployeeSelect.value;
        customerCanvassedList.innerHTML = '';
        customerDetailsContent.querySelectorAll('[data-field]').forEach(el => el.textContent = '');
        customerDetailsContent.querySelector('.profile-text').textContent = 'Select a customer to view their detailed profile.';
        customerDetailsContent.querySelector('.remark-text').textContent = 'Select a customer to view their remarks.';

        if (selectedBranch && selectedEmployeeCode) {
            const customersCanvassed = allCanvassingData.filter(entry =>
                entry[HEADER_BRANCH_NAME] === selectedBranch &&
                entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode &&
                entry[HEADER_PROSPECT_NAME] // Ensure prospect name exists
            );

            if (customersCanvassed.length > 0) {
                // Sort by prospect name for consistent display
                customersCanvassed.sort((a, b) => a[HEADER_PROSPECT_NAME].localeCompare(b[HEADER_PROSPECT_NAME]));

                customersCanvassed.forEach((customer, index) => {
                    const li = document.createElement('li');
                    li.className = 'customer-list-item';
                    li.textContent = customer[HEADER_PROSPECT_NAME];
                    li.dataset.index = index; // Store index to retrieve full details later
                    li.addEventListener('click', () => {
                        // Remove active class from all list items
                        document.querySelectorAll('.customer-list-item').forEach(item => item.classList.remove('active'));
                        // Add active class to the clicked item
                        li.classList.add('active');
                        displayCustomerDetails(customer);
                    });
                    customerCanvassedList.appendChild(li);
                });
            } else {
                customerCanvassedList.innerHTML = '<p class="no-participation-message">No customers canvassed by this employee in this branch.</p>';
            }
        }
    });

    function displayCustomerDetails(customer) {
        // Clear previous details
        customerDetailsContent.querySelectorAll('[data-field]').forEach(el => el.textContent = '');

        // Populate Customer Overview
        customerDetailsContent.querySelector('[data-field="customerName"]').textContent = customer[HEADER_PROSPECT_NAME] || 'N/A';
        customerDetailsContent.querySelector('[data-field="customerAge"]').textContent = customer[HEADER_DOB_WD] || 'N/A'; // Assuming DOB/WD is for age if no explicit age field
        customerDetailsContent.querySelector('[data-field="customerProfession"]').textContent = customer[HEADER_PROFESSION] || 'N/A';
        customerDetailsContent.querySelector('[data-field="customerMobile"]').textContent = customer[HEADER_PHONE_NUMBER_WHATSAPP] || 'N/A';
        customerDetailsContent.querySelector('[data-field="customerAddress"]').textContent = customer[HEADER_ADDRESS] || 'N/A';

        // Populate Canvassing Activity
        customerDetailsContent.querySelector('[data-field="canvassedDate"]').textContent = formatDate(customer[HEADER_DATE]) || 'N/A';
        customerDetailsContent.querySelector('[data-field="canvassedBy"]').textContent = employeeCodeToNameMap[customer[HEADER_EMPLOYEE_CODE]] || customer[HEADER_EMPLOYEE_CODE] || 'N/A';
        customerDetailsContent.querySelector('[data-field="canvassingMedium"]').textContent = customer[HEADER_HOW_CONTACTED] || 'N/A'; // Assuming How Contacted is canvassing medium
        customerDetailsContent.querySelector('[data-field="currentStatus"]').textContent = customer[HEADER_TYPE_OF_CUSTOMER] || 'N/A'; // Assuming Type of Customer reflects current status
        customerDetailsContent.querySelector('[data-field="followUpDate"]').textContent = formatDate(customer[HEADER_NEXT_FOLLOW_UP_DATE]) || 'N/A';
        customerDetailsContent.querySelector('[data-field="productInterest"]').textContent = customer[HEADER_PRODUCT_INTERESTED] || 'N/A';

        // Populate Family Details
        customerDetailsContent.querySelector('[data-field="maritalStatus"]').textContent = customer[HEADER_RELATION_WITH_STAFF] || 'N/A'; // Re-purposing for marital status if applicable
        customerDetailsContent.querySelector('[data-field="spouseName"]').textContent = customer[HEADER_FAMILY_DETAILS_1] || 'N/A';
        customerDetailsContent.querySelector('[data-field="children"]').textContent = customer[HEADER_FAMILY_DETAILS_3] || 'N/A'; // Concatenate children details if needed
        customerDetailsContent.querySelector('[data-field="familyIncome"]').textContent = customer[HEADER_FAMILY_DETAILS_2] || 'N/A'; // Re-purposing for family income if applicable

        // Profile and Remarks
        customerDetailsContent.querySelector('[data-field="customerProfile"]').textContent = customer[HEADER_PROFILE_OF_CUSTOMER] || 'No detailed profile available.';
        customerDetailsContent.querySelector('[data-field="remarks"]').textContent = customer[HEADER_REMARKS] || 'No remarks available.';
    }


    // Google Apps Script Web App Communication
    async function sendDataToGoogleAppsScript(action, data) {
        displayEmployeeManagementMessage('Processing request...', false);
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
                throw new Error(`Server responded with an error: ${response.status} - ${errorText}`);
            }

            const result = await response.json();
            if (result.status === 'success') {
                displayEmployeeManagementMessage(result.message, false);
                // Re-process data to update dropdowns and reports after successful modification
                await processData(); 
                // After successful add/delete, if on employee management tab, clear forms
                if (employeeManagementSection.style.display === 'block') {
                    if (action === 'add_employee') {
                        addEmployeeForm.reset();
                    } else if (action === 'add_bulk_employees') {
                        bulkAddEmployeeForm.reset();
                    } else if (action === 'delete_employee') {
                        deleteEmployeeForm.reset();
                    }
                }
                return true;
            } else {
                displayEmployeeManagementMessage(`Error: ${result.message}`, true);
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Apps Script:', error);
            displayEmployeeManagementMessage(`Failed to send data: ${error.message}. Check network and Apps Script deployment.`, true);
            return false;
        }
    }


    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const newEmployeeName = newEmployeeNameInput.value.trim();
            const newEmployeeCode = newEmployeeCodeInput.value.trim();
            const newBranchName = newBranchNameInput.value.trim();
            const newDesignation = newDesignationInput.value.trim();

            if (!newEmployeeName || !newEmployeeCode || !newBranchName) {
                displayEmployeeManagementMessage('Name, Code, and Branch Name are required.', true);
                return;
            }

            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: newEmployeeName,
                [HEADER_EMPLOYEE_CODE]: newEmployeeCode,
                [HEADER_BRANCH_NAME]: newBranchName,
                [HEADER_DESIGNATION]: newDesignation
            };

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);
            if (success) {
                addEmployeeForm.reset(); // Clear form on success
            }
        });
    }

    // Event Listener for Bulk Add Employees Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const bulkDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName || !bulkDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk entry.', true);
                return;
            }

            const lines = bulkDetails.split('\n').filter(line => line.trim() !== '');
            const employeesToAdd = [];

            for (const line of lines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length >= 2) { // Minimum Name, Code
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
