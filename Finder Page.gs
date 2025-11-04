// NEW: Define emails that have admin access.
// These users will see ALL data, bypassing the date and other filters.
// IMPORTANT: Every email listed here MUST also have an entry in the 'validUsers' object below.
const SUPER_USERS = [
  "jecastro@megaworld-lifestyle.com",
  "jmpizarro@megaworld-lifestyle.com",
  "rrlandoy@megaworld-lifestyle.com",
  "mppiamonte@megaworld-lifestyle.com"
  // Add other admin emails here, separated by commas
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle("MCD Memo Finder");
}

let cachedSheetData = null;

/**
 * Fetches and processes data from the spreadsheet.
 * Includes a check to see if the user has special access permissions.
 * @param {number[]} selectedColumns Array of column numbers to fetch.
 * @param {number} startRow The row number to start fetching data from.
 * @param {string} userEmail The email of the user to check against the SUPER_USERS list.
 * @returns {Array<Array<string>>} A 2D array of the processed data.
 */
function getSheetData(selectedColumns, startRow, userEmail) {
  // Check if the current user has superuser privileges
  const isSuperUser = SUPER_USERS.includes(userEmail);
  
  const dateColumnNumber = selectedColumns[1]; 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("importselectedcol");
  const lastRow = sheet.getLastRow();
  startRow = startRow || 1;

  const numRows = lastRow - startRow + 1;
  if (numRows < 1) return [];

  const fullData = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  let dataToProcess;

  // Conditionally filter data based on user role
  if (isSuperUser) {
    // If the user is a superuser, bypass filters and use the full dataset
    Logger.log(`User ${userEmail} is a superuser. Loading all ${fullData.length} rows.`);
    dataToProcess = fullData;
  } else {
    // If a regular user, apply the original filters
    Logger.log(`User ${userEmail} is a regular user. Applying filters.`);
    const threeMonthsAgo = new Date();
    threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);

    dataToProcess = fullData.filter(row => {
      const col8Valid = row[7] !== undefined && row[7] !== null && String(row[7]).trim() !== '';
      const dateValue = row[dateColumnNumber - 1]; 
      const dateInColumn = dateValue instanceof Date ? dateValue : null;
      const isDateValid = dateInColumn ? dateInColumn >= threeMonthsAgo : true;
      return col8Valid && isDateValid;
    });
  }

  // Map the selected columns from the (now conditionally filtered) data
  return dataToProcess.map(row => {
    while (row.length < Math.max(...selectedColumns)) {
      row.push('');
    }
    return selectedColumns.map(colIndex => {
      const cell = row[colIndex - 1];
      if (colIndex === dateColumnNumber && cell instanceof Date) {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), "MMM d, yyyy H:mm:ss");
      }
      return cell;
    });
  });
}

function clearSheetDataCache() {
  cachedSheetData = null;
}

/**
 * MODIFIED: Validates user credentials and now also returns their role.
 */
function validateLogin(username, password) {
  const validUsers = {
  "mmcontractsreceiptstorage@megaworld-lifestyle.com":{password:"#approvena",actualName:"Jeron Luther"},
"jbarante@megaworld-lifestyle.com":{password:"JBBxxx1989",actualName:"Judylin Bondoc-Arante"},
"reginearcovia@gmail.com":{password:"JULY91999",actualName:"REGINE OPENA"},
"srnacional@megaworld-lifestyle.com":{password:"shai1234",actualName:"Shaira Danica Nacional"},
"fplayug@megaworld-lifestyle.com":{password:"flor0802",actualName:"FLORYTESS P. LAYUG"},
"jdfulgar@megaworld-lifestyle.com":{password:"061518",actualName:"Judy Ann Fulgar "},
"rpfermaran@megaworld-lifestyle.com":{password:"10091999",actualName:"Racquel May Fermaran"},
"llformentera@megaworld-lifestyle.com":{password:"Lf010189",actualName:"LANIBEL FORMENTERA"},
"ppabella@megaworld-lifestyle.com":{password:"858585",actualName:"Praise Jan P. Abella"},
"llformentera@megaworld-lifestyle.com":{password:"Lf010189",actualName:"Lanibel Formentera"},
"ppabella@megaworld-lifestyle.com":{password:"858585",actualName:"Praise Jan P. Abella"},
"llformentera@megaworld-lifestyle.com":{password:"Lf010189",actualName:"Lanibel Formentera"},
"ppabella@megaworld-lifestyle.com":{password:"858585",actualName:"Praise Jan P. Abella"},
"acchua@megaworld-lifestyle.com":{password:"Memopass1234",actualName:"Angeli Marie C. Chua"},
"acchua@megaworld-lifetyle.com":{password:"Memopass1234",actualName:"Angeli Marie C. Chua"},
"rrlandoy@megaworld-lifestyle.com":{password:"Mcd@123",actualName:"Rodalyn Landoy"},
"idmaximo@megaworld-lifestyle.com":{password:"1234",actualName:"Israel Maximo"},
"monocomora@megaworld-lifestyle.com":{password:"ForbesTown2025",actualName:"Melchor Nocomora"},
"monocomora@megaworld-lifestyle.com":{password:"ForbesTown2025",actualName:"Steffi C. Famatiga"},
"monocomora@megaworld-lifestyle.com":{password:"ForbesTown2025",actualName:"Jessie T. Carlos"},
"monocomora@megaworld-lifestyle.com":{password:"ForbesTown2025",actualName:"John Carlo J. Buzeta"},
"monocomora@megaworld-lifestyle.com":{password:"ForbesTown2025",actualName:"Liwayway Y. Molina"},
"monocomora@megaworld-lifestyle.com":{password:"ForbesTown2025",actualName:"Gonzalo D. Tolentino Jr."},
"monocomora@megaworld-lifestyle.com":{password:"ForbesTown2025",actualName:"Elaine N. Manimtim"},
"monocomora@megaworld-lifestyle.com":{password:"ForbesTown2025",actualName:"Joseph Z. Lagunday"},
"jscale@megaworld-lifestyle.com":{password:"013183",actualName:"JENAVE S. CALE"},
"mtdiola@megaworld-lifestyle.com":{password:"TUE@2025",actualName:"MARY GRACE T. DIOLA"},
"jfreynes@megaworld-lifestyle.com":{password:"@memoroutingapp_0214",actualName:"Jeremy F. Reynes"},
"rlborlongan@megaworld-lifestyle.com":{password:"Mcdbilling1",actualName:"RAN BORLONGAN"},
"smcervan@megaworld-lifestyle.com":{password:"chiching17!",actualName:"Shiela May Cervan"},
"smcervan@megaworld-lifestyle.com":{password:"chiching17!",actualName:"Shiela May Cervan"},
"jonulud@megaworld-lifestyle.com":{password:"jmnulud@2022",actualName:"John Michael Nulud"},
"dbpineda@megaworld-lifestyle.com":{password:"Darleen121724",actualName:"Darleen B Pineda"},
"dbpineda@megaworld-lifestyle.com":{password:"Darleen121724",actualName:"Darleen B Pineda"},
"nmmizoguchi@megaworld-lifestyle.com":{password:"@@@@@@",actualName:"Nori Mizoguchi"},
"jrabedes@megaworld-lifestyle.com":{password:"janinescyrelleav",actualName:"Janine Scyrelle Abedes"},
"bvvaldenarro@megaworld-lifestyle.com":{password:"brixton20",actualName:"Jan Brix V. Valdenarro"},
"pllangit@megaworld-lifestyle.com":{password:"LCT2025",actualName:"Paola Camille Langit"},
"rrmagdaluyo@megaworld-lifestyle.com":{password:"080684",actualName:"Ryan Magdaluyo"},
"jdmarcos@megaworld-lifestyle.com":{password:"@Jeff123",actualName:"Jeffry Marcos"},
"mobonales@megaworld-lifestyle.com":{password:"Mb090100*",actualName:"Michaela Bonales"},
"accastillo@megaworld-lifestyle.com":{password:"2015-12-0143",actualName:"ALEXANDER CASTILLO"},
"jjcastillo@megaworld-lifestyle.com":{password:"J@nuary1979",actualName:"Jose Maria Castillo"},
"rfbarce@megaworld-lifestyle.com":{password:"Rovsmemo2025",actualName:"ROVELIN F. BARCE"},
"dcpo@megaworld-lifestyle.com":{password:"rawrrawrrawr",actualName:"Danica C. Po"},
"cspuyat@megaworld-lifestyle.com":{password:"@CSPccfmemo2025",actualName:"Charisse S. Puyat"},
"zqcordero@megaworld-lifestyle.com":{password:"Ccfdrew@2025",actualName:"Zandro Q. Cordero"},
"jebeltran@megaworld-lifestyle.com":{password:"jamjambeltran",actualName:"JAMIL BELTRAN"},
"metianero@megaworld-lifestyle.com":{password:"mkt050817",actualName:"Mitchelle Keith Tianero"},
"lbcicat@megaworld-lifestyle.com":{password:"eaanne2017",actualName:"Lorraine Ann Cicat"},
"tstan@megaworld-lifestyle.com":{password:"CGMOps@888",actualName:"Tyrone Jason S Tan"},
"jicanilao@megaworld-lifestyle.com":{password:"ballpen0.8",actualName:"Joel I. Canilao"},
"etmayuga@megaworld-lifestyle.com":{password:"etmayuga@megaworld",actualName:"Elwin James T. Mayuga"},
"magregorio@megaworld-lifestyle.com":{password:"KaluskusngKawayan",actualName:"MHIKE A. GREGORIO"},
"rbalmorfe@megaworld-lifestyle.com":{password:"Newport@2022",actualName:"RIZZA B. ALMORFE"},
"lbcicat@megaworld-lifestyle.com":{password:"earaine2017",actualName:"Lorraine Ann B. Cicat"},
"plarmada@megaworld-lifestyle.com":{password:"Ccf@2025",actualName:"Paul Christian L. Armada"},
"jdfulgar@megaworld-lifestyle.com":{password:"mallops",actualName:"Judy Ann Fulgar"},
"cmbacay@megaworld-lifestyle.com":{password:"cmmb2509",actualName:"Charmina M. Bacay"},
"rsaguilar@megaworld-lifestyle.com":{password:"ronron042022",actualName:"Ron Mickhael S. Aguilar"},
"lctorres@megaworld-lifestyle.com":{password:"CCF_2025",actualName:"Leonard Torres"},
"mjsumilong@megaworld-lifestyle.com":{password:"Nheloggs_13",actualName:"Marnelle Sumilong"},
"jvcruz@megaworld-lifestyle.com":{password:"pepsi123",actualName:"Josebb Angelo V. Cruz"},
"jvcruz@megaworld-lifestyle.com":{password:"pepsi123",actualName:"Josebb Angelo V. Cruz"},
"bgmiagao@megaworld-lifestyle.com":{password:"b3r202013",actualName:"BRIAN G. MIAG-AO"},
"aclabindalawa@megaworld-lifestyle.com":{password:"Meg0507",actualName:"Aepreha May C. Labindalawa"},
"meanaud@megaworld-lifestyle.com":{password:"Meanaud1234!@#$",actualName:"Mart Loie Estrada Anaud "},
"meanaud@megaworld-lifestyle.com":{password:"Meanaud1234!@#$",actualName:"Mart Loie Estrada Anaud "},
"jmmolo@megaworld-lifestyle.com":{password:"Koniii28",actualName:"John Lester Molo"},
"ssgarcia@megaworld-lifestyle.com":{password:"shingarcia.06",actualName:"Shenna Mae S. Garcia"},
"mlnerie@megaworld-lifestyle.com":{password:"CLUSTER12u!",actualName:"MELODY L. NERIE"},
"lcesguerra@megaworld-lifestyle.com":{password:"Lauricemae0!!",actualName:"Laurice Mae C. Esguerra"},
"gltolentino@megaworld-lifestyle.com":{password:"P@ssw0rd123!",actualName:"Gjamil Zyris L. Tolentino"},
"jmcordovales@megaworld-lifestyle.com":{password:"P@ssw0rd23",actualName:"Janinna M. Cordovales"},
"gltolentino@megaworld-lifestyle.com":{password:"P@ssw0rd123!",actualName:"Gjamil Zyris L. Tolentino"},
"matanedo@megaworld-lifestyle.com":{password:"gh1234",actualName:"Ma. Engelie Tanedo"},
"mamagallanes@megaworld-lifestyle.com":{password:"9756680164Aviso",actualName:"Mariel Magallanes"},

"spmontemayor@megaworld-lifestyle.com":{password:"PASEO 123",actualName:"Stephen Montemayor"},
"spmontemayor@megaworld-lifestyle.com":{password:"SAP123",actualName:"Stephen Montemayor"},
"spmontemayor@megaworld-lifestyle.com":{password:"THREECEN123",actualName:"STEPHEN MONTEMAYOR"},
"spmontemayor@megaworld-lifestyle.com":{password:"SLP123",actualName:"STEPHEN P. MONTEMAYOR"},
"mvcruz@megaworld-lifestyle.com":{password:"Welcome123*",actualName:"Ma. Angelica V. Cruz"},
"hbhispano@megaworld-lifestyle.com":{password:"haneyhispano2025",actualName:"Hanelyn Hispano"},
"mvcruz@megaworld-lifestyle.com":{password:"Welcome123*",actualName:"Ma. Angelica V. Cruz"},
"mvcruz@megaworld-lifestyle.com":{password:"Welcome123*",actualName:"Ma. Angelica V. Cruz"},
"msvillamor@megaworld-lifestyle.com":{password:"jane2025",actualName:"Mary Jane Villamor"},
"gqmaningding@megaworld-lifestyle.com":{password:"090909",actualName:"Gerly Q. Maningding"},
"gmcoates@megaworld-lifestyle.com":{password:"pass1",actualName:"Graham Coates"},
"mglao@megaworld-lifestyle.com":{password:"pass3",actualName:"Michael Lao"},
"mppiamonte@megaworld-lifestyle.com":{password:"pass3a",actualName:"Louise Piamonte"},
"jmpizarro@megaworld-lifestyle.com":{password:"pass6",actualName:"Jaye Trich Pizarro"},
"jcmelitante@megaworld-lifestyle.com":{password:"pass7",actualName:"Jocelyn Melitante"},
"apflores@megaworld-lifestyle.com":{password:"pass8",actualName:"Alex Flores"},
"mdpastolero@megaworld-lifestyle.com":{password:"pass9",actualName:"Aurora Palostero"},
"mbcaleja@megaworld-lifestyle.com":{password:"pass10",actualName:"Mariano Caleja"},
"celvina@megaworld-lifestyle.com":{password:"camsy123",actualName:"Camsy Elvina"},
"mvvivo@megaworld-lifestyle.com":{password:"mikee123",actualName:"Mikee Vivo"},
"frpacampara@megaworld-lifestyle.com":{password:"fermie123",actualName:"Fermila Pacampara"},
"vrvicente@megaworld-lifestyle.com":{password:"vanessa123",actualName:"Vanessa Vicente"},
"jecastro@megaworld-lifestyle.com":{password:"test123",actualName:"Jeron Luther Castro"},
"burofornisandrea30@gmail.com":{password:"@Fornis022920",actualName:"Andrea Buro"},
"bdrellis@megaworld-lifestyle.com":{password:"bdrellis123",actualName:"Brisbane Rellis"},
"pbaclan@megaworld-lifestyle.com":{password:"FTC1234",actualName:"Prescila Aclan"},
"faignas@megaworld-lifestyle.com":{password:"Leasing2023",actualName:"Frencie Ignas"},

    
  };

  const correctedUsername = username.replace(/lifetyle\.com|lifestye\.com/, "lifestyle.com");

  if (validUsers[correctedUsername] && validUsers[correctedUsername].password === password) {
    // NEW: Check if the validated user is in the SUPER_USERS list
    const isSuperUser = SUPER_USERS.includes(correctedUsername);
    return {
      valid: true,
      actualName: validUsers[correctedUsername].actualName,
      isSuperUser: isSuperUser // Add this flag to the result
    };
  } else {
    return {
      valid: false
    };
  }
}
