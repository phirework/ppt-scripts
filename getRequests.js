const UTC_OFFSET = 4;
const HR_MS = 60 * 60 * 1000;
const NA = "DATA_NOT_FOUND";

const REQUESTS_SHEET = "1mvH_VWwClOzIakdbBTkbMb2fZNqcU_OBXCYWAwU3Mjo";
const GROCERIES_TAB = "Groceries - PENDING";
const MEALS_TAB = "Meals - PENDING";

const CHEF_SHEET = "1xI966-4oWt84ZiU_46I5fYDXQ7wscua1_pVLf9wDDMs";
const CHEF_SIGNUP_TAB = "MEALS SIGN-UP";
const RECURRING_TAB = "RECURRENT MEALS";

const VOLUNTEER_SHEET = "1Xf1p_o7SMdbKxEJWR_TvjMJG1qdRZRUtnw_6w2tyhGg";
const CHEF_INFO_TAB = "New Chefs list";

let dateRange = getDateRange();

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('People\'s Pantry')
      .addItem('Get Request Data', 'getRequestData')
      .addToUi();
}

const getRequestData = () => {
  const LOGISTICS_TAB = "Testing automation";
  let logisticsData = [];

  // requests data
  const mealsData = transformData(filterByDate(retrieveData(REQUESTS_SHEET, MEALS_TAB), "D"), "B");
  const groceriesData = transformData(filterByDate(retrieveData(REQUESTS_SHEET, GROCERIES_TAB), "D"), "B");
  const requests = {...mealsData}; // TODO - leaving off groceries for now

  // chef data
  const recurrentData = transformData(filterByDate(retrieveData(CHEF_SHEET, RECURRING_TAB), "N"), "B");
  const chefData = transformData(filterByDate(retrieveData(CHEF_SHEET, CHEF_SIGNUP_TAB), "O"), "B");

  // volunteer data
  const volunteerData = transformData(retrieveData(VOLUNTEER_SHEET, CHEF_INFO_TAB), "C");

  Object.keys(requests).map((id) => {
    const requester = requests[id];
    const delIntersection = requester[getColIndex("S")];
    const delZone = requester[getColIndex("Q")];
    const delNotes = requester[getColIndex("AO")];

    let pickIntersection = NA
        pickContact = NA,
        pickTime = NA;

    if (chefData[id]) {
      const supplier = chefData[id];
      pickContact = supplier[getColIndex("K")].trim();
      pickTime = supplier[getColIndex("P")];
    } else if (recurrentData[id]) {
      const supplier = recurrentData[id];
      pickContact = supplier[getColIndex("J")].trim();
      pickTime = supplier[getColIndex("O")];
    }

    if (volunteerData[pickContact]) {
      const volunteer = volunteerData[pickContact];
      pickIntersection = volunteer[getColIndex("G")];
    }

    const row = [id, dateRange.start, truncateName(pickContact), pickIntersection, getMapLink(pickIntersection), pickTime, '', '', '', delIntersection, delZone, getMapLink(delIntersection), delNotes];
    logisticsData.push(row);
  });

  const activeRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOGISTICS_TAB).getRange(`A2:M${logisticsData.length + 1}`);
  activeRange.setValues(logisticsData);
}

/*
 * Returns two-dimensional array of spreadsheet data in rows
 *
 * @params
 *   {string} sheetId - the alphanumeric ID that identifies a Google Sheet
 *   {string} tabName - the name of the tab you're retrieving data from
 */
const retrieveData = (sheetId, tabName) => {
  return SpreadsheetApp.openById(sheetId).getSheetByName(tabName).getDataRange().getValues();
}


/*
 * Returns two-dimensional array of data filtered by date range
 *
 * @params
 *   {Array[][]} dataArray - two dimensional array of spreadsheet data
 *   {string} dateCol - a string representing the column to filter dateRange by
 */
const filterByDate = (dataArray, dateCol) => {
  const dateColNum = getColIndex(dateCol);

  const filtered = dataArray.filter((row) => {
    const requestTime = new Date(row[dateColNum]);
    const valid = requestTime >= dateRange.start && requestTime < dateRange.end;
    return valid;
  });

  return filtered;
}


/*
 * Returns object in the form of { [identifier] : [row values] }, filtered by date
 *
 * @params
 *   {Array[][]} dataArray - two dimensional array of spreadsheet data
 *   {string} idCol - a string representing the column to use as object key
 */

const transformData = (dataArray, idCol) => {
  const transformedObj = {};

  dataArray.map((row) => {
    const key = row[getColIndex(idCol)].trim();
    transformedObj[key] = row;
  });

  return transformedObj;
}
/*
 * Returns index value of column name, starting at 0
 *
 * @params
 *   {string} colName - the alphanumeric index name of column, e.g. "AA"
 */
const getColIndex = (colName) => {
  let index = 0;
  let cols = [];

  colName.toLowerCase().split('').map((char) => {
    cols.push(char.charCodeAt(0) - 96);
  });

  cols.reverse().map((colIndex, i) => {
    index += Math.pow(26, i) * colIndex;
  });

  return index - 1;
}

/*
 * Returns HTML for link element given intersection
 *
 * @params
 *   {string} intersection - the intersection in Toronto to find
 */
const getMapLink = (intersection) => {
  const encodedAddr = `"https://maps.google.com?q=${escape(intersection)},+Toronto"`;
  return `=hyperlink(${encodedAddr}, "View on Google Maps")`;
}

const truncateName = (name) => {
  const parts = name.split(" ");
  return parts.length > 1 ? `${parts[0]} ${parts[1][0]}` : name;
}

/*
 * Returns { start: Date, end: Date } object based on today's date in EDT
 *
 */
function getDateRange() {
  // Working with UTC time to prevent weird client-side time from screwing this up
  // To get correct date value need to make sure UTC is on the same day as EDT
  const current = new Date(Date.now() - (UTC_OFFSET * HR_MS));
  const start = current.setUTCHours(UTC_OFFSET, 0, 0, 0) + 24 * HR_MS;
  const end = new Date(start).getTime() + 24 * HR_MS;

  return { start: new Date(start), end: new Date(end) };
}