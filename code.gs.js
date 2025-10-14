// Compiled using undefined undefined (TypeScript 4.9.5)
// 請將 'YOUR_SPREADSHEET_ID_HERE' 替換成您的 Google Sheet 檔案 ID
var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
function onEdit(e) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(30000);
        if (e.range.getFormula().toUpperCase() == "=MY_OBJECT_NUMBER()") {
            var activeSheet = e.source.getActiveSheet();
            var objectType = activeSheet.getName().toUpperCase();
            e.range.setValue(createObjectNumber(objectType));
        }
    }
    catch (e) {
    }
    finally {
        lock.releaseLock();
    }
}
function doGet(request) {
    var path = request === null || request === void 0 ? void 0 : request.pathInfo;
    switch (path) {
        case 'index':
        default:
            var template = HtmlService.createTemplateFromFile('index');
            return template.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    }
}
function showObjectInfo(objectType, objectNumber) {
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            var template = HtmlService.createTemplateFromFile('buildingInfo');
            var dataString = searchObjectInfo(objectType, objectNumber);
            var buildingObject = JSON.parse(dataString);
            template.buildingObject = buildingObject;
            console.log(JSON.stringify(buildingObject));
            return template.evaluate().getContent();
        case 'LAND':
            var landTemplate = HtmlService.createTemplateFromFile('landInfo');
            var landDataString = searchObjectInfo(objectType, objectNumber);
            var landObject = JSON.parse(landDataString);
            landTemplate.landObject = landObject;
            console.log(JSON.stringify(landObject));
            return landTemplate.evaluate().getContent();
    }
    return "";
}
function showObjectA4Info(objectType, objectNumber) {
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            // const buildingTemplate = HtmlService.createTemplateFromFile('buildingA4')
            var dataString = searchObjectInfo(objectType, objectNumber);
            var buildingObject = JSON.parse(dataString);
            // buildingTemplate.buildingObject = buildingObject
            // console.log(JSON.stringify(buildingObject))
            return createContract(objectType, buildingObject);
        // return buildingTemplate.evaluate()
        case 'LAND':
            // const landTemplate = HtmlService.createTemplateFromFile('landA4')
            var landDataString = searchObjectInfo(objectType, objectNumber);
            var landObject = JSON.parse(landDataString);
            // landTemplate.landObject = landObject
            // console.log(JSON.stringify(landObject))
            return createContract(objectType, landObject);
        // return landTemplate.evaluate()
    }
    return "";
}
function searchObjectInfo(objectType, objectNumber) {
    // 1. 快取機制
    var cache = CacheService.getScriptCache();
    var cacheKey = "info_".concat(objectType, "_").concat(objectNumber);
    var cached = cache.get(cacheKey);
    if (cached) {
        console.log("Serving info from cache for ".concat(objectNumber));
        return cached;
    }
    console.log("Fetching info from sheet for ".concat(objectNumber));
    // 2. 使用 GQL 查詢
    var sheetName = objectType;
    var headersEnum = (sheetName.toUpperCase() === 'BUILDING') ? BuildingHeaders : LnadHeaders;
    var objectNumberCol = getColumnLetter(headersEnum.OBJECT_NUMBER);
    var query = "SELECT * WHERE ".concat(objectNumberCol, " = '").concat(objectNumber, "'");
    var url = "https://docs.google.com/spreadsheets/d/".concat(SPREADSHEET_ID, "/gviz/tq?sheet=").concat(sheetName, "&tq=").concat(encodeURIComponent(query));
    var response;
    try {
        response = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer ".concat(ScriptApp.getOAuthToken()) } });
    }
    catch (e) {
        console.error("GQL Fetch Error for ".concat(objectNumber, ": ").concat(e));
        return ""; // 發生錯誤時回傳空字串
    }
    var text = response.getContentText();
    var json = JSON.parse(text.substring(text.indexOf('(') + 1, text.lastIndexOf(')')));
    if (!json.table || !json.table.rows || json.table.rows.length === 0) {
        console.log("Object not found: ".concat(objectNumber));
        return ""; // 找不到物件
    }
    var row = json.table.rows[0].c.map(function (cell) { return cell ? (cell.f || cell.v) : null; });
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            var buildingObject = {
                createTime: row[BuildingHeaders.CREATE_TIME],
                objectNumber: row[BuildingHeaders.OBJECT_NUMBER],
                objectName: row[BuildingHeaders.OBJECT_NAME],
                contractType: row[BuildingHeaders.CONTRACT_TYPE],
                location: row[BuildingHeaders.LOCATION],
                buildingType: row[BuildingHeaders.BUILDING_TYPE],
                housePattern: row[BuildingHeaders.HOUSE_PATTERN],
                floor: row[BuildingHeaders.FLOOR],
                address: row[BuildingHeaders.ADDRESS],
                position: row[BuildingHeaders.POSITION],
                valuation: row[BuildingHeaders.VALUATION],
                landSize: row[BuildingHeaders.LAND_SIZE],
                buildingSize: row[BuildingHeaders.BUILDING_SIZE],
                direction: row[BuildingHeaders.DIRECTION],
                vihecleParkingType: row[BuildingHeaders.VIHECLE_PARKING_TYPE],
                vihecleParkingNumber: row[BuildingHeaders.VIHECLE_PARKING_NUMBER],
                waterSupply: row[BuildingHeaders.WATER_SUPPLY],
                roadNearby: row[BuildingHeaders.ROAD_NEARBY],
                width: row[BuildingHeaders.WIDTH],
                buildingAge: row[BuildingHeaders.BUILDING_AGE],
                memo: row[BuildingHeaders.MEMO],
                contactPerson: row[BuildingHeaders.CONTACT_PERSON],
                pictureLink: row[BuildingHeaders.PICTURE_LINK],
                contractDateFrom: formatDateString(row[LnadHeaders.CONTRACT_DATE_FROM]),
                contractDateTo: formatDateString(row[LnadHeaders.CONTRACT_DATE_TO])
            };
            var resultString_1 = JSON.stringify(buildingObject);
            cache.put(cacheKey, resultString_1, 10800); // 存入快取，3小時
            return resultString_1;
        // const template = HtmlService.createTemplateFromFile('buildingInfo')
        // template.buildingObject = buildingObject
        // console.log(JSON.stringify(buildingObject))
        // return template.evaluate().getContent()
        case 'LAND':
            var landObject = {
                createTime: row[LnadHeaders.CREATE_TIME],
                objectNumber: row[LnadHeaders.OBJECT_NUMBER],
                objectName: row[LnadHeaders.OBJECT_NAME],
                contractType: row[LnadHeaders.CONTRACT_TYPE],
                location: row[LnadHeaders.LOCATION],
                landPattern: row[LnadHeaders.LAND_PATTERN],
                landUsage: row[LnadHeaders.LNAD_USAGE],
                landType: row[LnadHeaders.LNAD_TYPE],
                address: row[LnadHeaders.ADDRESS],
                position: row[LnadHeaders.POSITION],
                valuation: row[LnadHeaders.VALUATION],
                landSize: row[LnadHeaders.LAND_SIZE],
                numberOfOwner: row[LnadHeaders.NUMBER_OF_OWNER],
                roadNearby: row[LnadHeaders.ROAD_NEARBY],
                direction: row[LnadHeaders.DIRECTION],
                waterElectricitySupply: row[LnadHeaders.WATER_ELECTRICITY_SUPPLY],
                width: row[LnadHeaders.WIDTH],
                depth: row[LnadHeaders.DEEPTH],
                buildingCoverageRate: row[LnadHeaders.BUILDING_COVERAGE_RATE],
                volumeRate: row[LnadHeaders.VOLUME_RATE],
                memo: row[LnadHeaders.MEMO],
                contactPerson: row[LnadHeaders.CONTACT_PERSON],
                pictureLink: row[LnadHeaders.PICTURE_LINK],
                contractDateFrom: formatDateString(row[LnadHeaders.CONTRACT_DATE_FROM]),
                contractDateTo: formatDateString(row[LnadHeaders.CONTRACT_DATE_TO])
            };
            var resultString_2 = JSON.stringify(landObject);
            cache.put(cacheKey, resultString_2, 10800); // 存入快取，3小時
            return resultString_2;
        // const landTemplate = HtmlService.createTemplateFromFile('landInfo')
        // landTemplate.landObject = landObject
        // console.log(JSON.stringify(landObject))
        // return landTemplate.evaluate().getContent()
    }
    return "";
}
function formatDateString(date) {
    try {
        return Utilities.formatDate(date, 'GMT+8', 'yyyy/MM/dd');
    }
    catch (error) {
        return "";
    }
}
function createObjectNumber(objectType) {
    var objectNumberPrefix = '';
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            objectNumberPrefix = 'A';
            break;
        case 'LAND':
            objectNumberPrefix = 'B';
            break;
        default:
    }
    return objectNumberPrefix + (searchLastNumOfNumberedObjects(objectType) + 1);
}
function createContract(objectType, data) {
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            return createBuildingContract(data);
        case 'LAND':
            return createLandContract(data);
    }
    return "";
}
function createBuildingContract(data) {
    var googleDocId = '1fE0OZZQ00rcYU38vQWCl4h9kE2oJbHmz5uhb_FtP6Gs'; // google doc ID
    var outputFolderId = '1f-hfkEk0lxp2ha7-hcQ5E3mnsKRfMMUH'; // google drive資料夾ID
    // var googleDocId = '1fBHyUGHH0-hVNq2fTZVXKVxCJ0UYHjkpdOhM1jefQgI'; // 測試google doc ID
    // var outputFolderId = '1lSczRQ0HEKQrcK8PgqHvqD2kLRmRtsrH'; // 測試google drive資料夾ID
    var fileName = "".concat(data.objectName);
    var doc = createDoc(googleDocId, outputFolderId, fileName);
    renderBuildingDoc(doc, data);
    return doc.getUrl();
}
function createLandContract(data) {
    var googleDocId = '1MkGlxmbkGtMayj1ZqHd5y9kIwigZ5ky_ZlwRR1h0hH0'; // google doc ID
    var outputFolderId = '1f-hfkEk0lxp2ha7-hcQ5E3mnsKRfMMUH'; // google drive資料夾ID
    // var googleDocId = '1noZPLBuWEowiDHni3p-6RoafbOV45BylHkdocQ39p0Y'; // 測試google doc ID
    // var outputFolderId = '1lSczRQ0HEKQrcK8PgqHvqD2kLRmRtsrH'; // 測試google drive資料夾ID
    var fileName = "".concat(data.objectName);
    var doc = createDoc(googleDocId, outputFolderId, fileName);
    renderLandDoc(doc, data);
    return doc.getUrl();
}
// 先從樣板合約中複製出一個全新的google doc(this.doc)
function createDoc(googleDocId, outputFolderId, fileName) {
    var file = DriveApp.getFileById(googleDocId);
    var outputFolder = DriveApp.getFolderById(outputFolderId);
    var copy = file.makeCopy(fileName, outputFolder);
    var doc = DocumentApp.openById(copy.getId());
    return doc;
}
function renderBuildingDoc(doc, data) {
    var body = doc.getBody();
    body.replaceText("{{\u7DE8\u865F}}", data.objectNumber);
    body.replaceText("{{\u6848\u540D}}", data.objectName);
    body.replaceText("{{\u5408\u7D04\u985E\u578B}}", data.contractType);
    body.replaceText("{{\u5730\u5340}}", data.location);
    body.replaceText("{{\u5F62\u614B}}", data.buildingType);
    body.replaceText("{{\u683C\u5C40}}", data.housePattern);
    body.replaceText("{{\u6A13\u5C64}}", data.floor.toString());
    body.replaceText("{{\u5730\u5740}}", data.address);
    body.replaceText("{{\u4F4D\u7F6E}}", data.position);
    body.replaceText("{{\u7E3D\u50F9}}", data.valuation.toString());
    body.replaceText("{{\u5730\u576A}}", data.landSize.toString());
    body.replaceText("{{\u5EFA\u576A}}", data.buildingSize.toString());
    body.replaceText("{{\u5EA7\u5411}}", data.direction);
    body.replaceText("{{\u8ECA\u4F4D}}", data.vihecleParkingType);
    body.replaceText("{{\u8ECA\u4F4D\u865F\u78BC}}", data.vihecleParkingNumber.toString());
    body.replaceText("{{\u6C34\u96FB}}", data.waterSupply);
    body.replaceText("{{\u81E8\u8DEF}}", data.roadNearby);
    body.replaceText("{{\u9762\u5BEC}}", data.width.toString());
    body.replaceText("{{\u5B8C\u6210\u65E5}}", data.buildingAge);
    body.replaceText("{{\u5099\u8A3B}}", data.memo);
    body.replaceText("{{\u806F\u7D61\u4EBA}}", data.contactPerson);
    body.replaceText("{{\u5716\u7247\u9023\u7D50}}", data.pictureLink);
    body.replaceText("{{\u5408\u7D04\u958B\u59CB\u65E5\u671F}}", data.contractDateFrom);
    body.replaceText("{{\u5408\u7D04\u7D50\u675F\u65E5\u671F}}", data.contractDateTo);
    doc.saveAndClose();
}
function renderLandDoc(doc, data) {
    var body = doc.getBody();
    body.replaceText("{{\u7DE8\u865F}}", data.objectNumber);
    body.replaceText("{{\u6848\u540D}}", data.objectName);
    body.replaceText("{{\u5408\u7D04\u985E\u578B}}", data.contractType);
    body.replaceText("{{\u5730\u5340}}", data.location);
    body.replaceText("{{\u985E\u5225}}", data.landType);
    body.replaceText("{{\u5206\u5340}}", data.landUsage);
    body.replaceText("{{\u5F62\u614B}}", data.landPattern);
    body.replaceText("{{\u5730\u5740}}", data.address);
    body.replaceText("{{\u4F4D\u7F6E}}", data.position);
    body.replaceText("{{\u7E3D\u50F9}}", data.valuation.toString());
    body.replaceText("{{\u5730\u576A_1}}", data.landSize.toString());
    body.replaceText("{{\u5730\u576A_2}}", (Math.round((data.landSize / 293.4) * 100) / 100).toString());
    body.replaceText("{{\u6240\u6709\u6B0A\u4EBA\u6578}}", data.numberOfOwner.toString());
    body.replaceText("{{\u81E8\u8DEF}}", data.roadNearby);
    body.replaceText("{{\u5EA7\u5411}}", data.direction);
    body.replaceText("{{\u6C34\u96FB}}", data.waterElectricitySupply);
    body.replaceText("{{\u9762\u5BEC}}", data.width.toString());
    body.replaceText("{{\u7E31\u6DF1}}", data.depth.toString());
    body.replaceText("{{\u5EFA\u853D\u7387}}", data.buildingCoverageRate.toString());
    body.replaceText("{{\u5BB9\u7A4D\u7387}}", data.volumeRate.toString());
    body.replaceText("{{\u5099\u8A3B}}", data.memo);
    body.replaceText("{{\u806F\u7D61\u4EBA}}", data.contactPerson);
    body.replaceText("{{\u5716\u7247\u9023\u7D50}}", data.pictureLink);
    body.replaceText("{{\u5408\u7D04\u958B\u59CB\u65E5\u671F}}", data.contractDateFrom);
    body.replaceText("{{\u5408\u7D04\u7D50\u675F\u65E5\u671F}}", data.contractDateTo);
    doc.saveAndClose();
}
var BuildingHeaders;
(function (BuildingHeaders) {
    BuildingHeaders[BuildingHeaders["CREATE_TIME"] = 0] = "CREATE_TIME";
    BuildingHeaders[BuildingHeaders["OBJECT_NUMBER"] = 1] = "OBJECT_NUMBER";
    BuildingHeaders[BuildingHeaders["OBJECT_NAME"] = 2] = "OBJECT_NAME";
    BuildingHeaders[BuildingHeaders["CONTRACT_TYPE"] = 3] = "CONTRACT_TYPE";
    BuildingHeaders[BuildingHeaders["LOCATION"] = 4] = "LOCATION";
    BuildingHeaders[BuildingHeaders["BUILDING_TYPE"] = 5] = "BUILDING_TYPE";
    BuildingHeaders[BuildingHeaders["HOUSE_PATTERN"] = 6] = "HOUSE_PATTERN";
    BuildingHeaders[BuildingHeaders["FLOOR"] = 7] = "FLOOR";
    BuildingHeaders[BuildingHeaders["ADDRESS"] = 8] = "ADDRESS";
    BuildingHeaders[BuildingHeaders["POSITION"] = 9] = "POSITION";
    BuildingHeaders[BuildingHeaders["VALUATION"] = 10] = "VALUATION";
    BuildingHeaders[BuildingHeaders["LAND_SIZE"] = 11] = "LAND_SIZE";
    BuildingHeaders[BuildingHeaders["BUILDING_SIZE"] = 12] = "BUILDING_SIZE";
    BuildingHeaders[BuildingHeaders["DIRECTION"] = 13] = "DIRECTION";
    BuildingHeaders[BuildingHeaders["VIHECLE_PARKING_TYPE"] = 14] = "VIHECLE_PARKING_TYPE";
    BuildingHeaders[BuildingHeaders["VIHECLE_PARKING_NUMBER"] = 15] = "VIHECLE_PARKING_NUMBER";
    BuildingHeaders[BuildingHeaders["WATER_SUPPLY"] = 16] = "WATER_SUPPLY";
    BuildingHeaders[BuildingHeaders["ROAD_NEARBY"] = 17] = "ROAD_NEARBY";
    BuildingHeaders[BuildingHeaders["WIDTH"] = 18] = "WIDTH";
    BuildingHeaders[BuildingHeaders["BUILDING_AGE"] = 19] = "BUILDING_AGE";
    BuildingHeaders[BuildingHeaders["MEMO"] = 20] = "MEMO";
    BuildingHeaders[BuildingHeaders["CONTACT_PERSON"] = 21] = "CONTACT_PERSON";
    BuildingHeaders[BuildingHeaders["PICTURE_LINK"] = 22] = "PICTURE_LINK";
    BuildingHeaders[BuildingHeaders["OBJECT_CREATE_DATE"] = 23] = "OBJECT_CREATE_DATE";
    BuildingHeaders[BuildingHeaders["CONTRACT_DATE_FROM"] = 24] = "CONTRACT_DATE_FROM";
    BuildingHeaders[BuildingHeaders["CONTRACT_DATE_TO"] = 25] = "CONTRACT_DATE_TO";
    BuildingHeaders[BuildingHeaders["OBJECT_UPDATE_DATE"] = 26] = "OBJECT_UPDATE_DATE";
})(BuildingHeaders || (BuildingHeaders = {}));
var LnadHeaders;
(function (LnadHeaders) {
    LnadHeaders[LnadHeaders["CREATE_TIME"] = 0] = "CREATE_TIME";
    LnadHeaders[LnadHeaders["OBJECT_NUMBER"] = 1] = "OBJECT_NUMBER";
    LnadHeaders[LnadHeaders["OBJECT_NAME"] = 2] = "OBJECT_NAME";
    LnadHeaders[LnadHeaders["CONTRACT_TYPE"] = 3] = "CONTRACT_TYPE";
    LnadHeaders[LnadHeaders["LOCATION"] = 4] = "LOCATION";
    LnadHeaders[LnadHeaders["LAND_PATTERN"] = 5] = "LAND_PATTERN";
    LnadHeaders[LnadHeaders["LNAD_USAGE"] = 6] = "LNAD_USAGE";
    LnadHeaders[LnadHeaders["LNAD_TYPE"] = 7] = "LNAD_TYPE";
    LnadHeaders[LnadHeaders["ADDRESS"] = 8] = "ADDRESS";
    LnadHeaders[LnadHeaders["POSITION"] = 9] = "POSITION";
    LnadHeaders[LnadHeaders["VALUATION"] = 10] = "VALUATION";
    LnadHeaders[LnadHeaders["LAND_SIZE"] = 11] = "LAND_SIZE";
    LnadHeaders[LnadHeaders["NUMBER_OF_OWNER"] = 12] = "NUMBER_OF_OWNER";
    LnadHeaders[LnadHeaders["ROAD_NEARBY"] = 13] = "ROAD_NEARBY";
    LnadHeaders[LnadHeaders["DIRECTION"] = 14] = "DIRECTION";
    LnadHeaders[LnadHeaders["WATER_ELECTRICITY_SUPPLY"] = 15] = "WATER_ELECTRICITY_SUPPLY";
    LnadHeaders[LnadHeaders["WIDTH"] = 16] = "WIDTH";
    LnadHeaders[LnadHeaders["DEEPTH"] = 17] = "DEEPTH";
    LnadHeaders[LnadHeaders["BUILDING_COVERAGE_RATE"] = 18] = "BUILDING_COVERAGE_RATE";
    LnadHeaders[LnadHeaders["VOLUME_RATE"] = 19] = "VOLUME_RATE";
    LnadHeaders[LnadHeaders["MEMO"] = 20] = "MEMO";
    LnadHeaders[LnadHeaders["CONTACT_PERSON"] = 21] = "CONTACT_PERSON";
    LnadHeaders[LnadHeaders["PICTURE_LINK"] = 22] = "PICTURE_LINK";
    LnadHeaders[LnadHeaders["OBJECT_CREATE_DATE"] = 23] = "OBJECT_CREATE_DATE";
    LnadHeaders[LnadHeaders["CONTRACT_DATE_FROM"] = 24] = "CONTRACT_DATE_FROM";
    LnadHeaders[LnadHeaders["CONTRACT_DATE_TO"] = 25] = "CONTRACT_DATE_TO";
    LnadHeaders[LnadHeaders["OBJECT_UPDATE_DATE"] = 26] = "OBJECT_UPDATE_DATE";
})(LnadHeaders || (LnadHeaders = {}));
/**
 * 將欄位索引轉換為 Google Sheet 的欄位字母 (A, B, C...)
 * @param {number} colIndex - 欄位索引 (0-based).
 * @returns {string} 欄位字母.
 */
function getColumnLetter(colIndex) {
    var temp, letter = '';
    var col = colIndex + 1;
    while (col > 0) {
        temp = (col - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        col = (col - temp - 1) / 26;
    }
    return letter;
}
function searchObjects(contractType, objectType, objectPattern, objectNmae, valuationFrom, valuationTo, landSizeFrom, landSizeTo, roadNearby, roomFrom, roomTo, isHasParkingSpace, buildingAgeFrom, buildingAgeTo, direction, objectWidthFrom, objectWidthTo, contactPerson) {
    // 根據所有搜尋參數建立一個快取鍵
    var cache = CacheService.getScriptCache();
    var cacheKey = "search_" + JSON.stringify(Array.from(arguments));
    var cachedResult = cache.get(cacheKey);
    // 如果快取中已有結果，直接回傳
    if (cachedResult) {
        console.log("Serving from cache.");
        return cachedResult;
    }
    console.log("Serving from live query.");
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetsToQuery = [];
    if (objectType) {
        var sheet = spreadsheet.getSheetByName(objectType);
        if (sheet)
            sheetsToQuery.push(sheet);
    }
    else {
        sheetsToQuery = spreadsheet.getSheets();
    }
    var allFilteredData = [];
    for (var _i = 0, sheetsToQuery_1 = sheetsToQuery; _i < sheetsToQuery_1.length; _i++) {
        var sheet = sheetsToQuery_1[_i];
        var sheetName = sheet.getName();
        var headers = (sheetName.toUpperCase() === 'BUILDING') ? BuildingHeaders : LnadHeaders;
        var queryParts = ['SELECT *'];
        var whereClauses = [];
        // 關鍵字查詢 (OR logic)
        if (objectNmae) {
            var keywordClauses = objectNmae.split(' ').filter(Boolean).map(function (keyword) {
                var upperKeyword = keyword.toUpperCase();
                return "(".concat(getColumnLetter(headers.OBJECT_NUMBER), " contains '").concat(upperKeyword, "' OR ")
                    + "".concat(getColumnLetter(headers.OBJECT_NAME), " contains '").concat(upperKeyword, "' OR ")
                    + "".concat(getColumnLetter(headers.LOCATION), " contains '").concat(upperKeyword, "' OR ")
                    + "".concat(getColumnLetter(headers.ADDRESS), " contains '").concat(upperKeyword, "')");
            });
            if (keywordClauses.length > 0) {
                whereClauses.push("(".concat(keywordClauses.join(' AND '), ")"));
            }
        }
        // 其他 AND 條件
        if (contractType)
            whereClauses.push("".concat(getColumnLetter(headers.CONTRACT_TYPE), " = '").concat(contractType, "'"));
        if (valuationFrom > 0)
            whereClauses.push("".concat(getColumnLetter(headers.VALUATION), " >= ").concat(valuationFrom));
        if (valuationTo > 0)
            whereClauses.push("".concat(getColumnLetter(headers.VALUATION), " <= ").concat(valuationTo));
        if (landSizeFrom > 0)
            whereClauses.push("".concat(getColumnLetter(headers.LAND_SIZE), " >= ").concat(landSizeFrom));
        if (landSizeTo > 0)
            whereClauses.push("".concat(getColumnLetter(headers.LAND_SIZE), " <= ").concat(landSizeTo));
        if (direction)
            whereClauses.push("".concat(getColumnLetter(headers.DIRECTION), " = '").concat(direction, "'"));
        if (objectWidthFrom > 0)
            whereClauses.push("".concat(getColumnLetter(headers.WIDTH), " >= ").concat(objectWidthFrom));
        if (objectWidthTo > 0)
            whereClauses.push("".concat(getColumnLetter(headers.WIDTH), " <= ").concat(objectWidthTo));
        if (contactPerson)
            whereClauses.push("".concat(getColumnLetter(headers.CONTACT_PERSON), " contains '").concat(contactPerson, "'"));
        if (roadNearby) {
            var range = roadNearby.split('|');
            if (range.length > 1) {
                whereClauses.push("".concat(getColumnLetter(headers.ROAD_NEARBY), " >= ").concat(range[0]));
                whereClauses.push("".concat(getColumnLetter(headers.ROAD_NEARBY), " <= ").concat(range[1]));
            }
        }
        if (sheetName.toUpperCase() === 'BUILDING') {
            if (objectPattern && objectPattern.length > 0) {
                var patternClauses = objectPattern.map(function (p) { return "".concat(getColumnLetter(BuildingHeaders.BUILDING_TYPE), " contains '").concat(p, "'"); });
                whereClauses.push("(".concat(patternClauses.join(' OR '), ")"));
            }
            if (isHasParkingSpace !== '') {
                var condition = isHasParkingSpace === '1' ? "!= '沒車位'" : "= '沒車位'";
                whereClauses.push("".concat(getColumnLetter(BuildingHeaders.VIHECLE_PARKING_TYPE), " ").concat(condition));
            }
            // 對於無法直接用 GQL 查詢的欄位 (例如需要分割字串)，在後續處理
        }
        else if (sheetName.toUpperCase() === 'LAND') {
            if (objectPattern && objectPattern.length > 0) {
                var patternClauses = objectPattern.map(function (p) { return "".concat(getColumnLetter(LnadHeaders.LNAD_USAGE), " contains '").concat(p, "'"); });
                whereClauses.push("(".concat(patternClauses.join(' OR '), ")"));
            }
        }
        if (whereClauses.length > 0) {
            queryParts.push('WHERE ' + whereClauses.join(' AND '));
        }
        var query = queryParts.join(' ');
        var url = "https://docs.google.com/spreadsheets/d/".concat(SPREADSHEET_ID, "/gviz/tq?sheet=").concat(sheetName, "&tq=").concat(encodeURIComponent(query));
        var response = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer ".concat(ScriptApp.getOAuthToken()) } });
        var text = response.getContentText();
        var json = JSON.parse(text.substring(text.indexOf('(') + 1, text.lastIndexOf(')')));
        if (json.table && json.table.rows.length > 0) {
            var rows = json.table.rows.map(function (r) { return r.c.map(function (cell) { return cell ? (cell.f || cell.v) : null; }); });
            // 後續過濾 (Post-filtering) for complex conditions
            var postFilteredRows = rows.filter(function (row) {
                if (sheetName.toUpperCase() === 'BUILDING') {
                    if (roomFrom > 0) {
                        var rooms = (row[BuildingHeaders.HOUSE_PATTERN] || '').toString().split('/')[0];
                        if (!rooms || parseInt(rooms) < roomFrom)
                            return false;
                    }
                    if (roomTo > 0) {
                        var rooms = (row[BuildingHeaders.HOUSE_PATTERN] || '').toString().split('/')[0];
                        if (!rooms || parseInt(rooms) > roomTo)
                            return false;
                    }
                    if (buildingAgeFrom > 0) {
                        var age = (row[BuildingHeaders.BUILDING_AGE] || '').toString().split('/').pop();
                        if (!age || parseInt(age) < buildingAgeFrom)
                            return false;
                    }
                    if (buildingAgeTo > 0) {
                        var age = (row[BuildingHeaders.BUILDING_AGE] || '').toString().split('/').pop();
                        if (!age || parseInt(age) > buildingAgeTo)
                            return false;
                    }
                }
                return true;
            });
            var extracted = postFilteredRows.map(function (row) {
                if (sheetName.toUpperCase() === 'BUILDING') {
                    return {
                        objectType: sheetName,
                        objectNumber: row[BuildingHeaders.OBJECT_NUMBER], // Pass objectNumber instead
                        objectNumber: row[BuildingHeaders.OBJECT_NUMBER],
                        objectName: row[BuildingHeaders.OBJECT_NAME],
                        valuation: row[BuildingHeaders.VALUATION],
                        landSize: row[BuildingHeaders.LAND_SIZE],
                        buildingSize: row[BuildingHeaders.BUILDING_SIZE],
                        housePattern: row[BuildingHeaders.HOUSE_PATTERN],
                        position: row[BuildingHeaders.POSITION],
                        location: row[BuildingHeaders.LOCATION],
                        address: row[BuildingHeaders.ADDRESS],
                        pictureLink: row[BuildingHeaders.PICTURE_LINK]
                    };
                }
                else { // LAND
                    return {
                        objectType: sheetName,
                        objectNumber: row[LnadHeaders.OBJECT_NUMBER], // Pass objectNumber instead
                        objectNumber: row[LnadHeaders.OBJECT_NUMBER],
                        objectName: row[LnadHeaders.OBJECT_NAME],
                        valuation: row[LnadHeaders.VALUATION],
                        landSize: row[LnadHeaders.LAND_SIZE],
                        buildingSize: 0,
                        housePattern: "",
                        position: row[LnadHeaders.POSITION],
                        location: row[LnadHeaders.LOCATION],
                        address: row[LnadHeaders.ADDRESS],
                        pictureLink: row[LnadHeaders.PICTURE_LINK]
                    };
                }
            });
            allFilteredData = allFilteredData.concat(extracted);
        }
    }
    // 因為 GQL 不回傳原始行號，我們需要重新查找
    // 為了簡化和效能，這裡暫時將 sequenceNumberInSheet 設為 0
    // 如果絕對需要行號，需要額外的一次性讀取來建立 map，但會降低效能
    var resultString = JSON.stringify(allFilteredData);
    // 將結果存入快取，設定 3 小時過期
    cache.put(cacheKey, resultString, 10800);
    return resultString;
}
var BuildingObjectData = /** @class */ (function () {
    function BuildingObjectData() {
    }
    return BuildingObjectData;
}());
var LandObjectData = /** @class */ (function () {
    function LandObjectData() {
    }
    return LandObjectData;
}());
function searchLastNumOfNumberedObjects(objectType) {
    var listOfSheet = new Array();
    if (objectType.toUpperCase() === 'BUILDING' || objectType.toUpperCase() === 'LAND') {
        var currentSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(objectType);
        if (currentSheet) {
            listOfSheet.push(currentSheet);
        }
    }
    else {
        listOfSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets();
    }
    // if object type is building, then the rule of object number is a 'A' + last number of numbered objects plus 1
    // if object type is land, then the rule of object number is a 'B' + last number of numbered objects plus 1
    var objectNumberPrefix = '';
    var objectNumberColumn = 0;
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            objectNumberPrefix = 'A';
            objectNumberColumn = BuildingHeaders.OBJECT_NUMBER;
            break;
        case 'LAND':
            objectNumberPrefix = 'B';
            objectNumberColumn = LnadHeaders.OBJECT_NUMBER;
            break;
        default:
    }
    var lastNumberOfObjectNumber = '';
    for (var _i = 0, listOfSheet_2 = listOfSheet; _i < listOfSheet_2.length; _i++) {
        var currentSheet = listOfSheet_2[_i];
        var dataRange = currentSheet.getDataRange();
        var values = dataRange.getValues();
        var headers = values.shift();
        var objectNumbers = values.map(function (row) {
            return row[objectNumberColumn];
        });
        lastNumberOfObjectNumber = objectNumbers.reduce(function (prev, current) {
            var isHasPrefix = current.toString().startsWith(objectNumberPrefix);
            var currentNumberPart = Number(current.toString().substring(1));
            var prevNumberPart = Number(prev.toString().substring(1));
            var isCurrentANumber = !isNaN(currentNumberPart);
            var isPrevANumber = !isNaN(prevNumberPart);
            if (!isPrevANumber) {
                prevNumberPart = 0;
            }
            if (isHasPrefix && isCurrentANumber) {
                return currentNumberPart > prevNumberPart ? current : prev;
            }
            return prev;
        });
        // const numberedObjectNumbers = objectNumbers.filter(function(objectNumber) {
        //     const isHasPrefix = objectNumber.toString().startsWith(objectNumberPrefix)
        //     const isNumber = !isNaN(Number(objectNumber.toString().substr(1)))
        //     return isHasPrefix && isNumber
        // })
        // numOfNumberedObjects += numberedObjectNumbers.length
    }
    return Number(lastNumberOfObjectNumber.toString().substring(1));
}
