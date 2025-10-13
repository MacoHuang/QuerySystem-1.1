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
function showObjectInfo(objectType, sequenceNumberInSheet) {
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            var template = HtmlService.createTemplateFromFile('buildingInfo');
            var dataString = searchObjectInfo(objectType, sequenceNumberInSheet);
            var buildingObject = JSON.parse(dataString);
            template.buildingObject = buildingObject;
            console.log(JSON.stringify(buildingObject));
            return template.evaluate().getContent();
        case 'LAND':
            var landTemplate = HtmlService.createTemplateFromFile('landInfo');
            var landDataString = searchObjectInfo(objectType, sequenceNumberInSheet);
            var landObject = JSON.parse(landDataString);
            landTemplate.landObject = landObject;
            console.log(JSON.stringify(landObject));
            return landTemplate.evaluate().getContent();
    }
    return "";
}
function showObjectA4Info(objectType, sequenceNumberInSheet) {
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            // const buildingTemplate = HtmlService.createTemplateFromFile('buildingA4')
            var dataString = searchObjectInfo(objectType, sequenceNumberInSheet);
            var buildingObject = JSON.parse(dataString);
            // buildingTemplate.buildingObject = buildingObject
            // console.log(JSON.stringify(buildingObject))
            return createContract(objectType, buildingObject);
        // return buildingTemplate.evaluate()
        case 'LAND':
            // const landTemplate = HtmlService.createTemplateFromFile('landA4')
            var landDataString = searchObjectInfo(objectType, sequenceNumberInSheet);
            var landObject = JSON.parse(landDataString);
            // landTemplate.landObject = landObject
            // console.log(JSON.stringify(landObject))
            return createContract(objectType, landObject);
        // return landTemplate.evaluate()
    }
    return "";
}
function searchObjectInfo(objectType, sequenceNumberInSheet) {
    var currentSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(objectType);
    var dataRange = currentSheet === null || currentSheet === void 0 ? void 0 : currentSheet.getDataRange();
    var values = dataRange === null || dataRange === void 0 ? void 0 : dataRange.getValues();
    var headers = values === null || values === void 0 ? void 0 : values.shift();
    var row = values === null || values === void 0 ? void 0 : values.find(function (row) {
        return values.indexOf(row) === sequenceNumberInSheet - 1;
    });
    console.log("row:".concat(row));
    if (!row) {
        return "";
    }
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
            return JSON.stringify(buildingObject);
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
            return JSON.stringify(landObject);
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
function searchObjects(contractType, objectType, objectPattern, objectNmae, valuationFrom, valuationTo, landSizeFrom, landSizeTo, roadNearby, roomFrom, roomTo, isHasParkingSpace, buildingAgeFrom, buildingAgeTo, direction, objectWidthFrom, objectWidthTo, contactPerson) {
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
    var filteredValues = new Map();
    var _loop_1 = function (currentSheet) {
        var dataRange = currentSheet.getDataRange();
        var values = dataRange.getValues();
        var headers = values.shift();
        console.log(objectPattern);
        currentfilteredValues = values
            .map(function (row) {
            var obj = {};
            obj = [values.indexOf(row) + 1, row];
            return obj;
        })
            .filter(function (row) {
            var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m;
            var andConditionList = new Array();
            var orConditionList = new Array();
            var roadNearbyRange = roadNearby.split('|');
            var objectNameKeywordList = objectNmae.split(' ');
            var sheetName = currentSheet.getName().toUpperCase();
            switch (sheetName) {
                case 'BUILDING':
                    andConditionList.push(((_a = row[1][BuildingHeaders.CONTRACT_TYPE]) === null || _a === void 0 ? void 0 : _a.toString().indexOf(contractType)) > -1);
                    // andConditionList.push(
                    //     row[1][BuildingHeaders.OBJECT_NAME]?.toString().indexOf(objectNmae) > -1 ||
                    //     row[1][BuildingHeaders.LOCATION]?.toString().indexOf(objectNmae) > -1 ||
                    //     row[1][BuildingHeaders.ADDRESS]?.toString().indexOf(objectNmae) > -1
                    // )
                    orConditionList.push(objectNameKeywordList.some(function (keywords) {
                        var _a;
                        return ((_a = row[1][LnadHeaders.OBJECT_NUMBER]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                    }));
                    orConditionList.push(objectNameKeywordList.some(function (keywords) {
                        var _a;
                        return ((_a = row[1][BuildingHeaders.OBJECT_NAME]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                    }));
                    orConditionList.push(objectNameKeywordList.some(function (keywords) {
                        var _a;
                        return ((_a = row[1][BuildingHeaders.LOCATION]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                    }));
                    orConditionList.push(objectNameKeywordList.some(function (keywords) {
                        var _a;
                        return ((_a = row[1][BuildingHeaders.ADDRESS]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                    }));
                    var buildingUsageList = (_b = row[1][BuildingHeaders.BUILDING_TYPE]) === null || _b === void 0 ? void 0 : _b.toString().split(',');
                    // andConditionList.push(objectPattern.includes(row[1][BuildingHeaders.BUILDING_TYPE]?.toString()))
                    andConditionList.push(objectPattern.some(function (pattern) {
                        return buildingUsageList.includes(pattern);
                    }));
                    if (roadNearbyRange && roadNearbyRange.length > 1) {
                        andConditionList.push(row[1][BuildingHeaders.ROAD_NEARBY] >= roadNearbyRange[0] && row[1][BuildingHeaders.ROAD_NEARBY] <= roadNearbyRange[1]);
                    }
                    if (valuationFrom > 0) {
                        andConditionList.push(row[1][BuildingHeaders.VALUATION] >= valuationFrom);
                    }
                    if (valuationTo > 0) {
                        andConditionList.push(row[1][BuildingHeaders.VALUATION] <= valuationTo);
                    }
                    if (landSizeFrom > 0) {
                        andConditionList.push(row[1][BuildingHeaders.LAND_SIZE] >= landSizeFrom);
                    }
                    if (landSizeTo > 0) {
                        andConditionList.push(row[1][BuildingHeaders.LAND_SIZE] <= landSizeTo);
                    }
                    var roomOfBuilding = row[1][BuildingHeaders.HOUSE_PATTERN].toString().split('/');
                    if (roomFrom > 0 && roomOfBuilding.length > 0) {
                        andConditionList.push(roomOfBuilding[0] >= roomFrom);
                    }
                    if (roomTo > 0 && roomOfBuilding.length > 0) {
                        andConditionList.push(roomOfBuilding[0] <= roomTo);
                    }
                    //console.log(`isHasParkingSpace:${isHasParkingSpace}`)
                    if (isHasParkingSpace !== '') {
                        var matchCondition = isHasParkingSpace === '1';
                        console.log("matchCondition:".concat(matchCondition));
                        console.log("VIHECLE_PARKING_TYPE:".concat((_c = row[1][BuildingHeaders.VIHECLE_PARKING_TYPE]) === null || _c === void 0 ? void 0 : _c.toString().trim()));
                        console.log("VIHECLE_PARKING_TYPE:".concat(((_d = row[1][BuildingHeaders.VIHECLE_PARKING_TYPE]) === null || _d === void 0 ? void 0 : _d.toString().trim()) != '沒車位'));
                        andConditionList.push((((_e = row[1][BuildingHeaders.VIHECLE_PARKING_TYPE]) === null || _e === void 0 ? void 0 : _e.toString().trim()) != '沒車位') == matchCondition);
                    }
                    // andConditionList.push(row[1][BuildingHeaders.WATER_SUPPLY]?.toString().indexOf(waterSupply) > -1)
                    andConditionList.push(((_f = row[1][BuildingHeaders.DIRECTION]) === null || _f === void 0 ? void 0 : _f.toString().indexOf(direction)) > -1);
                    if (objectWidthFrom > 0) {
                        andConditionList.push(row[1][BuildingHeaders.WIDTH] >= objectWidthFrom);
                    }
                    if (objectWidthTo > 0) {
                        andConditionList.push(row[1][BuildingHeaders.WIDTH] <= objectWidthTo);
                    }
                    var buildingAge = ((_g = row[1][BuildingHeaders.BUILDING_AGE]) === null || _g === void 0 ? void 0 : _g.toString().split('/').pop()) || '0';
                    if (buildingAgeFrom > 0) {
                        andConditionList.push(buildingAge >= buildingAgeFrom);
                    }
                    if (buildingAgeTo > 0) {
                        andConditionList.push(buildingAge <= buildingAgeTo);
                    }
                    andConditionList.push(((_h = row[1][BuildingHeaders.CONTACT_PERSON]) === null || _h === void 0 ? void 0 : _h.toString().indexOf(contactPerson)) > -1);
                    break;
                case 'LAND':
                    andConditionList.push(((_j = row[1][LnadHeaders.CONTRACT_TYPE]) === null || _j === void 0 ? void 0 : _j.toString().indexOf(contractType)) > -1);
                    // andConditionList.push(
                    //     row[1][LnadHeaders.OBJECT_NAME]?.toString().indexOf(objectNmae) > -1 ||
                    //     row[1][LnadHeaders.LOCATION]?.toString().indexOf(objectNmae) > -1 ||
                    //     row[1][LnadHeaders.ADDRESS]?.toString().indexOf(objectNmae) > -1
                    // )
                    orConditionList.push(objectNameKeywordList.some(function (keywords) {
                        var _a;
                        return ((_a = row[1][LnadHeaders.OBJECT_NUMBER]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                    }));
                    orConditionList.push(objectNameKeywordList.some(function (keywords) {
                        var _a;
                        return ((_a = row[1][LnadHeaders.OBJECT_NAME]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                    }));
                    orConditionList.push(objectNameKeywordList.some(function (keywords) {
                        var _a;
                        return ((_a = row[1][LnadHeaders.LOCATION]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                    }));
                    orConditionList.push(objectNameKeywordList.some(function (keywords) {
                        var _a;
                        return ((_a = row[1][LnadHeaders.ADDRESS]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                    }));
                    var landUsageList = (_k = row[1][LnadHeaders.LNAD_USAGE]) === null || _k === void 0 ? void 0 : _k.toString().split(',');
                    // andConditionList.push(objectPattern.includes(row[1][LnadHeaders.LNAD_USAGE]?.toString()))
                    andConditionList.push(objectPattern.some(function (pattern) {
                        return landUsageList.includes(pattern);
                    }));
                    if (roadNearbyRange && roadNearbyRange.length > 1) {
                        andConditionList.push(row[1][LnadHeaders.ROAD_NEARBY] >= roadNearbyRange[0] && row[1][LnadHeaders.ROAD_NEARBY] <= roadNearbyRange[1]);
                    }
                    if (valuationFrom > 0) {
                        andConditionList.push(row[1][LnadHeaders.VALUATION] >= valuationFrom);
                    }
                    if (valuationTo > 0) {
                        andConditionList.push(row[1][LnadHeaders.VALUATION] <= valuationTo);
                    }
                    if (landSizeFrom > 0) {
                        andConditionList.push(row[1][LnadHeaders.LAND_SIZE] >= landSizeFrom);
                    }
                    if (landSizeTo > 0) {
                        andConditionList.push(row[1][LnadHeaders.LAND_SIZE] <= landSizeTo);
                    }
                    // if(waterSupply !== '') {
                    //     let matchCondition = waterSupply === '1';
                    //     console.log(`matchCondition:${matchCondition}`)
                    //     console.log(`WATER_ELECTRICITY_SUPPLY:${row[1][LnadHeaders.WATER_ELECTRICITY_SUPPLY]?.toString().trim()}`)
                    //     console.log(`WATER_ELECTRICITY_SUPPLY:${row[1][LnadHeaders.WATER_ELECTRICITY_SUPPLY]?.toString().trim() !== ''}`)
                    //     andConditionList.push((row[1][LnadHeaders.WATER_ELECTRICITY_SUPPLY]?.toString().trim() !== '') == matchCondition)
                    // }
                    // andConditionList.push(row[1][LnadHeaders.WATER_ELECTRICITY_SUPPLY]?.toString().indexOf(waterSupply) > -1)
                    andConditionList.push(((_l = row[1][LnadHeaders.DIRECTION]) === null || _l === void 0 ? void 0 : _l.toString().indexOf(direction)) > -1);
                    if (objectWidthFrom > 0) {
                        andConditionList.push(row[1][LnadHeaders.WIDTH] >= objectWidthFrom);
                    }
                    if (objectWidthTo > 0) {
                        andConditionList.push(row[1][LnadHeaders.WIDTH] <= objectWidthTo);
                    }
                    andConditionList.push(((_m = row[1][LnadHeaders.CONTACT_PERSON]) === null || _m === void 0 ? void 0 : _m.toString().indexOf(contactPerson)) > -1);
                    break;
                default:
            }
            andConditionList.forEach(function (value, index) {
                console.log("".concat(sheetName, ":").concat(index, " ").concat(value));
            });
            var orCondition = orConditionList.some(Boolean);
            console.log("orCondition:".concat(orCondition));
            return andConditionList.every(Boolean) && orCondition;
        });
        filteredValues = filteredValues.set(currentSheet.getName(), currentfilteredValues);
    };
    var currentfilteredValues;
    for (var _i = 0, listOfSheet_1 = listOfSheet; _i < listOfSheet_1.length; _i++) {
        var currentSheet = listOfSheet_1[_i];
        _loop_1(currentSheet);
    }
    console.log("filteredValues.size:".concat(filteredValues.size));
    var extractedData = [];
    Array.from(filteredValues).map(function (_a) {
        var key = _a[0], filteredData = _a[1];
        console.log("key:".concat(key, ", filteredData.length:").concat(filteredData.length));
        var temp = filteredData.map(function (row) {
            var data = {};
            switch (key.toUpperCase()) {
                case 'BUILDING':
                    data = {
                        objectType: key,
                        sequenceNumberInSheet: row[0],
                        objectNumber: row[1][BuildingHeaders.OBJECT_NUMBER],
                        objectName: row[1][BuildingHeaders.OBJECT_NAME],
                        valuation: row[1][BuildingHeaders.VALUATION],
                        landSize: row[1][BuildingHeaders.LAND_SIZE],
                        buildingSize: row[1][BuildingHeaders.BUILDING_SIZE],
                        housePattern: row[1][BuildingHeaders.HOUSE_PATTERN],
                        position: row[1][BuildingHeaders.POSITION],
                        location: row[1][BuildingHeaders.LOCATION],
                        address: row[1][BuildingHeaders.ADDRESS],
                        pictureLink: row[1][BuildingHeaders.PICTURE_LINK]
                    };
                    break;
                case 'LAND':
                    data = {
                        objectType: key,
                        sequenceNumberInSheet: row[0],
                        objectNumber: row[1][LnadHeaders.OBJECT_NUMBER],
                        objectName: row[1][LnadHeaders.OBJECT_NAME],
                        valuation: row[1][LnadHeaders.VALUATION],
                        landSize: row[1][LnadHeaders.LAND_SIZE],
                        buildingSize: 0,
                        housePattern: "",
                        position: row[1][LnadHeaders.POSITION],
                        location: row[1][LnadHeaders.LOCATION],
                        address: row[1][LnadHeaders.ADDRESS],
                        pictureLink: row[1][LnadHeaders.PICTURE_LINK]
                    };
                    break;
                default:
                    break;
            }
            //console.log(data)
            return data;
        });
        // console.log(temp)
        extractedData = extractedData.concat(temp);
    });
    // console.log("BuildingHeaders[0]:" + BuildingHeaders[0])
    // console.log(extractedData)
    return JSON.stringify(extractedData);
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
