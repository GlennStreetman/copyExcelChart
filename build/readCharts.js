var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import fs from 'fs';
import xml2js from 'xml2js';
import AdmZip from 'adm-zip';
function buildChartList(chartList) {
    let allCharts = [];
    Object.values(chartList).forEach((el) => {
        allCharts = allCharts.concat(el);
    });
    return allCharts;
}
function buildDrawingList(drawingList) {
    let allDrawings = [];
    Object.values(drawingList).forEach((el) => {
        allDrawings.push(el);
    });
    return allDrawings;
}
function buildFileList(chartRels, ref) {
    const returnList = [];
    Object.values(chartRels).forEach((chart) => {
        returnList.push(chart[ref]);
    });
    return returnList;
}
function buildChartDetails(tempFolder, worksheetNames, drawingList, drawingXMLs, chartList, cellRefs, definedNameRefs, chartRels, drawingrIds, definedNameKeys) {
    const workbook = {
        tempDir: tempFolder,
        drawingList: buildDrawingList(drawingList),
        drawingXMLs: drawingXMLs,
        chartList: buildChartList(chartList),
        defineNameRefs: definedNameKeys ? definedNameKeys : [],
        colorList: buildFileList(chartRels, 'colors'),
        styleList: buildFileList(chartRels, 'style'),
        worksheets: function () {
            const worksheetList = Object.entries(worksheetNames).reduce((acc, [key, val]) => {
                return Object.assign(Object.assign({}, acc), { [key]: {
                        name: val,
                        drawing: '',
                        drawingRels: {},
                        charts: {},
                    } });
            }, {});
            Object.keys(worksheetList).forEach((worksheet) => {
                if (drawingList[worksheet])
                    worksheetList[worksheet]['drawing'] = drawingList[worksheet];
                if (worksheetList[worksheet].drawing) {
                    const drawingName = worksheetList[worksheet].drawing;
                    worksheetList[worksheet].drawingRels = Object.entries(drawingrIds[drawingName]).reduce((acc, [key, val]) => { return Object.assign(Object.assign({}, acc), { [val]: key }); }, {}); //returns drawingName: rId
                    // worksheetList[worksheet].charts = {}
                    chartList[worksheetList[worksheet].drawing].forEach((chart) => {
                        worksheetList[worksheet].charts[chart] = {
                            chartRels: chartRels[chart],
                            cellRefs: cellRefs[chart],
                            definedNameRefs: definedNameRefs[chart],
                        };
                    });
                }
            });
            return worksheetList;
        }(),
        summary: function () {
            const summaryObj = {};
            Object.keys(this.worksheets).forEach((worksheet) => {
                if (this.worksheets[worksheet].charts) {
                    summaryObj[worksheet] = Object.entries(this.worksheets[worksheet].charts).reduce((acc, [key, val]) => {
                        const rtnObj = { [key]: [...new Set(val.cellRefs.concat(val.definedNameRefs))] };
                        return Object.assign(Object.assign({}, acc), rtnObj);
                    }, {});
                }
                else {
                    summaryObj[worksheet] = {};
                }
            });
            return summaryObj;
        }
    };
    return workbook;
}
function findChartRels(chartList, tempFolder) {
    const chartRels = {};
    Object.values(chartList).forEach((list) => {
        list.forEach((chart) => {
            chartRels[chart] = { colors: '', style: '' };
            const chartRelsXML = fs.readFileSync(`${tempFolder}/xl/charts/_rels/${chart}.xml.rels`, { encoding: 'utf-8' });
            xml2js.parseString(chartRelsXML, (error, res) => __awaiter(this, void 0, void 0, function* () {
                res.Relationships.Relationship.forEach((rel) => {
                    var _a, _b;
                    if (((_a = rel['$']) === null || _a === void 0 ? void 0 : _a.Target) && rel['$'].Target.includes('color'))
                        chartRels[chart]['colors'] = rel['$'].Target.replace('.xml', '');
                    if (((_b = rel['$']) === null || _b === void 0 ? void 0 : _b.Target) && rel['$'].Target.includes('style'))
                        chartRels[chart]['style'] = rel['$'].Target.replace('.xml', '');
                });
            }));
        });
    });
    return chartRels;
}
function findChartCellRefs(chartList, aliaslist, tempFolder, definedNames) {
    const cellRefs = {};
    const definedNameRefs = {};
    const tempRefs = {};
    Object.values(chartList).forEach((chartList) => {
        chartList.forEach((chart) => {
            cellRefs[chart] = [];
            const chartXML = fs.readFileSync(`${tempFolder}/xl/charts/${chart}.xml`, { encoding: 'utf-8' });
            aliaslist.forEach((alias) => {
                //find matching cell ranges
                //check for matching without commas around names 
                let worksheetCellRefRange = `${alias}!\\$[A-Z]{1,3}\\$[0-9]{1,7}:\\$[A-Z]{1,3}\\$[0-9]{1,7}`;
                let findCellRef = new RegExp(worksheetCellRefRange, 'g');
                let matchListNoCommas = [...new Set(chartXML.match(findCellRef))];
                //check for matching WITH commas around names 
                worksheetCellRefRange = `'${alias}'!\\$[A-Z]{1,3}\\$[0-9]{1,7}:\\$[A-Z]{1,3}\\$[0-9]{1,7}`;
                findCellRef = new RegExp(worksheetCellRefRange, 'g');
                let matchListCommas = [...new Set(chartXML.match(findCellRef))];
                let matchList = [...new Set(matchListNoCommas.concat(matchListCommas))];
                //find matching cell refs.
                let worksheetCellRef = `${alias}!\\$[A-Z]{1,3}\\$[0-9]{1,7}<`;
                let findCellRefCell = new RegExp(worksheetCellRef, 'g');
                let matchListCellNoCommas = [...new Set(chartXML.match(findCellRefCell))];
                matchListCellNoCommas = matchListCellNoCommas.map(el => el.slice(0, el.length - 1));
                //check for matching WITH commas around names 
                worksheetCellRef = `'${alias}'!\\$[A-Z]{1,3}\\$[0-9]{1,7}<`;
                findCellRefCell = new RegExp(worksheetCellRef, 'g');
                let matchListCellCommas = [...new Set(chartXML.match(findCellRefCell))];
                matchListCellCommas = matchListCellCommas.map(el => el.slice(0, el.length - 1));
                matchList = [...new Set(matchList.concat(matchListCellNoCommas).concat(matchListCellCommas))];
                cellRefs[chart] = [...new Set(cellRefs[chart].concat(matchList))];
            });
            //some chart types use named ranges that are stored in worbook.xml. The refs to workbook.xml look like _xlchart.v?.?
            definedNameRefs[chart] = [];
            let refRegex = new RegExp(`>_xlchart.v[0-9]{1,9}.[0-9]{1,10}<`, 'g');
            let matchingDefinedNameRefs = [...new Set(chartXML.match(refRegex))].forEach((el) => {
                const foundRef = el.slice(1, el.length - 1);
                tempRefs[foundRef] = chart;
            });
        });
    });
    Object.entries(definedNames).forEach(([key, val]) => {
        if (tempRefs[key])
            definedNameRefs[tempRefs[key]].push(val);
    });
    const definedNameKeys = Object.keys(tempRefs);
    return [cellRefs, definedNameRefs, definedNameKeys];
}
function findDrawingXML(drawingObj, sourceFolder) {
    const returnObj = {};
    Object.keys(drawingObj).forEach((drawingName) => {
        returnObj[drawingName] = {};
        const drawingXML = fs.readFileSync(`${sourceFolder}/xl/drawings/${drawingName}.xml`, { encoding: 'utf-8' });
        xml2js.parseString(drawingXML, (error, res) => __awaiter(this, void 0, void 0, function* () {
            const targetAnchors = res['xdr:wsDr']['xdr:twoCellAnchor'];
            targetAnchors.forEach((el) => {
                var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r, _s, _t, _u;
                const rIdRegular = (_j = (_h = (_g = (_f = (_e = (_d = (_c = (_b = (_a = el === null || el === void 0 ? void 0 : el['xdr:graphicFrame']) === null || _a === void 0 ? void 0 : _a[0]) === null || _b === void 0 ? void 0 : _b['a:graphic']) === null || _c === void 0 ? void 0 : _c[0]) === null || _d === void 0 ? void 0 : _d['a:graphicData']) === null || _e === void 0 ? void 0 : _e[0]) === null || _f === void 0 ? void 0 : _f['c:chart']) === null || _g === void 0 ? void 0 : _g[0]) === null || _h === void 0 ? void 0 : _h['$']) === null || _j === void 0 ? void 0 : _j['r:id']; //regular chart?.xml
                if (rIdRegular) {
                    const chartName = drawingObj[drawingName][rIdRegular];
                    returnObj[drawingName][chartName] = el;
                }
                const rIdRegularEx = (_u = (_t = (_s = (_r = (_q = (_p = (_o = (_m = (_l = (_k = el === null || el === void 0 ? void 0 : el['mc:AlternateContent']) === null || _k === void 0 ? void 0 : _k[0]) === null || _l === void 0 ? void 0 : _l['mc:Choice']) === null || _m === void 0 ? void 0 : _m[0]) === null || _o === void 0 ? void 0 : _o['xdr:graphicFrame']) === null || _p === void 0 ? void 0 : _p[0]) === null || _q === void 0 ? void 0 : _q['a:graphic']) === null || _r === void 0 ? void 0 : _r[0]) === null || _s === void 0 ? void 0 : _s['a:graphicData']) === null || _t === void 0 ? void 0 : _t[0]) === null || _u === void 0 ? void 0 : _u['cx:chart'][0]['$']['r:id']; //alternate chart type chartEx?.xml
                if (rIdRegularEx) {
                    const chartName = drawingObj[drawingName][rIdRegularEx];
                    returnObj[drawingName][chartName] = el;
                }
            });
        }));
    });
    return returnObj;
}
function parseDrawingRels(drawingList, tempFolder) {
    const chartList = {};
    const drawingrIds = {};
    Object.entries(drawingList).forEach(([alias, name]) => {
        const drawingRelsXML = fs.readFileSync(`${tempFolder}/xl/drawings/_rels/${name}.xml.rels`, { encoding: 'utf-8' });
        xml2js.parseString(drawingRelsXML, (error, res) => __awaiter(this, void 0, void 0, function* () {
            drawingrIds[name] = {};
            const chartRels = res.Relationships.Relationship.reduce((acc, el) => {
                if (el['$'].Target.includes('charts/')) {
                    const chartName = el['$'].Target.replace('../charts/', '').replace('.xml', '');
                    acc.push(chartName);
                    drawingrIds[name][el['$'].Id] = chartName;
                    return acc;
                }
                else {
                    return acc;
                }
            }, []);
            chartList[name] = chartRels;
        }));
    });
    return [chartList, drawingrIds];
}
function findDrawingRels(sheetNames, tempFolder) {
    let drawingRelsAcc = {};
    Object.entries(sheetNames).forEach(([alias, name]) => {
        if (fs.existsSync(`${tempFolder}/xl/worksheets/_rels/${name}.xml.rels`)) {
            const worksheetRelsXML = fs.readFileSync(`${tempFolder}/xl/worksheets/_rels/${name}.xml.rels`, { encoding: 'utf-8' });
            xml2js.parseString(worksheetRelsXML, (error, res) => __awaiter(this, void 0, void 0, function* () {
                const drawingRels = res.Relationships.Relationship.reduce((acc, el) => {
                    if (el['$'].Target.includes('drawings/')) {
                        return Object.assign(Object.assign({}, acc), { [alias]: el['$'].Target.replace('../drawings/', '').replace('.xml', '') });
                    }
                    else {
                        return acc;
                    }
                }, {});
                drawingRelsAcc = Object.assign(Object.assign({}, drawingRelsAcc), drawingRels);
            }));
        }
    });
    return drawingRelsAcc;
}
function readSheetNames(tempFolder) {
    //read workbook.xml and workbook.xml.rels to find list of worksheet file names and aliases.
    //aliases are worksheet names visable to excel users.
    let sheetIds = {};
    let definedName = {}; //named cell references, refered to some type of charts, often named chartEX?.xml
    const workbookXML = fs.readFileSync(`${tempFolder}xl/workbook.xml`, { encoding: 'utf-8' });
    xml2js.parseString(workbookXML, (error, res) => {
        var _a, _b;
        const sheetRelAlias = ((_a = res === null || res === void 0 ? void 0 : res.workbook) === null || _a === void 0 ? void 0 : _a.sheets) ? res.workbook.sheets[0].sheet.reduce((acc, el) => {
            return Object.assign(Object.assign({}, acc), { [el['$']['r:id']]: el['$'].name });
        }, {}) : {};
        sheetIds = sheetRelAlias;
        definedName = ((_b = res === null || res === void 0 ? void 0 : res.workbook) === null || _b === void 0 ? void 0 : _b.definedNames) ? res.workbook.definedNames[0].definedName.reduce((acc, el) => {
            return Object.assign(Object.assign({}, acc), { [el['$']['name']]: el['_'] });
        }, {}) : {};
    });
    const workbookXML_rels = fs.readFileSync(`${tempFolder}xl/_rels/workbook.xml.rels`, { encoding: 'utf-8' });
    xml2js.parseString(workbookXML_rels, (error, res) => {
        const sheetRelName = res.Relationships.Relationship.reduce((acc, el) => {
            if (sheetIds[el['$'].Id]) {
                return Object.assign(Object.assign({}, acc), { [sheetIds[el['$'].Id]]: el['$'].Target.replace('worksheets/', '').replace('.xml', '') });
            }
            else {
                return Object.assign({}, acc);
            }
        }, {});
        sheetIds = sheetRelName;
    });
    return [sheetIds, definedName];
}
export function readCharts(sourceFile, //location of file to read
tempFolder) {
    const filePath = sourceFile.replace(/\\/, 'g');
    const fileName = filePath.slice(filePath.lastIndexOf('/') + 1, filePath.length).replace('.xlsx', '/');
    const sourceFolder = `${tempFolder}/${fileName}`;
    if (fs.existsSync(sourceFolder))
        fs.rmdirSync(sourceFolder, { recursive: true }); //remove old files that have been parced at the same location.
    fs.mkdirSync(sourceFolder);
    return new Promise((resolve, reject) => {
        try { //under the hood, excel files are zip folders containing xml files.
            const zip = new AdmZip(sourceFile);
            zip.extractAllTo(sourceFolder, true); //unzip excel template file to dump folder so that we can access xml files.
            const [worksheetNames, definedNames] = readSheetNames(sourceFolder); //returns  {[alias]: worksheet.xml name}
            const drawingList = findDrawingRels(worksheetNames, sourceFolder); //returns {[alias]: drawings}
            const [chartList, drawingrIds] = Object.keys(drawingList).length > 0 ? parseDrawingRels(drawingList, sourceFolder) : [{}, {}]; //find associated charts and xml blob associated with drawing in drawing.xml
            const findDrawingXMLs = Object.keys(drawingrIds).length > 0 ? findDrawingXML(drawingrIds, sourceFolder) : drawingrIds; //drawing xml file.
            const [cellRefs, definedNameRefs, definedNameKeys] = Object.keys(chartList).length > 0 ? findChartCellRefs(chartList, Object.keys(worksheetNames), sourceFolder, definedNames) : [{}, {}]; //excel formula ranges and and cell refs.
            const chartRefs = Object.keys(chartList).length > 0 ? findChartRels(chartList, sourceFolder) : {}; //related chart xmls. Style & colors.
            const chartDetails = buildChartDetails(sourceFolder, worksheetNames, drawingList, findDrawingXMLs, chartList, cellRefs, definedNameRefs, chartRefs, drawingrIds, definedNameKeys);
            resolve(chartDetails);
        }
        catch (error) {
            console.log('Read file error. Path: ', sourceFile, 'Error: ', error);
            reject(error);
        }
    });
}
//# sourceMappingURL=readCharts.js.map