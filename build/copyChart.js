import fs from 'fs';
import xml2js from 'xml2js';
function copyDefineNames(//
sourceExcel, targetExcel, newDefinedNamesObj, stringOverrides) {
    if (Object.keys(newDefinedNamesObj).length > 0) {
        const sourceDir = sourceExcel.tempDir;
        const targetDir = targetExcel.tempDir;
        let addDefs = [];
        const sourceWookbook = `${sourceDir}xl/workbook.xml`;
        const souceDefs = fs.readFileSync(sourceWookbook, { encoding: 'utf-8' });
        xml2js.parseString(souceDefs, (error, editXML) => {
            editXML.workbook.definedNames[0].definedName.forEach((rel) => {
                if (newDefinedNamesObj[rel['$'].name])
                    rel['$'].name = newDefinedNamesObj[rel['$'].name];
                if (stringOverrides[rel['_']])
                    rel['_'] = stringOverrides[rel['_']];
                addDefs.push(rel);
            });
        });
        const outputWorkbook = `${targetDir}xl/workbook.xml`;
        const outputFile = fs.readFileSync(outputWorkbook, { encoding: 'utf-8' });
        xml2js.parseString(outputFile, (error, editXML) => {
            if (editXML.workbook.definedNames) {
                editXML.workbook.definedNames[0].definedName.concat(addDefs);
            }
            else { //need to copy from old xml, and insure that definedNames follows </sheets> tag
                const newWorkbookObj = {};
                Object.entries(editXML.workbook).forEach(([key, val]) => {
                    if (key !== 'sheets') {
                        newWorkbookObj[key] = val;
                    }
                    else {
                        newWorkbookObj[key] = val;
                        newWorkbookObj['definedNames'] = [{ definedName: addDefs }];
                    }
                });
                editXML.workbook = newWorkbookObj;
            }
            const builder = new xml2js.Builder();
            const xml = builder.buildObject(editXML);
            fs.writeFileSync(`${targetDir}/xl/workbook.xml`, xml);
        });
    }
}
function updateContentTypes(contentTypesUpdateObj, sourceExcel, targetExcel, sourceWorksheet, chartToCopy, newChartName) {
    const sourceDir = sourceExcel.tempDir;
    const targetDir = targetExcel.tempDir;
    const updateTags = [];
    //read source contentTypes, copy source tags to list after updating PartName
    const sourceContent = fs.readFileSync(`${sourceDir}/[Content_Types].xml`, { encoding: 'utf-8' });
    xml2js.parseString(sourceContent, (error, editXML) => {
        editXML.Types.Override.forEach((rel) => {
            if (contentTypesUpdateObj[rel['$'].PartName]) {
                rel['$'].PartName = contentTypesUpdateObj[rel['$'].PartName];
                updateTags.push(rel);
            }
        });
    });
    //update output contentTypes
    const targetContent = fs.readFileSync(`${targetDir}/[Content_Types].xml`, { encoding: 'utf-8' });
    xml2js.parseString(targetContent, (error, editXML) => {
        editXML.Types.Override = editXML.Types.Override.concat(updateTags);
        const builder = new xml2js.Builder();
        const xml = builder.buildObject(editXML);
        fs.writeFileSync(`${targetDir}/[Content_Types].xml`, xml);
    });
}
function getNewName(newName, targetList, iterator = 0) {
    if (!targetList.includes(newName)) {
        targetList.push(newName);
        return newName;
    }
    else {
        const updateIterator = iterator + 1;
        const updateNewName = `${newName.replace(new RegExp("[0-9]", "g"), "")}${updateIterator}`;
        return getNewName(updateNewName, targetList, updateIterator);
    }
}
function getNewDdefinedNameRef(newName, targetList, iterator = 0) {
    if (!targetList.includes(newName)) {
        targetList.push(newName);
        return newName;
    }
    else {
        const updateIterator = iterator + 1;
        const updateNewName = newName.slice(0, newName.lastIndexOf('.')) + iterator;
        return getNewDdefinedNameRef(updateNewName, targetList, updateIterator);
    }
}
function copyChartFiles(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, newChartName, stringOverrides, contentTypesUpdateObj) {
    //copy all four files, update names
    //update string overrides for chart.xml
    const sourceDir = sourceExcel.tempDir;
    const targetDir = targetExcel.tempDir;
    if (!fs.existsSync(`${targetDir}xl/charts/_rels/`))
        fs.mkdirSync(`${targetDir}xl/charts/_rels/`, { recursive: true });
    const getNewColorsFileName = getNewName('colors1', targetExcel.colorList);
    const getNewStyleFileName = getNewName('style1', targetExcel.styleList);
    //COPY SOURCE RELS FILE
    const sourceRelsFile = `${sourceDir}xl/charts/_rels/${chartToCopy}.xml.rels`;
    const sourceRelsXML = fs.readFileSync(sourceRelsFile, { encoding: 'utf-8' });
    xml2js.parseString(sourceRelsXML, (error, editXML) => {
        editXML.Relationships.Relationship.forEach((rel) => {
            if (rel['$'].Target.includes('colors'))
                rel['$'].Target = `${getNewColorsFileName}.xml`;
            if (rel['$'].Target.includes('style'))
                rel['$'].Target = `${getNewStyleFileName}.xml`;
        });
        const builder = new xml2js.Builder();
        const xml = builder.buildObject(editXML);
        fs.writeFileSync(`${targetDir}/xl/charts/_rels/${newChartName}.xml.rels`, xml);
    });
    //COPY SOURCE CHART FILE. Update any cell references refs (ex. A1:B2) AND definedName name refs.(ex. '_xlchart.v1.0')
    const sourceChartFile = `${sourceDir}xl/charts/${chartToCopy}.xml`;
    let sourceChartXML = fs.readFileSync(sourceChartFile, { encoding: 'utf-8' });
    //update cell references.
    Object.entries(stringOverrides).forEach(([key, val]) => {
        const newKey = key.replace(/\$/g, '\\$');
        const regExKey = new RegExp(`>${newKey}<`, 'g');
        sourceChartXML = sourceChartXML.replace(regExKey, `>${val}<`);
    });
    //update definedName Refs
    const newDefinedNamesObj = sourceExcel.worksheets[sourceWorksheet].charts[chartToCopy].definedNameRefs.reduce((acc, el) => {
        const newDefinedNameRef = getNewDdefinedNameRef(el, targetExcel.defineNames);
        sourceChartXML = sourceChartXML.replace(new RegExp(`>${el}<`, 'g'), `>${newDefinedNameRef}<`);
        return Object.assign(Object.assign({}, acc), { [el]: newDefinedNameRef });
    }, {});
    fs.writeFileSync(`${targetDir}/xl/charts/${newChartName}.xml`, sourceChartXML);
    contentTypesUpdateObj[`/xl/charts/${chartToCopy}.xml`] = `/xl/charts/${newChartName}.xml`;
    //COPY Chart colors?.xml and style?.xml
    Object.entries(sourceExcel.worksheets[sourceWorksheet].charts[chartToCopy].chartRels).forEach(([key, val]) => {
        // const updateFileName = `${key}${newChartName.replace(/[A-z]/g, '')}.xml`
        const thisFileName = key === 'colors' ? getNewColorsFileName : getNewStyleFileName;
        fs.copyFileSync(`${sourceDir}xl/charts/${val}.xml`, `${targetDir}xl/charts/${thisFileName}.xml`);
        contentTypesUpdateObj[`/xl/charts/${val}.xml`] = `/xl/charts/${thisFileName}.xml`;
    });
    return newDefinedNamesObj;
}
function addWorksheetRelsFile(rId, newDrawingName, target, source, targetWorksheet, sourceWorksheet) {
    const sourceDir = source.tempDir;
    const targetDir = target.tempDir;
    if (!fs.existsSync(`${targetDir}xl/worksheets/_rels/`))
        fs.mkdirSync(`${targetDir}xl/worksheets/_rels/`, { recursive: true }); //make worksheet rels directory if it doesnt exist yet.
    //copy worksheet rels file over
    const relList = [];
    const worksheetXMLRels = fs.readFileSync(`${sourceDir}xl/worksheets/_rels/${source.worksheets[sourceWorksheet].name}.xml.rels`, { encoding: 'utf-8' });
    xml2js.parseString(worksheetXMLRels, (error, editXML) => {
        editXML.Relationships.Relationship.forEach((rel) => {
            if (rel['$'].Target.includes(`../drawings/`)) {
                rel['$'].Target = `../drawings/${newDrawingName}.xml`;
                rel['$'].Id = rId;
                relList.push(rel);
            }
        });
        editXML.Relationships.Relationship = relList;
        const builder = new xml2js.Builder();
        const xml = builder.buildObject(editXML);
        fs.writeFileSync(`${targetDir}/xl/worksheets/_rels/${target.worksheets[targetWorksheet].name}.xml.rels`, xml);
    });
}
function addWorksheetDrawingTag(rId, newDrawingName, target, targetWorksheet) {
    const targetDir = target.tempDir;
    const worksheetXML = fs.readFileSync(`${targetDir}xl/worksheets/${target.worksheets[targetWorksheet].name}.xml`, { encoding: 'utf-8' });
    xml2js.parseString(worksheetXML, (error, editXML) => {
        editXML.worksheet.drawing = { $: { ['r:id']: rId } };
        // editXML.drawing['r:id'] = rId
        const builder = new xml2js.Builder();
        const xml = builder.buildObject(editXML);
        fs.writeFileSync(`${targetDir}/xl/worksheets/${target.worksheets[targetWorksheet].name}.xml`, xml);
    });
}
function newDrawingXML(source, target, sourceWorksheet, chartToCopy, targetWorksheet, rId, newDrawingName, contentTypesUpdateObj) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r, _s, _t, _u;
    const sourceDir = source.tempDir;
    const targetDir = target.tempDir;
    //update rId tag for sourceDrawingXML section
    const sourceDrawingRef = source.worksheets[sourceWorksheet].drawing;
    const drawingSource = source.drawingXMLs[sourceDrawingRef][chartToCopy]; //xml2Js object representing source drawing.xml sub section.
    const rIdRegular = (_j = (_h = (_g = (_f = (_e = (_d = (_c = (_b = (_a = drawingSource === null || drawingSource === void 0 ? void 0 : drawingSource['xdr:graphicFrame']) === null || _a === void 0 ? void 0 : _a[0]) === null || _b === void 0 ? void 0 : _b['a:graphic']) === null || _c === void 0 ? void 0 : _c[0]) === null || _d === void 0 ? void 0 : _d['a:graphicData']) === null || _e === void 0 ? void 0 : _e[0]) === null || _f === void 0 ? void 0 : _f['c:chart']) === null || _g === void 0 ? void 0 : _g[0]) === null || _h === void 0 ? void 0 : _h['$']) === null || _j === void 0 ? void 0 : _j['r:id']; //regular chart?.xml
    if (rIdRegular) { //update rID to match drawing?.xml.rels rId
        drawingSource['xdr:graphicFrame'][0]['a:graphic'][0]['a:graphicData'][0]['c:chart'][0]['$']['r:id'] = rId;
    }
    const rIdRegularEx = (_u = (_t = (_s = (_r = (_q = (_p = (_o = (_m = (_l = (_k = drawingSource === null || drawingSource === void 0 ? void 0 : drawingSource['mc:AlternateContent']) === null || _k === void 0 ? void 0 : _k[0]) === null || _l === void 0 ? void 0 : _l['mc:Choice']) === null || _m === void 0 ? void 0 : _m[0]) === null || _o === void 0 ? void 0 : _o['xdr:graphicFrame']) === null || _p === void 0 ? void 0 : _p[0]) === null || _q === void 0 ? void 0 : _q['a:graphic']) === null || _r === void 0 ? void 0 : _r[0]) === null || _s === void 0 ? void 0 : _s['a:graphicData']) === null || _t === void 0 ? void 0 : _t[0]) === null || _u === void 0 ? void 0 : _u['cx:chart'][0]['$']['r:id']; //alternate chart type chartEx?.xml
    if (rIdRegularEx) {
        drawingSource['mc:AlternateContent'][0]['mc:Choice'][0]['xdr:graphicFrame'][0]['a:graphic'][0]['a:graphicData'][0]['cx:chart'][0]['$']['r:id'] = rId;
    }
    //if drawing.xml does not exist for target worksheet, copy source drawing.xml and set Relationships.relation =  source.drawingXML
    //make sure to update drawingXML rId = new rID passed into function. File name should match new drawing name.
    //this cannot be a equal copy. Only one of the source drawing xml subsections needs to be copied over if new file.
    fs.copyFileSync(`${sourceDir}xl/drawings/${source.worksheets[sourceWorksheet].drawing}.xml`, `${targetDir}xl/drawings/${newDrawingName}.xml`);
    const drawingXML = fs.readFileSync(`${targetDir}xl/drawings/${newDrawingName}.xml`, { encoding: 'utf-8' });
    xml2js.parseString(drawingXML, (error, editXML) => {
        editXML['xdr:wsDr']['xdr:twoCellAnchor'] = drawingSource;
        const builder = new xml2js.Builder();
        const xml = builder.buildObject(editXML);
        fs.writeFileSync(`${targetDir}/xl/drawings/${newDrawingName}.xml`, xml);
        contentTypesUpdateObj[`/xl/drawings/${source.worksheets[sourceWorksheet].drawing}.xml`] = `/xl/drawings/${newDrawingName}.xml`;
    });
}
function updateDrawingXML(//if drawing.xml exists for target worksheet combine <xdr:twoCellAnchor> tags from source and target drawing file. New cellAnchor needs to have its rID updated.
source, target, sourceWorksheet, chartToCopy, targetWorksheet, rId, newDrawingName) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r, _s, _t, _u;
    const targetDir = target.tempDir;
    //update rId tag for sourceDrawingXML section
    const sourceDrawingRef = source.worksheets[sourceWorksheet].drawing;
    const drawingSource = source.drawingXMLs[sourceDrawingRef][chartToCopy]; //xml2Js object representing source drawing.xml sub section.
    const rIdRegular = (_j = (_h = (_g = (_f = (_e = (_d = (_c = (_b = (_a = drawingSource === null || drawingSource === void 0 ? void 0 : drawingSource['xdr:graphicFrame']) === null || _a === void 0 ? void 0 : _a[0]) === null || _b === void 0 ? void 0 : _b['a:graphic']) === null || _c === void 0 ? void 0 : _c[0]) === null || _d === void 0 ? void 0 : _d['a:graphicData']) === null || _e === void 0 ? void 0 : _e[0]) === null || _f === void 0 ? void 0 : _f['c:chart']) === null || _g === void 0 ? void 0 : _g[0]) === null || _h === void 0 ? void 0 : _h['$']) === null || _j === void 0 ? void 0 : _j['r:id']; //regular chart?.xml
    if (rIdRegular) { //update rID to match drawing?.xml.rels rId
        drawingSource['xdr:graphicFrame'][0]['a:graphic'][0]['a:graphicData'][0]['c:chart'][0]['$']['r:id'] = rId;
    }
    const rIdRegularEx = (_u = (_t = (_s = (_r = (_q = (_p = (_o = (_m = (_l = (_k = drawingSource === null || drawingSource === void 0 ? void 0 : drawingSource['mc:AlternateContent']) === null || _k === void 0 ? void 0 : _k[0]) === null || _l === void 0 ? void 0 : _l['mc:Choice']) === null || _m === void 0 ? void 0 : _m[0]) === null || _o === void 0 ? void 0 : _o['xdr:graphicFrame']) === null || _p === void 0 ? void 0 : _p[0]) === null || _q === void 0 ? void 0 : _q['a:graphic']) === null || _r === void 0 ? void 0 : _r[0]) === null || _s === void 0 ? void 0 : _s['a:graphicData']) === null || _t === void 0 ? void 0 : _t[0]) === null || _u === void 0 ? void 0 : _u['cx:chart'][0]['$']['r:id']; //alternate chart type chartEx?.xml
    if (rIdRegularEx) {
        drawingSource['mc:AlternateContent'][0]['mc:Choice'][0]['xdr:graphicFrame'][0]['a:graphic'][0]['a:graphicData'][0]['cx:chart'][0]['$']['r:id'] = rId;
    }
    const drawingXML = fs.readFileSync(`${targetDir}xl/drawings/${target.worksheets[targetWorksheet].drawing}.xml`, { encoding: 'utf-8' });
    xml2js.parseString(drawingXML, (error, editXML) => {
        //replace source drawing ref with new ref. Remember to update drawing ref in target. 
        editXML['xdr:wsDr']['xdr:twoCellAnchor'] = editXML['xdr:wsDr']['xdr:twoCellAnchor'].concat(drawingSource);
        const builder = new xml2js.Builder();
        const xml = builder.buildObject(editXML);
        fs.writeFileSync(`${targetDir}/xl/drawings/${target.worksheets[targetWorksheet].drawing}.xml`, xml); //
    });
}
function newDrawingRels(//if drawing.xml does not exist for target worksheet
source, target, sourceWorksheet, chartToCopy, targetWorksheet) {
    const rIdOutputList = []; //list of rIds in output drawing file.
    const sourceDir = source.tempDir;
    const targetDir = target.tempDir;
    const sourceDrawingName = source.worksheets[sourceWorksheet].drawing;
    const newChartName = getNewName(chartToCopy, target.chartList); //used for naming drawing.xml and drawing.xml.rels
    const newDrawingName = getNewName('drawing1', target.drawingList); //used for naming drawing.xml and drawing.xml.rels
    if (!fs.existsSync(`${targetDir}xl/drawings/_rels/`)) {
        fs.mkdirSync(`${targetDir}xl/drawings/_rels/`, { recursive: true }); //make drawing directory if it doesnt exist yet.
    }
    else { //if drawing directory exists and target worksheet has drawing file, read drawing file and update rID Output list so that we can find a new rId for drawing relation.
        const targetFile_Rels = `${targetDir}xl/drawings/_rels/${sourceDrawingName}.xml.rels`;
        if (fs.existsSync(targetFile_Rels)) {
            const drawingTargetSource = fs.readFileSync(targetFile_Rels, { encoding: 'utf-8' });
            xml2js.parseString(drawingTargetSource, (error, editXML) => {
                editXML.Relationships.Relationship.forEach((rel) => {
                    rIdOutputList.push(rel['$'].Id);
                });
            });
        }
    }
    let rId = getNewName('rId1', rIdOutputList);
    const drawingSourceRelsXML = fs.readFileSync(`${sourceDir}xl/drawings/_rels/${sourceDrawingName}.xml.rels`, { encoding: 'utf-8' }); //`${targetDir}xl/drawings/${drawingName}.xml`
    xml2js.parseString(drawingSourceRelsXML, (error, editXML) => {
        editXML.Relationships.Relationship.forEach((rel) => {
            const refChartName = rel['$'].Target.replace('../charts/', '').replace('.xml', '');
            if (refChartName === chartToCopy) {
                rel['$'].Target = `../charts/${newChartName}.xml`;
                rel['$'].Id = rId;
                target.worksheets[targetWorksheet][newChartName] = rId;
                editXML.Relationships.Relationship = [rel]; //if match, create file with single relationship, representing new chart. rId can stay the same.
            }
        });
        const builder = new xml2js.Builder();
        const xml = builder.buildObject(editXML);
        fs.writeFileSync(`${targetDir}/xl/drawings/_rels/${newDrawingName}.xml.rels`, xml);
    });
    return [rId, newChartName, newDrawingName];
}
function updateDrawingRels(//if drawing.xml exists for target worksheet combine <xdr:twoCellAnchor> tags from source and target drawing file. Update rId and ChartName
source, target, sourceWorksheet, chartToCopy, targetWorksheet) {
    let rId = '';
    const sourceDir = source.tempDir;
    const targetDir = target.tempDir;
    const sourceDrawingName = source.worksheets[sourceWorksheet].drawing;
    const drawingSourceRelsXML = fs.readFileSync(`${sourceDir}xl/drawings/_rels/${sourceDrawingName}.xml.rels`, { encoding: 'utf-8' }); //`${targetDir}xl/drawings/${drawingName}.xml`
    const newChartName = getNewName(chartToCopy, target.chartList); //used for naming drawing.xml and drawing.xml.rels
    let sourceRelTag;
    xml2js.parseString(drawingSourceRelsXML, (error, editXML) => {
        editXML.Relationships.Relationship.forEach((rel) => {
            const refChartName = rel['$'].Target.replace('../charts/', '').replace('.xml', '');
            if (refChartName === chartToCopy) {
                rId = getNewName('rId1', Object.values(target.worksheets[targetWorksheet].drawingRels));
                target.worksheets[targetWorksheet][newChartName] = rId;
                sourceRelTag = rel;
                sourceRelTag['$'].Id = rId;
                sourceRelTag['$'].Target = `../charts/${newChartName}.xml`;
            }
        });
    });
    const targetName = target.worksheets[targetWorksheet].drawing;
    const drawingTargetPath = `${targetDir}xl/drawings/_rels/${targetName}.xml.rels`;
    const drawingTargetRelsXML = fs.readFileSync(drawingTargetPath, { encoding: 'utf-8' });
    xml2js.parseString(drawingTargetRelsXML, (error, editXML) => {
        editXML.Relationships.Relationship = editXML.Relationships.Relationship.concat(sourceRelTag);
        const builder = new xml2js.Builder();
        const xml = builder.buildObject(editXML);
        fs.writeFileSync(`${targetDir}/xl/drawings/_rels/${target.worksheets[targetWorksheet].drawing}.xml.rels`, xml);
    });
    return [rId, newChartName, ''];
}
export function copyChart(sourceExcel, //chart source object returned from readCharts. Includes chart details and source xml directory
targetExcel, //target excel object returned from readCharts. Includes chart details and source xml directory
sourceWorksheet, //alias of source worksheet
chartToCopy, //chart, from chartDetails, that is copied by this operation
targetWorksheet, //alias of sheet that chart will be copied to. Alias is the sheet name visable to an ecxel user.
stringOverrides) {
    const contentTypesUpdateObj = {}; //partNameSource : partNameOutput
    if (!targetExcel.worksheets[targetWorksheet].drawing) { //if no drawing for target worksheet.
        //add chart tag to worksheet
        console.log('1');
        const [rId, newChartName, newDrawingName] = newDrawingRels(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet);
        console.log('2');
        newDrawingXML(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet, rId, newDrawingName, contentTypesUpdateObj);
        console.log('3');
        addWorksheetDrawingTag(rId, newDrawingName, targetExcel, targetWorksheet);
        console.log('4');
        addWorksheetRelsFile(rId, newDrawingName, targetExcel, sourceExcel, targetWorksheet, sourceWorksheet);
        console.log('5');
        const newDefinedNamesObj = copyChartFiles(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, newChartName, stringOverrides, contentTypesUpdateObj);
        console.log('6');
        copyDefineNames(sourceExcel, targetExcel, newDefinedNamesObj, stringOverrides);
        console.log('7');
        updateContentTypes(contentTypesUpdateObj, sourceExcel, targetExcel, sourceWorksheet, chartToCopy, newChartName);
    }
    else {
        console.log('1a');
        const [rId, newChartName, newDrawingName] = updateDrawingRels(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet);
        console.log('2a');
        updateDrawingXML(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet, rId, newDrawingName);
        console.log('3a');
        const newDefinedNamesObj = copyChartFiles(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, newChartName, stringOverrides, contentTypesUpdateObj);
        console.log('4a');
        copyDefineNames(sourceExcel, targetExcel, newDefinedNamesObj, stringOverrides);
        console.log('5a');
        updateContentTypes(contentTypesUpdateObj, sourceExcel, targetExcel, sourceWorksheet, chartToCopy, newChartName);
    }
}
//# sourceMappingURL=copyChart.js.map