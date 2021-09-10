import fs from 'fs';
import xml2js from 'xml2js';

import { workbookChartDetails, worksheetObj } from 'readCharts'

interface stringOverrides {
    [key: string]: string
}

function copyDefineNames( //
    sourceExcel: workbookChartDetails,
    sourceWorksheet: string,
    chartToCopy: string,
    targetExcel: workbookChartDetails,
    targetWorksheet: string,
    newChartName: string,
    stringOverrides: stringOverrides,
    newDefinedNamesRefsObj: Object
) {

    if (Object.keys(sourceExcel.worksheets[sourceWorksheet].charts[chartToCopy].definedNameRefs).length > 0) { //copy defineNamed refs from source workbook.xml to target workbook.xml

        const sourceDir = sourceExcel.tempDir
        const targetDir = targetExcel.tempDir

        let newRelList: string[] = []

        let addDefs: any[] = []
        const sourceWookbook = `${sourceDir}xl/workbook.xml`
        const sourceXML = fs.readFileSync(sourceWookbook, { encoding: 'utf-8' })
        xml2js.parseString(sourceXML, (error, editXML) => { //read source workbook
            editXML.workbook.definedNames[0].definedName.forEach((rel) => {  //if source defineName in newDefinedNameObj, update definename.name and push to update list.
                if (newDefinedNamesRefsObj[rel['$'].name]) {
                    const newRefValue = stringOverrides[rel['_']]
                    if (newRefValue) newRelList.push(newRefValue)

                    rel['_'] = `'${stringOverrides[rel['_']]}`.replace("!", "'!")
                    rel['$'].name = newDefinedNamesRefsObj[rel['$'].name]
                    addDefs.push(rel)
                }
                // if (stringOverrides[rel['_']]) {
                //     const newRefValue = stringOverrides[rel['_']]
                //     rel['_'] = stringOverrides[rel['_']]
                //     newRelList.push(newRefValue)
                //     addDefs.push(rel)
                // }
            })
        })
        console.log('newRelList', newRelList, 'addDefs', addDefs)

        const outputWorkbook = `${targetDir}xl/workbook.xml`
        const outputFile = fs.readFileSync(outputWorkbook, { encoding: 'utf-8' })
        xml2js.parseString(outputFile, (error, editXML) => { //read source workbook
            if (editXML.workbook.definedNames) {
                editXML.workbook.definedNames[0].definedName.concat(addDefs)
            } else { //need to copy from old xml, and insure that definedNames follows </sheets> tag
                const newWorkbookObj = {}
                Object.entries(editXML.workbook).forEach(([key, val]) => {
                    if (key !== 'sheets') {
                        newWorkbookObj[key] = val
                    } else {
                        newWorkbookObj[key] = val
                        newWorkbookObj['definedNames'] = [{ definedName: addDefs }]
                    }
                })
                editXML.workbook = newWorkbookObj
            }
            const builder = new xml2js.Builder()
            const xml = builder.buildObject(editXML)

            targetExcel.defineNameRefs = [...targetExcel.defineNameRefs, ...newRelList]
            targetExcel.worksheets[targetWorksheet].charts[newChartName].definedNameRefs = newRelList

            fs.writeFileSync(`${targetDir}/xl/workbook.xml`, xml)
        })
    }
}

function updateContentTypes(
    contentTypesUpdateObj: { [key: string]: string },
    sourceExcel: workbookChartDetails,
    targetExcel: workbookChartDetails,
    sourceWorksheet: string,
    chartToCopy: string,
    newChartName: string,
) {

    const sourceDir = sourceExcel.tempDir
    const targetDir = targetExcel.tempDir

    const updateTags: any[] = []
    //read source contentTypes, copy source tags to list after updating PartName
    const sourceContent = fs.readFileSync(`${sourceDir}/[Content_Types].xml`, { encoding: 'utf-8' })
    xml2js.parseString(sourceContent, (error, editXML) => {

        editXML.Types.Override.forEach((rel) => {  //update rels with new chart name
            if (contentTypesUpdateObj[rel['$'].PartName]) {
                rel['$'].PartName = contentTypesUpdateObj[rel['$'].PartName]
                updateTags.push(rel)
            }
        })
    })



    //update output contentTypes
    const targetContent = fs.readFileSync(`${targetDir}/[Content_Types].xml`, { encoding: 'utf-8' })
    xml2js.parseString(targetContent, (error, editXML) => {
        editXML.Types.Override = editXML.Types.Override.concat(updateTags)

        const builder = new xml2js.Builder()
        const xml = builder.buildObject(editXML)
        fs.writeFileSync(`${targetDir}/[Content_Types].xml`, xml)
    })

}

function getNewName(newName: string, targetList: string[], iterator = 0): string {
    if (!targetList.includes(newName)) {
        targetList.push(newName)
        return newName
    } else {
        const updateIterator = iterator + 1
        const updateNewName = `${newName.replace(new RegExp("[0-9]", "g"), "")}${updateIterator}`
        return getNewName(updateNewName, targetList, updateIterator)
    }
}

function getNewDefinedNameRef(newName: string, targetList: string[], iterator = new Date().getTime()): string {
    const testName = `${newName}.${iterator}`
    if (!targetList.includes(testName)) {
        targetList.push(testName)
        return testName
    } else {
        const updateIterator = iterator + 1
        // const updateNewName = newName.slice(0, newName.lastIndexOf('.')) + iterator
        return getNewDefinedNameRef(newName, targetList, updateIterator)
    }
}


function findChartCellRefs(xml: string, worksheetList: string[]): string[] {

    let tempRefs: any = []
    worksheetList.forEach((worksheet) => {

        let worksheetCellRefRange = `${worksheet}!\\$[A-Z]{1,3}\\$[0-9]{1,7}:\\$[A-Z]{1,3}\\$[0-9]{1,7}`
        let findCellRef = new RegExp(worksheetCellRefRange, 'g')
        let matchListNoCommas = [...new Set(xml.match(findCellRef))]
        //check for matching WITH commas around names 
        worksheetCellRefRange = `'${worksheet}'!\\$[A-Z]{1,3}\\$[0-9]{1,7}:\\$[A-Z]{1,3}\\$[0-9]{1,7}`
        findCellRef = new RegExp(worksheetCellRefRange, 'g')
        let matchListCommas = [...new Set(xml.match(findCellRef))]

        tempRefs = [...new Set(tempRefs.concat(matchListNoCommas.concat(matchListCommas)))]
        //find matching cell refs.
        let worksheetCellRef = `${worksheet}!\\$[A-Z]{1,3}\\$[0-9]{1,7}<`
        let findCellRefCell = new RegExp(worksheetCellRef, 'g')
        let matchListCellNoCommas = [...new Set(xml.match(findCellRefCell))]
        matchListCellNoCommas = matchListCellNoCommas.map(el => el.slice(0, el.length - 1))
        //check for matching WITH commas around names 
        worksheetCellRef = `'${worksheet}'!\\$[A-Z]{1,3}\\$[0-9]{1,7}<`
        findCellRefCell = new RegExp(worksheetCellRef, 'g')
        let matchListCellCommas = [...new Set(xml.match(findCellRefCell))]
        matchListCellCommas = matchListCellCommas.map(el => el.slice(0, el.length - 1))

        tempRefs = [...new Set(tempRefs.concat(matchListCellNoCommas).concat(matchListCellCommas))]

        tempRefs = [...new Set(tempRefs.concat(tempRefs))]
    })

    return tempRefs
}

function copyChartFiles(
    sourceExcel: workbookChartDetails,
    targetExcel: workbookChartDetails,
    sourceWorksheet: string,
    chartToCopy: string,
    newChartName: string,
    stringOverrides: stringOverrides,
    contentTypesUpdateObj: { [key: string]: string },
    targetWorksheet: string,
) {

    const sourceDir = sourceExcel.tempDir
    const targetDir = targetExcel.tempDir

    if (!fs.existsSync(`${targetDir}xl/charts/_rels/`)) fs.mkdirSync(`${targetDir}xl/charts/_rels/`, { recursive: true }) //make directory if needed

    const getNewColorsFileName = getNewName('colors1', targetExcel.colorList)
    const getNewStyleFileName = getNewName('style1', targetExcel.styleList)


    //COPY SOURCE RELS FILE
    const sourceRelsFile = `${sourceDir}xl/charts/_rels/${chartToCopy}.xml.rels`
    const sourceRelsXML = fs.readFileSync(sourceRelsFile, { encoding: 'utf-8' })
    xml2js.parseString(sourceRelsXML, (error, editXML) => { //read list of related files.
        editXML.Relationships.Relationship.forEach((rel) => {  //update rels with new chart name
            if (rel['$'].Target.includes('colors')) rel['$'].Target = `${getNewColorsFileName}.xml`
            if (rel['$'].Target.includes('style')) rel['$'].Target = `${getNewStyleFileName}.xml`
        })
        const builder = new xml2js.Builder()
        const xml = builder.buildObject(editXML)
        fs.writeFileSync(`${targetDir}/xl/charts/_rels/${newChartName}.xml.rels`, xml) //write rels file
    })

    //COPY SOURCE CHART FILE. Update any cell references refs (ex. A1:B2) AND definedName name refs.(ex. '_xlchart.v1.0')
    const sourceChartFile = `${sourceDir}xl/charts/${chartToCopy}.xml`
    let sourceChartXML = fs.readFileSync(sourceChartFile, { encoding: 'utf-8' })

    Object.entries(stringOverrides).forEach(([key, val]) => { //replace all cell references with overrides.
        const newKey = key.replace(/\$/g, '\\$')
        const regExKey = new RegExp(`>${newKey}<`, 'g')
        sourceChartXML = sourceChartXML.replace(regExKey, `>${val}<`)
    })

    // create definedName ref object. {Old ref: new ref}
    let refRegex = new RegExp(`>_xlchart.v[0-9]{1,9}.[0-9]{1,10}<`, 'g')
    const foundDefineNameRefs: string[] = [...new Set(sourceChartXML.match(refRegex))]

    const newDefinedNamesRefsObj = foundDefineNameRefs.reduce((acc, el) => {
        const oldName = el.slice(1, el.length - 1)
        const newDefinedNameRef = getNewDefinedNameRef('_xlchart.v1', targetExcel.defineNameRefs)
        console.log('REPLACE', el, `>${newDefinedNameRef}<`)
        sourceChartXML = sourceChartXML.replace(new RegExp(`${el}`, 'g'), `>${newDefinedNameRef}<`) //override source reference with new reference.
        return { ...acc, [oldName]: newDefinedNameRef }
    }, {})

    fs.writeFileSync(`${targetDir}/xl/charts/${newChartName}.xml`, sourceChartXML)
    contentTypesUpdateObj[`/xl/charts/${chartToCopy}.xml`] = `/xl/charts/${newChartName}.xml`

    //COPY Chart colors?.xml and style?.xml
    Object.entries(sourceExcel.worksheets[sourceWorksheet].charts[chartToCopy].chartRels).forEach(([key, val]) => {
        // const updateFileName = `${key}${newChartName.replace(/[A-z]/g, '')}.xml`
        const thisFileName = key === 'colors' ? getNewColorsFileName : getNewStyleFileName
        fs.copyFileSync(`${sourceDir}xl/charts/${val}.xml`, `${targetDir}xl/charts/${thisFileName}.xml`)
        contentTypesUpdateObj[`/xl/charts/${val}.xml`] = `/xl/charts/${thisFileName}.xml`
    })

    const newExcelCellRefs = findChartCellRefs(sourceChartXML, [...Object.keys(sourceExcel.worksheets), ...Object.keys(targetExcel.worksheets)])

    const newChartObj = {
        chartRels: {
            colors: getNewColorsFileName,
            style: getNewStyleFileName
        },
        cellRefs: newExcelCellRefs,
        definedNameRefs: [],
    }

    if (!targetExcel.worksheets[targetWorksheet]?.charts) targetExcel.worksheets[targetWorksheet].charts = {}
    targetExcel.worksheets[targetWorksheet].charts[newChartName] = newChartObj

    return newDefinedNamesRefsObj
}

function addWorksheetRelsFile(
    rId: string,
    newDrawingName: string,
    target: workbookChartDetails,
    source: workbookChartDetails,
    targetWorksheet: string,
    sourceWorksheet: string,
) {
    const sourceDir = source.tempDir
    const targetDir = target.tempDir

    if (!fs.existsSync(`${targetDir}xl/worksheets/_rels/`)) fs.mkdirSync(`${targetDir}xl/worksheets/_rels/`, { recursive: true }) //make worksheet rels directory if it doesnt exist yet.
    //copy worksheet rels file over
    const relList: any[] = []
    const worksheetXMLRels = fs.readFileSync(`${sourceDir}xl/worksheets/_rels/${source.worksheets[sourceWorksheet].name}.xml.rels`, { encoding: 'utf-8' })

    xml2js.parseString(worksheetXMLRels, (error, editXML) => {
        editXML.Relationships.Relationship.forEach((rel) => {
            if (rel['$'].Target.includes(`../drawings/`)) {
                rel['$'].Target = `../drawings/${newDrawingName}.xml`
                rel['$'].Id = rId
                relList.push(rel)
            }
        })
        editXML.Relationships.Relationship = relList
        const builder = new xml2js.Builder()
        const xml = builder.buildObject(editXML)
        fs.writeFileSync(`${targetDir}/xl/worksheets/_rels/${target.worksheets[targetWorksheet].name}.xml.rels`, xml)
    })
}

function addWorksheetDrawingTag(
    rId: string,
    newDrawingName: string,
    target: workbookChartDetails,
    targetWorksheet: string,
) {
    const targetDir = target.tempDir
    const worksheetXML = fs.readFileSync(`${targetDir}xl/worksheets/${target.worksheets[targetWorksheet].name}.xml`, { encoding: 'utf-8' })
    xml2js.parseString(worksheetXML, (error, editXML) => {
        editXML.worksheet.drawing = { $: { ['r:id']: rId } }
        // editXML.drawing['r:id'] = rId
        const builder = new xml2js.Builder()
        const xml = builder.buildObject(editXML)
        fs.writeFileSync(`${targetDir}/xl/worksheets/${target.worksheets[targetWorksheet].name}.xml`, xml)
    })
}

function newDrawingXML( //if no drawing exists for target worksheet then the source file needs to be copied, with only a relation to the target chart.
    source: workbookChartDetails,
    target: workbookChartDetails,
    sourceWorksheet: string,
    chartToCopy: string,
    targetWorksheet: string,
    rId: string,
    newDrawingName: string,
    contentTypesUpdateObj: { [key: string]: string }
) {
    const sourceDir = source.tempDir
    const targetDir = target.tempDir
    //update rId tag for sourceDrawingXML section
    const sourceDrawingRef: string = source.worksheets[sourceWorksheet].drawing
    const drawingSource = source.drawingXMLs[sourceDrawingRef][chartToCopy] //xml2Js object representing source drawing.xml sub section.

    const rIdRegular = drawingSource?.['xdr:graphicFrame']?.[0]?.['a:graphic']?.[0]?.['a:graphicData']?.[0]?.['c:chart']?.[0]?.['$']?.['r:id'] //regular chart?.xml
    if (rIdRegular) { //update rID to match drawing?.xml.rels rId
        drawingSource['xdr:graphicFrame'][0]['a:graphic'][0]['a:graphicData'][0]['c:chart'][0]['$']['r:id'] = rId
    }
    const rIdRegularEx = drawingSource?.['mc:AlternateContent']?.[0]?.['mc:Choice']?.[0]?.['xdr:graphicFrame']?.[0]?.['a:graphic']?.[0]?.['a:graphicData']?.[0]?.['cx:chart'][0]['$']['r:id'] //alternate chart type chartEx?.xml
    if (rIdRegularEx) {
        drawingSource['mc:AlternateContent'][0]['mc:Choice'][0]['xdr:graphicFrame'][0]['a:graphic'][0]['a:graphicData'][0]['cx:chart'][0]['$']['r:id'] = rId
    }

    //if drawing.xml does not exist for target worksheet, copy source drawing.xml and set Relationships.relation =  source.drawingXML
    //make sure to update drawingXML rId = new rID passed into function. File name should match new drawing name.
    //this cannot be a equal copy. Only one of the source drawing xml subsections needs to be copied over if new file.
    fs.copyFileSync(`${sourceDir}xl/drawings/${source.worksheets[sourceWorksheet].drawing}.xml`, `${targetDir}xl/drawings/${newDrawingName}.xml`)
    const drawingXML = fs.readFileSync(`${targetDir}xl/drawings/${newDrawingName}.xml`, { encoding: 'utf-8' })
    xml2js.parseString(drawingXML, (error, editXML) => {
        editXML['xdr:wsDr']['xdr:twoCellAnchor'] = drawingSource
        const builder = new xml2js.Builder()
        const xml = builder.buildObject(editXML)
        fs.writeFileSync(`${targetDir}/xl/drawings/${newDrawingName}.xml`, xml)
        contentTypesUpdateObj[`/xl/drawings/${source.worksheets[sourceWorksheet].drawing}.xml`] = `/xl/drawings/${newDrawingName}.xml`
    })

}

function updateDrawingXML(         //if drawing.xml exists for target worksheet combine <xdr:twoCellAnchor> tags from source and target drawing file. New cellAnchor needs to have its rID updated.
    source: workbookChartDetails,
    target: workbookChartDetails,
    sourceWorksheet: string,
    chartToCopy: string,
    targetWorksheet: string,
    rId: string,
    newDrawingName: string,
) {
    const targetDir = target.tempDir

    //update rId tag for sourceDrawingXML section
    const sourceDrawingRef: string = source.worksheets[sourceWorksheet].drawing
    const drawingSource = source.drawingXMLs[sourceDrawingRef][chartToCopy] //xml2Js object representing source drawing.xml sub section.

    const rIdRegular = drawingSource?.['xdr:graphicFrame']?.[0]?.['a:graphic']?.[0]?.['a:graphicData']?.[0]?.['c:chart']?.[0]?.['$']?.['r:id'] //regular chart?.xml
    if (rIdRegular) { //update rID to match drawing?.xml.rels rId
        drawingSource['xdr:graphicFrame'][0]['a:graphic'][0]['a:graphicData'][0]['c:chart'][0]['$']['r:id'] = rId
    }
    const rIdRegularEx = drawingSource?.['mc:AlternateContent']?.[0]?.['mc:Choice']?.[0]?.['xdr:graphicFrame']?.[0]?.['a:graphic']?.[0]?.['a:graphicData']?.[0]?.['cx:chart'][0]['$']['r:id'] //alternate chart type chartEx?.xml
    if (rIdRegularEx) {
        drawingSource['mc:AlternateContent'][0]['mc:Choice'][0]['xdr:graphicFrame'][0]['a:graphic'][0]['a:graphicData'][0]['cx:chart'][0]['$']['r:id'] = rId
    }

    const drawingXML = fs.readFileSync(`${targetDir}xl/drawings/${target.worksheets[targetWorksheet].drawing}.xml`, { encoding: 'utf-8' })
    xml2js.parseString(drawingXML, (error, editXML) => {
        //replace source drawing ref with new ref. Remember to update drawing ref in target. 
        editXML['xdr:wsDr']['xdr:twoCellAnchor'] = editXML['xdr:wsDr']['xdr:twoCellAnchor'].concat(drawingSource)
        const builder = new xml2js.Builder()
        const xml = builder.buildObject(editXML)
        fs.writeFileSync(`${targetDir}/xl/drawings/${target.worksheets[targetWorksheet].drawing}.xml`, xml) //
    })

}

function newDrawingRels( //if drawing.xml does not exist for target worksheet
    source: workbookChartDetails,
    target: workbookChartDetails,
    sourceWorksheet: string,
    chartToCopy: string,
    targetWorksheet: string,
): [string, string, string] {

    const rIdOutputList: string[] = [] //list of rIds in output drawing file.
    const sourceDir = source.tempDir
    const targetDir = target.tempDir
    const sourceDrawingName = source.worksheets[sourceWorksheet].drawing
    const newChartName: string = getNewName(chartToCopy, target.chartList) //used for naming drawing.xml and drawing.xml.rels

    const newDrawingName = getNewName('drawing1', target.drawingList) //used for naming drawing.xml and drawing.xml.rels
    target.worksheets[targetWorksheet].drawing = newDrawingName
    if (!fs.existsSync(`${targetDir}xl/drawings/_rels/`)) {
        fs.mkdirSync(`${targetDir}xl/drawings/_rels/`, { recursive: true }) //make drawing directory if it doesnt exist yet.
    } else { //if drawing directory exists and target worksheet has drawing file, read drawing file and update rID Output list so that we can find a new rId for drawing relation.
        const targetFile_Rels = `${targetDir}xl/drawings/_rels/${sourceDrawingName}.xml.rels`
        if (fs.existsSync(targetFile_Rels)) {
            const drawingTargetSource = fs.readFileSync(targetFile_Rels, { encoding: 'utf-8' })
            xml2js.parseString(drawingTargetSource, (error, editXML) => {
                editXML.Relationships.Relationship.forEach((rel) => {
                    rIdOutputList.push(rel['$'].Id)
                })
            })
        }
    }

    let rId: string = getNewName('rId1', rIdOutputList)

    const drawingSourceRelsXML = fs.readFileSync(`${sourceDir}xl/drawings/_rels/${sourceDrawingName}.xml.rels`, { encoding: 'utf-8' }) //`${targetDir}xl/drawings/${drawingName}.xml`
    xml2js.parseString(drawingSourceRelsXML, (error, editXML) => {
        editXML.Relationships.Relationship.forEach((rel) => {
            const refChartName = rel['$'].Target.replace('../charts/', '').replace('.xml', '')
            if (refChartName === chartToCopy) {
                rel['$'].Target = `../charts/${newChartName}.xml`
                rel['$'].Id = rId

                editXML.Relationships.Relationship = [rel] //if match, create file with single relationship, representing new chart. rId can stay the same.
            }
        })
        const builder = new xml2js.Builder()
        const xml = builder.buildObject(editXML)
        fs.writeFileSync(`${targetDir}/xl/drawings/_rels/${newDrawingName}.xml.rels`, xml)

    })

    target.worksheets[targetWorksheet].drawingRels = { [newChartName]: rId }

    return [rId, newChartName, newDrawingName]

}

function updateDrawingRels(  //if drawing.xml exists for target worksheet combine <xdr:twoCellAnchor> tags from source and target drawing file. Update rId and ChartName
    source: workbookChartDetails,
    target: workbookChartDetails,
    sourceWorksheet: string,
    chartToCopy: string,
    targetWorksheet: string,
): [string, string, string] {
    let rId: string = ''
    const sourceDir = source.tempDir
    const targetDir = target.tempDir
    const sourceDrawingName = source.worksheets[sourceWorksheet].drawing
    const drawingSourceRelsXML = fs.readFileSync(`${sourceDir}xl/drawings/_rels/${sourceDrawingName}.xml.rels`, { encoding: 'utf-8' }) //`${targetDir}xl/drawings/${drawingName}.xml`
    const newChartName: string = getNewName(chartToCopy, target.chartList) //used for naming drawing.xml and drawing.xml.rels

    let sourceRelTag
    xml2js.parseString(drawingSourceRelsXML, (error, editXML) => { //make a copy of the source relationship tag after updating rId & target.
        editXML.Relationships.Relationship.forEach((rel) => {
            const refChartName = rel['$'].Target.replace('../charts/', '').replace('.xml', '')
            if (refChartName === chartToCopy) {
                rId = getNewName('rId1', Object.values(target.worksheets[targetWorksheet].drawingRels))
                target.worksheets[targetWorksheet][newChartName] = rId
                sourceRelTag = rel
                sourceRelTag['$'].Id = rId
                sourceRelTag['$'].Target = `../charts/${newChartName}.xml`
            }
        })
    })
    const targetName = target.worksheets[targetWorksheet].drawing
    const drawingTargetPath = `${targetDir}xl/drawings/_rels/${targetName}.xml.rels`
    const drawingTargetRelsXML = fs.readFileSync(drawingTargetPath, { encoding: 'utf-8' })
    xml2js.parseString(drawingTargetRelsXML, (error, editXML) => { //insert new relations tag into drawing?.xml.rel
        editXML.Relationships.Relationship = editXML.Relationships.Relationship.concat(sourceRelTag)
        const builder = new xml2js.Builder()
        const xml = builder.buildObject(editXML)
        fs.writeFileSync(`${targetDir}/xl/drawings/_rels/${target.worksheets[targetWorksheet].drawing}.xml.rels`, xml)
    })

    target?.worksheets?.[targetWorksheet]?.drawingRels?.[newChartName] ?
        target.worksheets[targetWorksheet].drawingRels[newChartName] = rId :
        target.worksheets[targetWorksheet].drawingRels = { [newChartName]: rId }

    return [rId, newChartName, '']
}

export function copyChart(
    sourceExcel: workbookChartDetails, //chart source object returned from readCharts. Includes chart details and source xml directory
    targetExcel: workbookChartDetails, //target excel object returned from readCharts. Includes chart details and source xml directory
    sourceWorksheet: string, //alias of source worksheet
    chartToCopy: string, //chart, from chartDetails, that is copied by this operation
    targetWorksheet: string, //alias of sheet that chart will be copied to. Alias is the sheet name visable to an ecxel user.
    stringOverrides: stringOverrides, //list of source worksheet cell references that need to be replaced. ex: {[worksheet1!A1:B2] : newWorksheet!A1:B2} 
) {
    const contentTypesUpdateObj = {} //partNameSource : partNameOutput

    if (!targetExcel.worksheets[targetWorksheet].drawing) { //if no drawing for target worksheet.
        //add chart tag to worksheet
        const [rId, newChartName, newDrawingName] = newDrawingRels(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet)
        newDrawingXML(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet, rId, newDrawingName, contentTypesUpdateObj)
        addWorksheetDrawingTag(rId, newDrawingName, targetExcel, targetWorksheet)
        addWorksheetRelsFile(rId, newDrawingName, targetExcel, sourceExcel, targetWorksheet, sourceWorksheet)
        const newDefinedNamesRefsObj = copyChartFiles(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, newChartName, stringOverrides, contentTypesUpdateObj, targetWorksheet)
        copyDefineNames(sourceExcel, sourceWorksheet, chartToCopy, targetExcel, targetWorksheet, newChartName, stringOverrides, newDefinedNamesRefsObj)
        updateContentTypes(contentTypesUpdateObj, sourceExcel, targetExcel, sourceWorksheet, chartToCopy, newChartName)
    } else {
        const [rId, newChartName, newDrawingName] = updateDrawingRels(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet)
        updateDrawingXML(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet, rId, newDrawingName)
        const newDefinedNamesRefsObj = copyChartFiles(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, newChartName, stringOverrides, contentTypesUpdateObj, targetWorksheet)
        copyDefineNames(sourceExcel, sourceWorksheet, chartToCopy, targetExcel, targetWorksheet, newChartName, stringOverrides, newDefinedNamesRefsObj)
        updateContentTypes(contentTypesUpdateObj, sourceExcel, targetExcel, sourceWorksheet, chartToCopy, newChartName)
    }
}