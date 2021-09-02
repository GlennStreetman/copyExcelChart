import fs from 'fs';
import xml2js from 'xml2js';
import AdmZip from 'adm-zip';

import { workbookChartDetails, worksheetObj } from 'readCharts'

interface stringOverrides {
    [key: string]: string
}

function getNewName(newName: string, targetList: string[], iterator = 0): string {
    console.log(newName)
    if (!targetList.includes(newName)) {
        targetList.push(newName)
        return newName
    } else {
        const updateIterator = iterator + 1
        const updateNewName = `${newName.replace(new RegExp("[0-9]", "g"), "")}${updateIterator}`
        return getNewName(updateNewName, targetList, updateIterator)
    }
}

function copyChartFiles(
    sourceExcel: workbookChartDetails,
    targetExcel: workbookChartDetails,
    sourceWorksheet: string,
    chartToCopy: string,
    newChartName: string,
    stringOverrides: stringOverrides) {
    //copy all four files, update names

    //update string overrides for chart.xml

    const sourceDir = sourceExcel.tempDir
    const targetDir = targetExcel.tempDir

    if (!fs.existsSync(`${targetDir}xl/charts/_rels/`)) fs.mkdirSync(`${targetDir}xl/charts/_rels/`, { recursive: true })

    const sourceRelsFile = `${sourceDir}xl/charts/_rels/${chartToCopy}.xml.rels`
    const sourceRelsXML = fs.readFileSync(sourceRelsFile, { encoding: 'utf-8' })
    xml2js.parseString(sourceRelsXML, (error, editXML) => {
        editXML.Relationships.Relationship.forEach((rel) => {  //update rels with new chart name
            if (rel['$'].Target.includes('colors')) rel['$'].Target = `colors${newChartName.replace(/[A-z]/g, '')}.xml`
            if (rel['$'].Target.includes('style')) rel['$'].Target = `style${newChartName.replace(/[A-z]/g, '')}.xml`
        })
        const builder = new xml2js.Builder()
        const xml = builder.buildObject(editXML)
        fs.writeFileSync(`${targetDir}/xl/charts/_rels/${newChartName}.xml.rels`, xml)
    })

    const sourceChartFile = `${sourceDir}xl/charts/${chartToCopy}.xml`
    let sourceChartXML = fs.readFileSync(sourceChartFile, { encoding: 'utf-8' })
    Object.entries(stringOverrides).forEach(([key, val]) => { //copy source chart fil, update cell references.
        sourceChartXML = sourceChartXML.replace(key, val)
    })
    fs.writeFileSync(`${targetDir}/xl/charts/${newChartName}.xml`, sourceChartXML)

    console.log('LAST', sourceExcel.worksheets[sourceWorksheet].charts[chartToCopy].chartRels)
    Object.entries(sourceExcel.worksheets[sourceWorksheet].charts[chartToCopy].chartRels).forEach(([key, val]) => {
        const updateFileName = `${key}${newChartName.replace(/[A-z]/g, '')}.xml`
        fs.copyFileSync(`${sourceDir}xl/charts/${val}.xml`, `${targetDir}xl/charts/${updateFileName}`)
    })

    //copy relations
}



function newDrawingXML(
    source: workbookChartDetails,
    target: workbookChartDetails,
    sourceWorksheet: string,
    chartToCopy: string,
    targetWorksheet: string,
    rId: string,
    newDrawingName: string,
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

    console.log('drawingSource', drawingSource, 'writing: ', target.worksheets[targetWorksheet].drawing)
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
    let rId: string = 'rId1'
    const sourceDir = source.tempDir
    const targetDir = target.tempDir
    const sourceDrawingName = source.worksheets[sourceWorksheet].drawing
    const drawingSourceRelsXML = fs.readFileSync(`${sourceDir}xl/drawings/_rels/${sourceDrawingName}.xml.rels`, { encoding: 'utf-8' }) //`${targetDir}xl/drawings/${drawingName}.xml`
    const newChartName: string = getNewName(chartToCopy, target.chartList) //used for naming drawing.xml and drawing.xml.rels

    const newDrawingName = getNewName('drawing1', target.drawingList) //used for naming drawing.xml and drawing.xml.rels
    if (!fs.existsSync(`${targetDir}xl/drawings/_rels/`)) fs.mkdirSync(`${targetDir}xl/drawings/_rels/`, { recursive: true }) //make drawing directory if it doesnt exist yet.

    xml2js.parseString(drawingSourceRelsXML, (error, editXML) => {
        editXML.Relationships.Relationship.forEach((rel) => {
            const refChartName = rel['$'].Target.replace('../charts/', '').replace('.xml', '')
            if (refChartName === chartToCopy) {
                console.log('FOUND MATCHING ', refChartName, chartToCopy)
                rel['$'].Target = `../charts/${newChartName}.xml`
                rel['$'].Id = rId
                target.worksheets[targetWorksheet][newChartName] = rId
                editXML.Relationships.Relationship = [rel] //if match, create file with single relationship, representing new chart. rId can stay the same.
            }
        })
        const builder = new xml2js.Builder()
        const xml = builder.buildObject(editXML)
        fs.writeFileSync(`${targetDir}/xl/drawings/_rels/${newDrawingName}.xml.rels`, xml)

    })

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
                console.log('MATCH FOUND', rel, refChartName, chartToCopy)
                rId = getNewName('rId1', Object.values(target.worksheets[targetWorksheet].drawingRels))
                target.worksheets[targetWorksheet][newChartName] = rId
                sourceRelTag = rel
                sourceRelTag['$'].Id = rId
                sourceRelTag['$'].Target = `../charts/${newChartName}.xml`
            }
        })
    })
    console.log('sourceRelTag', sourceRelTag)
    const targetName = target.worksheets[targetWorksheet].drawing
    const drawingTargetPath = `${targetDir}xl/drawings/_rels/${targetName}.xml.rels`
    const drawingTargetRelsXML = fs.readFileSync(drawingTargetPath, { encoding: 'utf-8' })
    xml2js.parseString(drawingTargetRelsXML, (error, editXML) => { //insert new relations tag into drawing?.xml.rel
        editXML.Relationships.Relationship = editXML.Relationships.Relationship.concat(sourceRelTag)
        const builder = new xml2js.Builder()
        const xml = builder.buildObject(editXML)
        fs.writeFileSync(`${targetDir}/xl/drawings/_rels/${target.worksheets[targetWorksheet].drawing}.xml.rels`, xml)
    })

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
    //EVERY NEW FILE NEEDS TO ALSO BE ADDED TO CONTENT TYPES
    //add drawing tag
    if (!targetExcel.worksheets[targetWorksheet].drawing) {
        const [rId, newChartName, newDrawingName] = newDrawingRels(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet) //need to add both of these files to content types
        newDrawingXML(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet, rId, newDrawingName) //need to add both of these files to content types
        copyChartFiles(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, newChartName, stringOverrides)
        //copyChartRels
        //copyDefineNames
        //updateContentTypes
    } else {
        const [rId, newChartName, newDrawingName] = updateDrawingRels(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet)
        updateDrawingXML(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet, rId, newDrawingName)
        copyChartFiles(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, newChartName, stringOverrides)
        //copyChartRels
        //copyDefineNames
        //updateContentTypes
    }

    //chart rels
    //chart <--remember overrides
    //definedNames <--remember to update overrides
}