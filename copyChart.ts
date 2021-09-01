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

function updateDrawingXML(
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
    if (!target.worksheets[targetWorksheet].drawing) {
        console.log('no drawing.xml')
        //if drawing.xml does not exist for target worksheet, copy source drawing.xml and set Relationships.relation =  source.drawingXML
        //make sure to update drawingXML rId = new rID passed into function. File name should match new drawing name.
        b
        //this cannot be a equal copy. Only one of the source drawing xml subsections needs to be copied over if new file.

        // const chartName = getNewName('drawing1', target)
        // if (!fs.existsSync(`${targetDir}xl/drawings/`)) fs.mkdirSync(`${targetDir}xl/drawings/`, { recursive: true })
        // fs.copyFileSync(`${sourceDir}xl/drawings/${source[sourceWorksheet].drawing}.xml`, `${targetDir}xl/drawings/${chartName}.xml`)
    } else { //if drawing.xml exists for target worksheet combine <xdr:twoCellAnchor> tags from source and target drawing file. New cellAnchor needs to have its rID updated.

        const sourceDrawingRef: string = source.worksheets[sourceWorksheet].drawing
        const drawingSource = source.drawingXMLs[sourceDrawingRef][chartToCopy] //xml2Js object representing source drawing.xml sub section.
        const drawingXML = fs.readFileSync(`${targetDir}xl/drawings/${target.worksheets[targetWorksheet].drawing}.xml`, { encoding: 'utf-8' })
        xml2js.parseString(drawingXML, (error, editXML) => {
            //replace source drawing ref with new ref. Remember to update drawing ref in target. 
            editXML['xdr:wsDr']['xdr:twoCellAnchor'] = editXML['xdr:wsDr']['xdr:twoCellAnchor'].concat(drawingSource)
            const builder = new xml2js.Builder()
            const xml = builder.buildObject(editXML)
            fs.writeFileSync(`${targetDir}/xl/drawings/${target.worksheets[targetWorksheet].drawing}.xml`, xml)
        })

    }
    //ADD DRAWING TO target.worksheets[targetWorksheet].drawing
}

function updateDrawingRels(
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
    if (!target.worksheets[targetWorksheet].drawing) { //if drawing.xml does not exist for target worksheet
        console.log('Parsing New drawingRels ')
        const newDrawingName = getNewName('drawing1', target.drawingList) //used for naming drawing.xml and drawing.xml.rels
        if (!fs.existsSync(`${targetDir}xl/drawings/_rels/`)) fs.mkdirSync(`${targetDir}xl/drawings/_rels/`, { recursive: true }) //make drawing directory if it doesnt exist yet.

        xml2js.parseString(drawingSourceRelsXML, (error, editXML) => {
            editXML.Relationship.Relationship.forEach((rel) => {
                const refChartName = rel.Target.replace('../charts/', '').replace('.xml.rels', '')
                if (refChartName === chartToCopy) {
                    rel.Target = `../charts/${newChartName}.xml.rels`
                    editXML.Relationship.Relationship = rel //if match, create file with single relationship, representing new chart. rId can stay the same.
                }
            })
            const builder = new xml2js.Builder()
            const xml = builder.buildObject(editXML)
            fs.writeFileSync(`${targetDir}/xl/drawings/_rels/${newDrawingName}.xml.rels`, xml)

        })

        return [rId, newChartName, newDrawingName]

    } else { //if drawing.xml exists for target worksheet combine <xdr:twoCellAnchor> tags from source and target drawing file. Update rId and ChartName
        console.log('Parsing Old drawingRels ')

        let sourceRelTag
        xml2js.parseString(drawingSourceRelsXML, (error, editXML) => { //make a copy of the source relationship tag after updating rId & target.
            editXML.Relationships.Relationship.forEach((rel) => {
                const refChartName = rel['$'].Target.replace('../charts/', '').replace('.xml', '')
                // console.log('------------REL-------------', rel, refChartName, chartToCopy)
                if (refChartName === chartToCopy) {
                    console.log('MATCH FOUND', rel, refChartName, chartToCopy)
                    rId = getNewName('rId1', Object.values(target.worksheets[targetWorksheet].drawingRels))
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
            console.log('editXML', editXML.Relationships.Relationship.concat(sourceRelTag))
            editXML.Relationships.Relationship = editXML.Relationships.Relationship.concat(sourceRelTag)
            const builder = new xml2js.Builder()
            const xml = builder.buildObject(editXML)
            fs.writeFileSync(`${targetDir}/xl/drawings/_rels/${target.worksheets[targetWorksheet].drawing}.xml.rels`, xml)
        })

        return [rId, newChartName, '']
    }
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
    //add drawing rel first, need to get rID and add tag for new drawing.
    const [rId, newChartName, newDrawingName] = updateDrawingRels(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet)
    // updateDrawingXML(sourceExcel, targetExcel, sourceWorksheet, chartToCopy, targetWorksheet, rId, newDrawingName)
    //chart rels
    //chart <--remember overrides
    //definedNames <--remember to update overrides
}