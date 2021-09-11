import fs from 'fs';
import xml2js from 'xml2js';
import AdmZip from 'adm-zip';


export interface workbookChartDetails {
    tempDir: string, //file path source.
    worksheets: worksheets,
    drawingList: string[],
    drawingXMLs: drawingId,
    chartList: string[],
    defineNameRefs: string[],
    summary: Function,
    colorList: string[],
    styleList: string[],
}

export interface worksheets {
    [key: string]: worksheetObj //key strings worksheet aliaes (do not include .xml)
}

export interface drawingObj {
    name: string,
    rId: string,
    xmlJS: any,
}

export interface worksheetObj {
    name: string, //chart file name
    drawing: string, //drawing file name
    drawingRels: drawingRels,
    charts: charts, //list object for all charts in worksheet
}

interface charts {
    [key: string]: chartObj //key string is chart names from excel xml file, not visable to user.
}

interface chartObj {
    chartRels: chartRelsObj,
    cellRefs: string[],
    definedNameRefs: string[],
}

interface chartRelsObj {
    colors: string,
    style: string,
}

interface sheetNames {
    [key: string]: string //{[alias]: worksheet.xml name}
}

interface defineNames {
    [key: string]: string //{[alias]: worksheet.xml name}
}

interface drawingRels {
    [key: string]: string //{[alias]: worksheet.xml name}
}

interface chartList {
    [key: string]: string[]
}

interface cellRefs {
    [key: string]: string[]
}

interface definedNameRefs {
    [key: string]: string[]
}

interface chartRelsObj {
    colors: string,
    style: string
}

interface chartRelList {
    [key: string]: chartRelsObj
}

interface drawObj {
    [key: string]: any //rId: xmlBlob
}

interface drawingId { //key is drawing name
    [key: string]: drawObj //worksheet: drawObj
}


function buildChartList(chartList: chartList): string[] {
    let allCharts: string[] = []
    Object.values(chartList).forEach((el) => {
        allCharts = allCharts.concat(el)
    })
    return allCharts
}

function buildDrawingList(drawingList: drawingRels): string[] {
    let allDrawings: string[] = []
    Object.values(drawingList).forEach((el) => {
        allDrawings.push(el)
    })
    return allDrawings
}

function buildFileList(chartRels: chartRelList, ref: string) {
    const returnList: string[] = []
    Object.values(chartRels).forEach((chart) => {
        returnList.push(chart[ref])

    })
    return returnList
}

function buildChartDetails(
    tempFolder: string,
    worksheetNames: sheetNames,
    drawingList: drawingRels,
    drawingXMLs: drawingId,
    chartList: chartList,
    cellRefs: cellRefs,
    definedNameRefs: definedNameRefs,
    chartRels: chartRelList,
    drawingrIds: drawingId,
    definedNameKeys: any,
): workbookChartDetails {
    const workbook: workbookChartDetails = {
        tempDir: tempFolder,
        drawingList: buildDrawingList(drawingList),
        drawingXMLs: drawingXMLs,
        chartList: buildChartList(chartList),
        defineNameRefs: definedNameKeys ? definedNameKeys : [], //buildChartList(definedNameRefs)
        colorList: buildFileList(chartRels, 'colors'),
        styleList: buildFileList(chartRels, 'style'),
        worksheets: function () {
            const worksheetList = Object.entries(worksheetNames).reduce((acc, [key, val]) => {
                return {
                    ...acc, [key]: {
                        name: val,
                        drawing: '',
                        drawingRels: {},
                        charts: {},
                    }
                }
            }, {})
            Object.keys(worksheetList).forEach((worksheet) => {

                if (drawingList[worksheet]) worksheetList[worksheet]['drawing'] = drawingList[worksheet]
                if (worksheetList[worksheet].drawing) {
                    const drawingName = worksheetList[worksheet].drawing
                    worksheetList[worksheet].drawingRels = Object.entries(drawingrIds[drawingName]).reduce((acc, [key, val]) => { return { ...acc, [val]: key } }, {}) //returns drawingName: rId
                    // worksheetList[worksheet].charts = {}
                    chartList[worksheetList[worksheet].drawing].forEach((chart) => {
                        worksheetList[worksheet].charts[chart] = {
                            chartRels: chartRels[chart],
                            cellRefs: cellRefs[chart],
                            definedNameRefs: definedNameRefs[chart],
                        }
                    })
                }
            })
            return worksheetList
        }(),
        summary: function () {
            const summaryObj = {}
            Object.keys(this.worksheets).forEach((worksheet) => {
                if (this.worksheets[worksheet].charts) {
                    summaryObj[worksheet] = Object.entries(this.worksheets[worksheet].charts).reduce((acc, [key, val]) => {
                        const rtnObj = { [key]: [...new Set(val.cellRefs.concat(val.definedNameRefs))] }
                        return { ...acc, ...rtnObj }
                    }, {})
                } else {
                    summaryObj[worksheet] = {}
                }
            })
            return summaryObj
        }
    }
    return workbook
}

function findChartRels(chartList: chartList, tempFolder: string): chartRelList {

    const chartRels: chartRelList = {}
    Object.values(chartList).forEach((list) => {
        list.forEach((chart) => {
            chartRels[chart] = { colors: '', style: '' }
            const chartRelsXML = fs.readFileSync(`${tempFolder}/xl/charts/_rels/${chart}.xml.rels`, { encoding: 'utf-8' })
            xml2js.parseString(chartRelsXML, async (error, res) => {
                res.Relationships.Relationship.forEach((rel) => {
                    if (rel['$']?.Target && rel['$'].Target.includes('color')) chartRels[chart]['colors'] = rel['$'].Target.replace('.xml', '')
                    if (rel['$']?.Target && rel['$'].Target.includes('style')) chartRels[chart]['style'] = rel['$'].Target.replace('.xml', '')
                })
            })
        })
    })
    return chartRels
}

function findChartCellRefs(chartList: chartList, aliaslist, tempFolder: string, definedNames: { [key: string]: string }): [cellRefs, definedNameRefs, string[]] {
    const cellRefs: cellRefs = {}
    const definedNameRefs: definedNameRefs = {}
    const tempRefs = {}
    Object.values(chartList).forEach((chartList) => {
        chartList.forEach((chart) => {
            cellRefs[chart] = []
            const chartXML = fs.readFileSync(`${tempFolder}/xl/charts/${chart}.xml`, { encoding: 'utf-8' })
            aliaslist.forEach((alias) => { //worksheet alias refs in cell formulas are inconsistent. Somtimes they have commas sometimes they dont: Ex. 'worksheet1'!A1:B2 OR worksheet1!A1:B2

                //find matching cell ranges
                //check for matching without commas around names 
                let worksheetCellRefRange = `${alias}!\\$[A-Z]{1,3}\\$[0-9]{1,7}:\\$[A-Z]{1,3}\\$[0-9]{1,7}`
                let findCellRef = new RegExp(worksheetCellRefRange, 'g')
                let matchListNoCommas = [...new Set(chartXML.match(findCellRef))]
                //check for matching WITH commas around names 
                worksheetCellRefRange = `'${alias}'!\\$[A-Z]{1,3}\\$[0-9]{1,7}:\\$[A-Z]{1,3}\\$[0-9]{1,7}`
                findCellRef = new RegExp(worksheetCellRefRange, 'g')
                let matchListCommas = [...new Set(chartXML.match(findCellRef))]

                let matchList = [...new Set(matchListNoCommas.concat(matchListCommas))]

                //find matching cell refs.
                let worksheetCellRef = `${alias}!\\$[A-Z]{1,3}\\$[0-9]{1,7}<`
                let findCellRefCell = new RegExp(worksheetCellRef, 'g')
                let matchListCellNoCommas = [...new Set(chartXML.match(findCellRefCell))]
                matchListCellNoCommas = matchListCellNoCommas.map(el => el.slice(0, el.length - 1))
                //check for matching WITH commas around names 
                worksheetCellRef = `'${alias}'!\\$[A-Z]{1,3}\\$[0-9]{1,7}<`
                findCellRefCell = new RegExp(worksheetCellRef, 'g')
                let matchListCellCommas = [...new Set(chartXML.match(findCellRefCell))]
                matchListCellCommas = matchListCellCommas.map(el => el.slice(0, el.length - 1))

                matchList = [...new Set(matchList.concat(matchListCellNoCommas).concat(matchListCellCommas))]

                cellRefs[chart] = [...new Set(cellRefs[chart].concat(matchList))]
            })
            //some chart types use named ranges that are stored in worbook.xml. The refs to workbook.xml look like _xlchart.v?.?
            definedNameRefs[chart] = []
            let refRegex = new RegExp(`>_xlchart.v[0-9]{1,9}.[0-9]{1,10}<`, 'g')
            let matchingDefinedNameRefs = [...new Set(chartXML.match(refRegex))].forEach((el) => {
                const foundRef = el.slice(1, el.length - 1)
                tempRefs[foundRef] = chart
            })
        })
    })

    Object.entries(definedNames).forEach(([key, val]) => {
        if (tempRefs[key]) definedNameRefs[tempRefs[key]].push(val)
    })

    const definedNameKeys: string[] = Object.keys(tempRefs)

    return [cellRefs, definedNameRefs, definedNameKeys]
}

function findDrawingXML(drawingObj: drawingId, sourceFolder: string): drawingId {
    const returnObj = {}
    Object.keys(drawingObj).forEach((drawingName) => {
        returnObj[drawingName] = {}
        const drawingXML = fs.readFileSync(`${sourceFolder}/xl/drawings/${drawingName}.xml`, { encoding: 'utf-8' })
        xml2js.parseString(drawingXML, async (error, res) => {
            const targetAnchors = res['xdr:wsDr']['xdr:twoCellAnchor']
            targetAnchors.forEach((el) => {
                const rIdRegular = el?.['xdr:graphicFrame']?.[0]?.['a:graphic']?.[0]?.['a:graphicData']?.[0]?.['c:chart']?.[0]?.['$']?.['r:id'] //regular chart?.xml
                if (rIdRegular) {
                    const chartName = drawingObj[drawingName][rIdRegular]
                    returnObj[drawingName][chartName] = el
                }
                const rIdRegularEx = el?.['mc:AlternateContent']?.[0]?.['mc:Choice']?.[0]?.['xdr:graphicFrame']?.[0]?.['a:graphic']?.[0]?.['a:graphicData']?.[0]?.['cx:chart'][0]['$']['r:id'] //alternate chart type chartEx?.xml
                if (rIdRegularEx) {
                    const chartName = drawingObj[drawingName][rIdRegularEx]
                    returnObj[drawingName][chartName] = el
                }

            })
        })
    })
    return returnObj
}

function parseDrawingRels(drawingList: drawingRels, tempFolder: string): [chartList, drawingId] { //reads drawing.xml.rels to find associated list of charts and xml blob in drawing?.xml

    const chartList = {}
    const drawingrIds: drawingId = {}
    Object.entries(drawingList).forEach(([alias, name]) => {
        const drawingRelsXML = fs.readFileSync(`${tempFolder}/xl/drawings/_rels/${name}.xml.rels`, { encoding: 'utf-8' })
        xml2js.parseString(drawingRelsXML, async (error, res) => {
            drawingrIds[name] = {}
            const chartRels = res.Relationships.Relationship.reduce((acc, el) => {
                if (el['$'].Target.includes('charts/')) {
                    const chartName = el['$'].Target.replace('../charts/', '').replace('.xml', '')
                    acc.push(chartName);
                    drawingrIds[name][el['$'].Id] = chartName
                    return acc
                } else {
                    return acc
                }
            }, [])
            chartList[name] = chartRels
        })
    })
    return [chartList, drawingrIds]
}

function findDrawingRels(sheetNames: sheetNames, tempFolder: string): drawingRels {
    let drawingRelsAcc = {}

    Object.entries(sheetNames).forEach(([alias, name]) => { //read worksheets relationss, check for link to drawing
        if (fs.existsSync(`${tempFolder}/xl/worksheets/_rels/${name}.xml.rels`)) {
            const worksheetRelsXML = fs.readFileSync(`${tempFolder}/xl/worksheets/_rels/${name}.xml.rels`, { encoding: 'utf-8' })
            xml2js.parseString(worksheetRelsXML, async (error, res) => {
                const drawingRels = res.Relationships.Relationship.reduce((acc, el) => {
                    if (el['$'].Target.includes('drawings/')) { return { ...acc, [alias]: el['$'].Target.replace('../drawings/', '').replace('.xml', '') } } else { return acc }
                }, {})
                drawingRelsAcc = { ...drawingRelsAcc, ...drawingRels }
            })
        }
    })
    return drawingRelsAcc
}

function readSheetNames(tempFolder: string): [sheetNames, defineNames] {

    //read workbook.xml and workbook.xml.rels to find list of worksheet file names and aliases.
    //aliases are worksheet names visable to excel users.
    let sheetIds = {}
    let definedName = {} //named cell references, refered to some type of charts, often named chartEX?.xml
    const workbookXML = fs.readFileSync(`${tempFolder}xl/workbook.xml`, { encoding: 'utf-8' })
    xml2js.parseString(workbookXML, (error, res) => { //read workbook.xml sheet list. Source of worksheet aliases.
        const sheetRelAlias = res?.workbook?.sheets ? res.workbook.sheets[0].sheet.reduce((acc, el) => { //aliases are names that are visable to user
            return { ...acc, [el['$']['r:id']]: el['$'].name }
        }, {}) : {}
        sheetIds = sheetRelAlias
        definedName = res?.workbook?.definedNames ? res.workbook.definedNames[0].definedName.reduce((acc, el) => { //aliases are names that are visable to user
            return { ...acc, [el['$']['name']]: el['_'] }
        }, {}) : {}
    })

    const workbookXML_rels = fs.readFileSync(`${tempFolder}xl/_rels/workbook.xml.rels`, { encoding: 'utf-8' })
    xml2js.parseString(workbookXML_rels, (error, res) => { //read workbook.xml.rels sheet list. Source of worksheet.xml file names
        const sheetRelName = res.Relationships.Relationship.reduce((acc, el) => {
            if (sheetIds[el['$'].Id]) {
                return { ...acc, [sheetIds[el['$'].Id]]: el['$'].Target.replace('worksheets/', '').replace('.xml', '') }
            } else { return { ...acc } }
        }, {})
        sheetIds = sheetRelName
    })
    return [sheetIds, definedName]
}

export function readCharts(
    sourceFile: string, //location of file to read
    tempFolder: string,  //location to store unzipped file.
) {
    const filePath = sourceFile.replace(/\\/, 'g')
    const fileName = filePath.slice(filePath.lastIndexOf('/') + 1, filePath.length).replace('.xlsx', '/')
    const sourceFolder = `${tempFolder}/${fileName}`
    if (fs.existsSync(sourceFolder)) fs.rmdirSync(sourceFolder, { recursive: true }) //remove old files that have been parced at the same location.
    fs.mkdirSync(sourceFolder)
    return new Promise((resolve, reject) => {
        try { //under the hood, excel files are zip folders containing xml files.
            const zip = new AdmZip(sourceFile)
            zip.extractAllTo(sourceFolder, true) //unzip excel template file to dump folder so that we can access xml files.
            const [worksheetNames, definedNames] = readSheetNames(sourceFolder) //returns  {[alias]: worksheet.xml name}
            const drawingList = findDrawingRels(worksheetNames, sourceFolder) //returns {[alias]: drawings}
            const [chartList, drawingrIds] = Object.keys(drawingList).length > 0 ? parseDrawingRels(drawingList, sourceFolder) : [{}, {}] //find associated charts and xml blob associated with drawing in drawing.xml
            const findDrawingXMLs = Object.keys(drawingrIds).length > 0 ? findDrawingXML(drawingrIds, sourceFolder) : drawingrIds //drawing xml file.
            const [cellRefs, definedNameRefs, definedNameKeys] = Object.keys(chartList).length > 0 ? findChartCellRefs(chartList, Object.keys(worksheetNames), sourceFolder, definedNames) : [{}, {}] //excel formula ranges and and cell refs.
            const chartRefs = Object.keys(chartList).length > 0 ? findChartRels(chartList, sourceFolder) : {} //related chart xmls. Style & colors.
            const chartDetails = buildChartDetails(sourceFolder, worksheetNames, drawingList, findDrawingXMLs, chartList, cellRefs, definedNameRefs, chartRefs, drawingrIds, definedNameKeys)
            resolve(chartDetails)
        } catch (error) {
            console.log('Read file error. Path: ', sourceFile, 'Error: ', error)
            reject(error)
        }
    })
}


