Copy charts between excel files using Node file system operations. Currently working with basic charts that pull data from cell ranges. Not yet tested with pivot charts or charts that reference named ranges or tables.

Dependancies:
xm2js
AdmZip

Usage: 

import {readCharts} from './build/readCharts.js'
import {copyChart} from './build/copyChart.js'
import {writeCharts} from './build/writeChart.js'

//Create a working folder.
if(!fs.existsSync('./output')) fs.mkdirSync('./output') 

//read excel file that contains source chart
const source = await readCharts('./source.xlsx', './output) //params: excel file location, working directory.

//example source workbook contains multiple worksheets, the worksheet "chartWorksheet" contains four charts.
//print summary object. {[WorksheetName(s)]: [chart(s)]: [cell reference list]}
console.log(source.summary())

RETURNS:
{
    recommendWorksheet2: {},
    earningsWorksheet1: {},
        cashWorksheet4: {},
    candleWorksheet3: {}
    chartWorksheet: {
        chart3: [
            'cashWorksheet4!$B$2:$B$22',
            'cashWorksheet4!$C$2:$C$22',
            'cashWorksheet4!$C$1'
        ],
        chart2: [
            'candleWorksheet3!$B$2:$B$26',
            'candleWorksheet3!$C$2:$C$26',
            'candleWorksheet3!$B$2:$B$27',
            'candleWorksheet3!$D$2:$D$26',
            'candleWorksheet3!$E$2:$E$26',
            'candleWorksheet3!$F$2:$F$26',
            'candleWorksheet3!$C$1',
            'candleWorksheet3!$D$1',
            'candleWorksheet3!$E$1',
            'candleWorksheet3!$F$1'
        ],
        chart1: [
            'recommendWorksheet2!$B$2:$B$42',
            'recommendWorksheet2!$C$2:$C$42',
            'recommendWorksheet2!$D$2:$D$42',
            'recommendWorksheet2!$E$2:$E$42',
            'recommendWorksheet2!$F$2:$F$42',
            'recommendWorksheet2!$G$2:$G$42',
            'recommendWorksheet2!$C$1',
            'recommendWorksheet2!$D$1',
            'recommendWorksheet2!$E$1',
            'recommendWorksheet2!$F$1',
            'recommendWorksheet2!$G$1'
        ],
        chartEx1: [
            'earningsWorksheet1!$B$2:$B$22',
            'earningsWorksheet1!$C$1',
            'earningsWorksheet1!$C$2:$C$22'
        ]
    },
}


<!-- read excel file that charts are going to be copied into -->
const output = await readCharts('./target.xlsx', './output') 
console.log(output.summary())

RETURNS:
{                                   //a workbook containing 4 worksheets and no charts.
  'worksheet-candle': {},
  'worksheet-Recommendation': {},
  'worksheet-EBIT': {},
  'worksheet-cashRatio': {}
}


//create a cell reference replacement object so that that charts cell references dont all break after copying the chart into a new workbook.
const replaceCellRefs = source.summary().sourceWorksheet['chart1'].reduce((acc, el)=>{
    return {...acc, [el]: el.replace('recommendWorksheet2', 'worksheet-Recommendation')}
}, {})

//copy chart from source to output files.
copyChart(
    source,
    output,
    'chartWorksheet', //worksheet, in source file, that chart will be copied from
    'chart1', //chart that will be copied
    'worksheet-Recommendation', //worksheet, in output file, that chart will be copied to
    replaceCellRefs, //object containing key value pairs of cell references that will be replaced while chart is being copied.
)

//from output, write new file: product.xlsx
writeCharts(output, './product.xlsx')
