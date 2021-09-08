## Copy charts between excel files using Node file system operations. 
 Currently working with basic excel .xlsx charts, that pull data from cell ranges. <br>
 Not yet tested with pivot charts or charts that reference named ranges or tables.

## Dependancies:
[xm2js](https://www.npmjs.com/package/xml2js) : Used to convert excel .xml source files into JSON objects. <br> 
[AdmZip](https://www.npmjs.com/package/adm-zip) : Used to unzip .xlsx files into individual .xml files <br> 

## Setup
Clone this repository then compile the source typescript files.
```
npm run tscd 
```

## Usage: 

```
import {readCharts} from './build/readCharts.js'
import {copyChart} from './build/copyChart.js'
import {writeCharts} from './build/writeChart.js'
import fs from 'fs';
```

Create a working folder:
```
if(!fs.existsSync('./working')) fs.mkdirSync('./working') 
```

Read an excel file that contains source charts useing the readCharts() function. <br>
```
Function readCharts(
    source File: string, 
    working directory: string
)
```
Returns an Object describing the source files charts.<br>
readCharts() also creates a working folder that contains the individual xml source files from the provided .xlsx file.<br>
```
const source = await readCharts('./source.xlsx', './working) 
```

Run source.summary() to get a list of worksheets, each worksheets charts, and each charts cell references. <br>
 source.summary() <br>
 returns {[WorksheetName(s)]: [chart(s)]: [cell reference list]} <br>

```
console.log('Worksheet Summary:', source.summary()) 

> Worksheet Summary: {
>    recommendWorksheet2: {}, //worksheet with no charts
>    earningsWorksheet1: {}, //worksheet with no charts
>        cashWorksheet4: {}, //worksheet with no charts
>    candleWorksheet3: {}, //worksheet with no charts
>    chartWorksheet: { //worksheetwith 4 charts.
>        chart3: [ //chart3 cell reference array
>            'cashWorksheet4!$B$2:$B$22', 
>            'cashWorksheet4!$C$2:$C$22',
>            'cashWorksheet4!$C$1'
>        ],
>        chart2: [ //chart2 cell reference array
>            'candleWorksheet3!$B$2:$B$26', 
>            'candleWorksheet3!$C$2:$C$26',
>            'candleWorksheet3!$B$2:$B$27',
>            'candleWorksheet3!$D$2:$D$26',
>            'candleWorksheet3!$E$2:$E$26',
>            'candleWorksheet3!$F$2:$F$26',
>            'candleWorksheet3!$C$1',
>            'candleWorksheet3!$D$1',
>            'candleWorksheet3!$E$1',
>            'candleWorksheet3!$F$1'
>        ],
>        chart1: [ //chart1 cell reference array
>            'recommendWorksheet2!$B$2:$B$42', 
>            'recommendWorksheet2!$C$2:$C$42',
>            'recommendWorksheet2!$D$2:$D$42',
>            'recommendWorksheet2!$E$2:$E$42',
>            'recommendWorksheet2!$F$2:$F$42',
>            'recommendWorksheet2!$G$2:$G$42',
>            'recommendWorksheet2!$C$1',
>            'recommendWorksheet2!$D$1',
>            'recommendWorksheet2!$E$1',
>            'recommendWorksheet2!$F$1',
>            'recommendWorksheet2!$G$1'
>        ],
>        chartEx1: [ //chartEx1 cell reference array
>            'earningsWorksheet1!$B$2:$B$22', 
>            'earningsWorksheet1!$C$1',
>            'earningsWorksheet1!$C$2:$C$22'
>        ]
>    }
> }
```

Repeat the steps of above for the excel xlsx file that your will be copying charts into. <br>
```
const output = await readCharts('./target.xlsx', './working') 
console.log('Worksheet Summary:', output.summary())

> Worksheet Summary: { 
>  'worksheet-candle': {}, //worksheet with no charts
>  'worksheet-Recommendation': {}, //worksheet with no charts
>  'worksheet-EBIT': {}, //worksheet with no charts
>  'worksheet-cashRatio': {} //worksheet with no charts
> }
```

Create a cell reference replacement object. <br>
This step is necesarry if the new chart needs cell references that point to a new location. <br>
Replacement Object: {[old reference]: new reference} <br>
example: {oldworksheet!A1:B20: newWorksheet!A1:B15}<br>
```
const replaceCellRefs = source.summary()['chartWorksheet']['chart1'].reduce((acc, el)=>{
    return {...acc, [el]: el.replace('recommendWorksheet2', 'worksheet-Recommendation')}
}, {})
console.log('Cell Reference overrides:', replaceCellRefs)
> Cell Reference overrides:
> {
>  'recommendWorksheet2!$B$2:$B$42': 'worksheet-Recommendation!$B$2:$B$42',
>  'recommendWorksheet2!$C$2:$C$42': 'worksheet-Recommendation!$C$2:$C$42',
>  'recommendWorksheet2!$D$2:$D$42': 'worksheet-Recommendation!$D$2:$D$42',
>  'recommendWorksheet2!$E$2:$E$42': 'worksheet-Recommendation!$E$2:$E$42',
>  'recommendWorksheet2!$F$2:$F$42': 'worksheet-Recommendation!$F$2:$F$42',
>  'recommendWorksheet2!$G$2:$G$42': 'worksheet-Recommendation!$G$2:$G$42',
>  'recommendWorksheet2!$C$1': 'worksheet-Recommendation!$C$1',
>  'recommendWorksheet2!$D$1': 'worksheet-Recommendation!$D$1',
>  'recommendWorksheet2!$E$1': 'worksheet-Recommendation!$E$1',
>  'recommendWorksheet2!$F$1': 'worksheet-Recommendation!$F$1',
>  'recommendWorksheet2!$G$1': 'worksheet-Recommendation!$G$1'
> }

```
Copy a chart from source working file to output working file using the copyChart() function.<br>
```
Function copyChart( 
    from Object: readCharts() return object, 
    to Object: readCharts() return object,  
    source worksheet: string,
    source chart: string,  
    move to worksheet: string, 
    cell reference overrides: {[key: string]: string} 
)
```
copyChart edits the to Objects working file .xmls
```
copyChart(
    source, 
    output, 
    'chartWorksheet', 
    'chart1', 
    'worksheet-Recommendation', 
    replaceCellRefs, 
)
```

If additional charts need to be copied do so here by performing addtional copyChart() operations. <br>
```
Function writeChart(
    to Object: readCharts() return object,
    file name: string
)
```
Write a new excel file: product.xlsx from the output working file using the writeChart() function <br>
```
writeCharts(output, './product.xlsx') 
```

Clean up old files <br>
```
fs.rmdirSync('./working', { recursive: true })
```