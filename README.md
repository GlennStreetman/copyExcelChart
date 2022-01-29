# WARNING: The XML schema can very between charts and additional testing is necessary to prove that this library works for all chart types

Consider using [excels internal JavaScript API](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-charts) BEFORE using this library.  
If Excel is not available in your working environment and you still need to copy excel charts using JavaScript, consider using this library as a code example and thoroughly test your use cases before deploying.  
XML schema can differ between each chart type and this library has not been thoroughly tested using all possible chart type and user configurations. Use with discretion.

## Copy charts between excel files using Node file system operations.

Currently working with basic excel .xlsx charts, that pull data from cell ranges. <br>
Not yet tested with pivot charts or charts that reference named ranges or tables.<br>

## Dependencies:

[xm2js](https://www.npmjs.com/package/xml2js) : Used to convert excel .xml source files into JSON objects. <br>
[AdmZip](https://www.npmjs.com/package/adm-zip) : Used to unzip .xlsx files into individual .xml files <br>

## Installation

npm i copy-excel-chart

### Working Code Example, additional explanation shown below under "Usage" heading: <br>

All Excel files used in these examples can be found in the test folder of this package. <br>

```
import copyExcelChart from 'copy-excel-chart'
import fs from 'fs'
const readCharts = copyExcelChart.readCharts
const copyChart = copyExcelChart.copyChart
const writeCharts = copyExcelChart.writeCharts

async function test(){

    if(!fs.existsSync('./working')) fs.mkdirSync('./working')

    const source = await readCharts('./source.xlsx', './working')
    const output = await readCharts('./target.xlsx', './working')

    const replaceCellRefs = source.summary()['chartWorksheet']['chart1'].reduce((acc, el)=>{
        return {...acc, [el]: el.replace('recommendWorksheet2', 'worksheet-Recommendation')}
    }, {})

    copyChart(
        source,
        output,
        'chartWorksheet',
        'chart1',
        'worksheet-Recommendation',
        replaceCellRefs,
    )

    writeCharts(output, './product.xlsx')

    fs.rmdirSync('./working', { recursive: true })
}

test()

```

## Usage:

Setup imports

```
import copyExcelChart from 'copy-excel-chart'
import fs from 'fs'
```

Setup shortcuts:

```
const readCharts = copyExcelChart.readCharts
const copyChart = copyExcelChart.copyChart
const writeCharts = copyExcelChart.writeCharts
```

Create a working folder. Make sure fs is imported!

```
if(!fs.existsSync('./working')) fs.mkdirSync('./working')
```

Read an excel file that contains source charts using the readCharts() function. <br>

```
const source = await readCharts('./source.xlsx', './working')
```

Run source.summary() to get a list of worksheets, each worksheets charts, and each charts cell references.<br>

```
Function source.summary()
```

<br>
Returns {[WorksheetName(s)]: [chart(s)]: [cell reference list]} <br>
Summary return objects are the easiest way to access worksheet names, chart names, and cell references. <br>
You can always choose to console.log(source) to get a full list of everything related to a workbook that this package uses.

```
console.log('Worksheet Summary:', source.summary())

> Worksheet Summary: {
>    recommendWorksheet2: {}, //worksheet with no charts
>    earningsWorksheet1: {}, //worksheet with no charts
>        cashWorksheet4: {}, //worksheet with no charts
>    candleWorksheet3: {}, //worksheet with no charts
>    chartWorksheet: { //worksheet with 4 charts.
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

Repeat the steps from above for the excel xlsx file that your will be copying charts into. <br>
Note that excel assigns chart names behind the scenes and sequences their names based on schema type. <br>
Chart and ChartEx and charts that use different XML schemas. Any chart that is not of type "chart" or "chartEx" probable has not been tested and this library might not be able to copy.

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
This step is necessary if the chart being copied needs updated cell references that point to a new location. <br>
Replacement Object: {[old reference]: new reference} <br>
example: {oldworksheet!A1:B20: newWorksheet!A1:B15}<br>
The Reducer function below creates an object that will be used to replace chart1's cell references with new references that point to the worksheet 'worksheet-Recommendation' instead of worksheet 'recommendWorksheet2'.

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

Copy a chart from the source working files to the output working files using the copyChart() function.<br>
Note that each time copyChart runs it edits the output file.xml(s) and updates the output object with all changes.

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

Write a new excel file: product.xlsx from the output working file using the writeChart() function <br>

```
writeCharts(output, './product.xlsx')
```

Clean up old files <br>

```
fs.rmdirSync('./working', { recursive: true })
```

# API

### copyExcelChart.readCharts() <br>

Returns an object that DETAILS worksheet chart relationships.<br>
Return type includes helper method, .summary(), that SUMMARIZES worksheet chart relationships<br>
.summary() should be significantly easier to run looping functions against than the detail.<br>

```
Function readCharts(
    source File: string,        //The path of the excel file you will be copying charts from
    working directory: string,  //a temporary working directory for file read/write operations.
)
```

### copy-excel-chart.copyChart() <br>

Updates toObject .xml files and updates updates toObject relationships. <br>
Copies a single chart. Run multiple times, with additional chart names, to copy multiple charts.

```
Function copyChart(
    fromObject: readCharts() return object,
    toObject: readCharts() return object,
    source worksheet: string,                    //source worksheet name is the worksheet alias viewable in an excel workbook
    source chart: string,                        //source chart name can be found using readCharts().summary()
    move to worksheet: string,                   //worksheet name visible in the output excel workbook.
    cell reference overrides:{[string]: string}, //object containing key value pairs that are used to update cell references. ex: {worksheet1!A1:B2: newWorksheet: C1:D2}
)
```

### copy-excel-chart.writeChart()

From the provided objects .xml files, write a new excel file.

```
Function writeChart(
    toObject: readCharts() return object,
    file name: string
)
```

## Run the tests:

```
> git clone https://github.com/GlennStreetman/copyExcelChart.git
> cd copyExcelChart
> npm install
```

Use your file explorer to open the tests subdirectory and review each sub-directories files.

```
> node runTests
```

At this point each sub-directory will have a new "product.xlsx" file as well as a new sub-folder named "working".  
product.xlsx is a copy of target.xlsx with source.xlsx's charts copied over.  
The "working" sub-directory contains that chart source, and destination XLSX files, XML source files.
