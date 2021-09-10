import {readCharts} from './../../build/readCharts.js'
import {copyChart} from './../../build/copyChart.js'
import {writeCharts} from './../../build/writeChart.js'
import util from 'util'
import fs from 'fs';

const testrun = async ()=>{ 
    //copy charts from one .xlsx worksheet to many .xlsx worksheets.
    
    if(!fs.existsSync('./tests/copyFromOneToManySheets/working')) fs.mkdirSync('./tests/copyFromOneToManySheets/working')

    const source = await readCharts('./tests/copyFromOneToManySheets/source.xlsx', './tests/copyFromOneToManySheets/working')
    // console.log(util.inspect(source, false, null, true))
    // console.log('source', source.summary())

    const output = await readCharts('./tests/copyFromOneToManySheets/target.xlsx', './tests/copyFromOneToManySheets/working')
    // console.log(util.inspect(output, false, null, true))
    // console.log('output', output.summary())

    const replaceCellRefs1 = source.summary().chartWorksheet['chart1'].reduce((acc, el)=>{
        return {...acc, [el]: el.replace('recommendWorksheet2', 'worksheet-Recommendation')}
    }, {})

    await copyChart(
        source,
        output,
        'chartWorksheet', //worksheet, in source file, that chart will be copied from
        'chart1', //chart that will be copied
        'worksheet-Recommendation', //worksheet, in output file, that chart will be copied to
        replaceCellRefs1, //object containing key value pairs of cell references that will be replaced while chart is being copied.
    )

    const replaceCellRefs2 = source.summary().chartWorksheet['chartEx1'].reduce((acc, el)=>{
        return {...acc, [el]: el.replace('earningsWorksheet1', "worksheet-EBIT")}
    }, {})

    await copyChart(
        source,
        output,
        'chartWorksheet', //worksheet, in source file, that chart will be copied from
        'chartEx1', //chart that will be copied
        'worksheet-EBIT', //worksheet, in output file, that chart will be copied to
        replaceCellRefs2, //object containing key value pairs of cell references that will be replaced while chart is being copied.
    )

    const replaceCellRefs3 = source.summary().chartWorksheet['chart2'].reduce((acc, el)=>{
        return {...acc, [el]: el.replace('candleWorksheet3', "worksheet-candle")}
    }, {})

    await copyChart(
        source,
        output,
        'chartWorksheet', //worksheet, in source file, that chart will be copied from
        'chart2', //chart that will be copied
        'worksheet-candle', //worksheet, in output file, that chart will be copied to
        replaceCellRefs3, //object containing key value pairs of cell references that will be replaced while chart is being copied.
    )

    const replaceCellRefs4 = source.summary().chartWorksheet['chart3'].reduce((acc, el)=>{
        return {...acc, [el]: el.replace('cashWorksheet4', "worksheet-cashRatio")}
    }, {})

    await copyChart(
        source,
        output,
        'chartWorksheet', //worksheet, in source file, that chart will be copied from
        'chart3', //chart that will be copied
        'worksheet-cashRatio', //worksheet, in output file, that chart will be copied to
        replaceCellRefs4, //object containing key value pairs of cell references that will be replaced while chart is being copied.
    )

    // console.log(util.inspect(output, false, null, true))
    // console.log('output', output.summary())

    await writeCharts(output, './tests/copyFromOneToManySheets/product.xlsx')
    // fs.rmdirSync('./tests/copyFromOneToManySheets/working', { recursive: true })
}

testrun()