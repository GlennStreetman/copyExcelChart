import {readCharts} from './../../build/readCharts.js'
import {copyChart} from './../../build/copyChart.js'
import {writeCharts} from './../../build/writeChart.js'
import util from 'util'
import fs from 'fs';

export const copyNoOverrides = async ()=>{ 
    //copy charts from one .xlsx worksheet to many .xlsx worksheets.
    console.log('starting test')
    if(!fs.existsSync('./tests/noOverridesOneToOne/working')) fs.mkdirSync('./tests/noOverridesOneToOne/working')

    const source = await readCharts('./tests/noOverridesOneToOne/source.xlsx', './tests/noOverridesOneToOne/working/')
    // console.log(util.inspect(source, false, null, true))
    // console.log('source', source.summary())

    const output = await readCharts('./tests/noOverridesOneToOne/target.xlsx', './tests/noOverridesOneToOne/working')
    // console.log(util.inspect(output, false, null, true))
    // console.log('output', output.summary())

    await copyChart(
        source,
        output,
        'chartWorksheet', //worksheet, in source file, that chart will be copied from
        'chart1', //chart that will be copied
        'chartWorksheet', //worksheet, in output file, that chart will be copied to
        {}, //object containing key value pairs of cell references that will be replaced while chart is being copied.
    )

    await copyChart(
        source,
        output,
        'chartWorksheet', //worksheet, in source file, that chart will be copied from
        'chartEx1', //chart that will be copied
        'chartWorksheet', //worksheet, in output file, that chart will be copied to
        {}, //object containing key value pairs of cell references that will be replaced while chart is being copied.
    )

    await copyChart(
        source,
        output,
        'chartWorksheet', //worksheet, in source file, that chart will be copied from
        'chart2', //chart that will be copied
        'chartWorksheet', //worksheet, in output file, that chart will be copied to
        {}, //object containing key value pairs of cell references that will be replaced while chart is being copied.
    )

    await copyChart(
        source,
        output,
        'chartWorksheet', //worksheet, in source file, that chart will be copied from
        'chart3', //chart that will be copied
        'chartWorksheet', //worksheet, in output file, that chart will be copied to
        {}, //object containing key value pairs of cell references that will be replaced while chart is being copied.
    )

    // console.log(util.inspect(output, false, null, true))
    // console.log('output', output.summary())

    await writeCharts(output, './tests/noOverridesOneToOne/product.xlsx')
    // fs.rmdirSync('./tests/copyFromOneToManySheets/working', { recursive: true })
}
