import {readCharts} from './build/readCharts.js'
import {copyChart} from './build/copyChart.js'
import {writeCharts} from './build/writeChart.js'
// import util from 'util'
import fs from 'fs';

const testrun = async ()=>{
    
    if(!fs.existsSync('./working')) fs.mkdirSync('./working')

    const source = await readCharts('./source.xlsx', './working')
    // console.log(util.inspect(source, false, null, true))
    console.log('source', source.summary())

    const output = await readCharts('./target.xlsx', './working')
    // console.log(util.inspect(output, false, null, true))
    console.log('source', output.summary())

    const replaceCellRefs = source.summary().chartWorksheet['chart1'].reduce((acc, el)=>{
        return {...acc, [el]: el.replace('recommendWorksheet2', 'worksheet-Recommendation')}
    }, {})

    copyChart(
        source,
        output,
        'chartWorksheet', //worksheet, in source file, that chart will be copied from
        'chart1', //chart that will be copied
        'worksheet-Recommendation', //worksheet, in output file, that chart will be copied to
        replaceCellRefs, //object containing key value pairs of cell references that will be replaced while chart is being copied.
    )
    
    writeCharts(output, './product.xlsx')
    fs.rmdirSync('./working', { recursive: true })
}

testrun()