import {readCharts} from './build/readCharts.js'
import {copyChart} from './build/copyChart.js'
import {writeCharts} from './build/writeChart.js'
import {cleanup} from './build/cleanup.js'
import util from 'util'
import fs from 'fs';

const testrun = async ()=>{
    
    if(!fs.existsSync('./output')) fs.mkdirSync('./output')

    fs.copyFileSync(`./testSource1.xlsx`, `./test.xlsx`)
    // fs.copyFileSync(`./testSource2.xlsx`, `./test.xlsx`)

    const source = await readCharts('./test.xlsx', './output')
    // console.log(util.inspect(source, false, null, true))
    console.log('source', source.summary())
    const output = await readCharts('./testOutput.xlsx', './output')
    // console.log('output', util.inspect(output, false, null, true))
    // console.log('output', output)
    console.log('----------Starting on chart1: -----------')
    copyChart(
        source,
        output,
        'sourceWorksheet',
        'chartEx1',
        'EBIT',
        {
            [`'EBIT-US-TSLA'!$B$2:$B$22`]: `'EBIT'!$B$2:$B$16`,
            [`'EBIT-US-TSLA'!$C$1`]: `'EBIT'!$C$1`,
            [`'EBIT-US-TSLA'!$C$2:$C$22`]: `'EBIT'!$C$2:$C$16`,
        },
    )
    console.log('----------Starting on chart2: -----------')
    let newOverrides = source.summary().sourceWorksheet['chart2'].reduce((acc, el)=>{
        return {...acc, [el]: el.replace('candle-US-TSLA', 'candle')}
    }, {})
    console.log('new overrides', newOverrides)
    copyChart(
        source,
        output,
        'sourceWorksheet',
        'chart2',
        'candle',
        newOverrides,
    )
    writeCharts(output, './product.xlsx')
    // fs.rmdirSync('./output', { recursive: true })
}

testrun()