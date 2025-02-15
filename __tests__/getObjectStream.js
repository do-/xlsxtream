const xlsx = require ('..')

test ('2_____with_chart', async () => {

	const wb = await xlsx.open ('__data__/2_____with_chart.xlsx')	
	const ws = wb.sheetByName.Tabelle1
	
	const ns = await ws.getObjectStream ()
	
	const a = []; for await (n of ns) a.push (n)
	expect (a).toStrictEqual ( [['1', '10'], ['2', '20']])
 	
})

test ('file_example_XLSX_10', async () => {

	const wb = await xlsx.open ('__data__/file_example_XLSX_10.xlsx')
	expect (wb.sharedStrings.at (-1)).toBe ('')

	const ws = wb.sheets [0]
	const ns = await ws.getObjectStream ()
	
	const a = []; for await (n of ns) {
		a.push (n)
		if (a.length === 2) break
	}
	
	expect (a).toStrictEqual (      [
        [
          '0',         'First Name',
          'Last Name', 'Gender',
          'Country',   'Age',
          'Date',      'Id'
        ],
        [
          '1',
          'Dulce',
          'Abril',
          'Female',
          'United States',
          '32',
          '15/10/2017',
          '1562'
        ]
      ]
    )
 	
})

test ('formats', async () => {

	const wb = await xlsx.open ('__data__/formats.xlsx')	
	const ws = wb.sheetByName.Sheet1
	
	const ns = await ws.getObjectStream ()
	
	const a = []; for await (n of ns) a.push (n)

	expect (a [0]).toStrictEqual (['2015-12-31', '23:59:00', '1.125', '1.125', 'Test', '1.2345'])

})

test ('inlineStr', async () => {

	const wb = await xlsx.open ('__data__/test-issue-1575.xlsx')	
	const ws = wb.sheets [0]
	
	const ns = await ws.getObjectStream ()
	
	const a = []; for await (n of ns) {
		a.push (n)
	}

	expect (a).toStrictEqual ([['A', 'B', 'C'], ['1.0', '2.0', '3.0'], ['4.0', '5.0', '6.0']])
 	
})

test ('skip', async () => {

	const wb = await xlsx.open ('__data__/d1.xlsx')	
	const ws = wb.sheets [0]
	
	const ns = await ws.getObjectStream ()
	
	const a = []; for await (n of ns) {
		a.push (n)
	}

	expect (a).toStrictEqual ([['', '', '', 'D1']])
 	
})
