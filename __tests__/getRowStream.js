const xlsx = require ('..')

test ('2_____with_chart', async () => {

	const wb = await xlsx.open ('__data__/1_____loreyna126.xlsx')	
	const ws = wb.sheetByName.POST_DSENDS
	
	const r = await ws.getLastRow ()

	expect (r.num).toBe (798)
	expect (r.index).toBe (798)
		
})

test ('err', async () => {

	const wb = await xlsx.open ('__data__/err.xlsx')

    const rows = await wb.sheets [0].getRowStream ()                        

    ROWS: for await (const {index, cells} of rows) if (index === 4) {

		for (const cell of cells) if (cell.index === 2) {

			const v = cell.valueOf ()

			expect (parseFloat (v) === 1946003509).toBe (true)

			expect (v == 1946003509).toBe (true)

			expect (v).toBe ('1.946003509E9')

		}

	}
		
})
