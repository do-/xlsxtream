const xlsx = require ('..')

test ('2_____with_chart', async () => {

	const wb = await xlsx.open ('__data__/1_____loreyna126.xlsx')	
	const ws = wb.sheetByName.POST_DSENDS
	
	const r = await ws.getLastRow ()

	expect (r.num).toBe (798)
	expect (r.index).toBe (798)
		
})

