const xlsx = require ('..')
const {StringDecoder} = require ('string_decoder'); 

test ('basic', async () => {

	const wb = await xlsx.open ('__data__/2_____with_chart.xlsx')
	
	const [ws] = wb.sheets
	
	const is = await ws.getXMLStream ()
	
	const d = new StringDecoder ()
	
	let s = ''
	
	for await (b of is) s += d.write (b)
	
	s += d.end ()
	
	expect (s.slice (0, 6)).toBe ('<?xml ')

	expect (ws.position).toBe (ws.size)
 	
})
