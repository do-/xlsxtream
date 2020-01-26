const xxx = require ('../index.js')
const fs  = require ('fs')
const Saxophone = require ('saxophone')

test ('scanStyles', async () => {

	let n = 0

	await xxx.scanStyles (
	
		(path) => fs.createReadStream ('./__data__/' + path),
		
		(node) => n += parseInt (Saxophone.parseAttrs (node.attrs).numFmtId)
	
	)

	expect (n).toBe (14);

})
