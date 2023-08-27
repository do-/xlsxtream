const RowTransformer = require ('../lib/RowTransformer.js')

test ('error', async () => {

	const t = new RowTransformer ({})
		
	await expect (new Promise ((ok, fail) => {
	
		t.on ('error', fail)
		t.on ('end', ok)
	
		t.end ({})		
	})).rejects.toThrow ('nnot re')
	 	
})
