const {XMLNode} = require ('xml-toolkit')
const xlsx = require ('..')

test ('multiple', async () => {

	const wb = await xlsx.open ('__data__/2_____with_chart.xlsx')	
	const ws = wb.sheetByName.Tabelle1
	
	const ns = await ws.getXMLNodeStream ()
	
	const dump = XMLNode.toObject (), a = []
		
	for await (n of ns) a.push (dump (n))

	expect (a).toStrictEqual (
      [
        {
          "r": "1",
          "spans": "1:2",
          "c": [
            {
              "r": "A1",
              "v": "1"
            },
            {
              "r": "B1",
              "v": "10"
            }
          ]
        },
        {
          "r": "2",
          "spans": "1:2",
          "c": [
            {
              "r": "A2",
              "v": "2"
            },
            {
              "r": "B2",
              "v": "20"
            }
          ]
        }
      ]		
	)
 	
})
