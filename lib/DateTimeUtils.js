const MS_IN_DAY = 24 * 60 * 60 * 1000

const BASE_DATE = new Date ('1900-01-01T00:00:00Z')

const toYMD = v => {

	const n = parseInt (v); if (isNaN (n)) return null

	const d = new Date (Math.round ((n - 25569) * MS_IN_DAY))

    return d.toJSON ().slice (0, 10)

}

const toHMS = v => {

	const f = parseFloat (v); if (isNaN (f)) return null

	const d = new Date (BASE_DATE)
	
	d.setMilliseconds (MS_IN_DAY * f)
	
	return d.toJSON ().slice (11, 19)

}

module.exports = {
	toYMD,
	toHMS,
}
