const MS_IN_DAY = 24 * 60 * 60 * 1000
const BASE_DATE = new Date ('1900-01-01T00:00:00Z')

const COL_IDX_BASE  = 'A'.charCodeAt (0) - 1
const COL_IDX_RADIX = 'Z'.charCodeAt (0) - COL_IDX_BASE

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

const colIdx = s => {

	const {length} = s

	let v = 0; for (let i = 0; i < length; i ++) {

		const delta = s.charCodeAt (i) - COL_IDX_BASE; if (delta < 0) break

		v *= COL_IDX_RADIX

		v += delta

	}

	return v

}

module.exports = {
	toYMD,
	toHMS,
	colIdx
}
