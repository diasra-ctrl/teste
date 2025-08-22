const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const EXCEL_PATH = path.resolve(__dirname, 'base_casas_capitais_brasil_5000.xlsx');

function coerceNumber(value) {
	if (value === null || value === undefined) return NaN;
	if (typeof value === 'number') return value;
	if (typeof value === 'string') {
		const cleaned = value
			.replace(/\s/g, '')
			.replace(/[Rr]\$/g, '')
			.replace(/\./g, '')
			.replace(/,/g, '.');
		const n = Number(cleaned);
		return Number.isFinite(n) ? n : NaN;
	}
	return NaN;
}

function main() {
	if (!fs.existsSync(EXCEL_PATH)) {
		console.error('Arquivo não encontrado:', EXCEL_PATH);
		process.exit(1);
	}
	const wb = xlsx.readFile(EXCEL_PATH);
	const sheetNames = wb.SheetNames;
	console.log('SHEETS:', JSON.stringify(sheetNames));
	const ws = wb.Sheets[sheetNames[0]];
	const rows = xlsx.utils.sheet_to_json(ws, { defval: null });
	console.log('ROWS:', rows.length);
	if (rows.length === 0) return;

	// Columns from header
	const columns = Object.keys(rows[0]);
	console.log('COLUMNS:', JSON.stringify(columns));
	console.log('SAMPLE_ROW:', JSON.stringify(rows[0], null, 2));

	// Numeric profile per column
	const limit = Math.min(rows.length, 2000);
	const profile = {};
	for (const col of columns) {
		let nonNull = 0;
		let numericCount = 0;
		let sum = 0;
		let min = Number.POSITIVE_INFINITY;
		let max = Number.NEGATIVE_INFINITY;
		for (let i = 0; i < limit; i++) {
			const v = rows[i][col];
			if (v === null || v === undefined || v === '') continue;
			nonNull++;
			const n = coerceNumber(v);
			if (Number.isFinite(n)) {
				numericCount++;
				sum += n;
				if (n < min) min = n;
				if (n > max) max = n;
			}
		}
		profile[col] = {
			nonNull,
			numeric: numericCount,
			avg: numericCount ? sum / numericCount : null,
			min: numericCount ? min : null,
			max: numericCount ? max : null,
		};
	}
	console.log('PROFILE:', JSON.stringify(profile, null, 2));

	fs.mkdirSync(path.resolve(__dirname, 'data'), { recursive: true });
	fs.writeFileSync(
		path.resolve(__dirname, 'data', 'profile.json'),
		JSON.stringify({ sheetNames, rowCount: rows.length, columns, sample: rows[0], profile }, null, 2),
		'utf8'
	);
}

main();