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

function safeDateFromExcelSerial(serial) {
	// Excel serial date (days since 1899-12-30); handle if numeric
	const n = coerceNumber(serial);
	if (!Number.isFinite(n)) return null;
	const epoch = new Date(Date.UTC(1899, 11, 30));
	const millis = epoch.getTime() + n * 24 * 60 * 60 * 1000;
	return new Date(millis);
}

function groupBy(items, key) {
	const map = new Map();
	for (const it of items) {
		const k = it[key];
		if (!map.has(k)) map.set(k, []);
		map.get(k).push(it);
	}
	return map;
}

function calcSummary(rows) {
	const n = rows.length;
	const metragem = rows.map(r => coerceNumber(r.metragem_m2)).filter(Number.isFinite);
	const custo_m2 = rows.map(r => coerceNumber(r.custo_m2_brl)).filter(Number.isFinite);
	const preco_total = rows.map(r => coerceNumber(r.preco_total_brl)).filter(Number.isFinite);
	const renda = rows.map(r => coerceNumber(r.renda_familiar_media_mensal_brl)).filter(Number.isFinite);
	const quartos = rows.map(r => coerceNumber(r.quartos)).filter(Number.isFinite);

	const avg = arr => arr.length ? arr.reduce((a,b)=>a+b,0)/arr.length : null;
	const med = arr => {
		if (!arr.length) return null;
		const s = [...arr].sort((a,b)=>a-b);
		const mid = Math.floor(s.length/2);
		return s.length % 2 ? s[mid] : (s[mid-1]+s[mid])/2;
	};
	const min = arr => arr.length ? Math.min(...arr) : null;
	const max = arr => arr.length ? Math.max(...arr) : null;

	function corr(x, y) {
		const m = Math.min(x.length, y.length);
		if (m < 3) return null;
		// align by index (drop NaNs already)
		const xs = x.slice(0, m);
		const ys = y.slice(0, m);
		const meanX = avg(xs);
		const meanY = avg(ys);
		let num = 0, denX = 0, denY = 0;
		for (let i = 0; i < m; i++) {
			const dx = xs[i] - meanX;
			const dy = ys[i] - meanY;
			num += dx * dy;
			denX += dx * dx;
			denY += dy * dy;
		}
		const den = Math.sqrt(denX * denY);
		return den ? num / den : null;
	}

	// mix de quartos
	const mixQuartos = {};
	for (const r of rows) {
		const q = coerceNumber(r.quartos);
		if (!Number.isFinite(q)) continue;
		mixQuartos[q] = (mixQuartos[q] || 0) + 1;
	}
	const totalMix = Object.values(mixQuartos).reduce((a,b)=>a+b,0) || 1;
	for (const k of Object.keys(mixQuartos)) {
		mixQuartos[k] = mixQuartos[k] / totalMix;
	}

	return {
		count: n,
		metragem_m2: { avg: avg(metragem), med: med(metragem), min: min(metragem), max: max(metragem) },
		custo_m2_brl: { avg: avg(custo_m2), med: med(custo_m2), min: min(custo_m2), max: max(custo_m2) },
		preco_total_brl: { avg: avg(preco_total), med: med(preco_total), min: min(preco_total), max: max(preco_total) },
		renda_familiar_media_mensal_brl: { avg: avg(renda), med: med(renda), min: min(renda), max: max(renda) },
		quartos: { avg: avg(quartos), med: med(quartos), min: min(quartos), max: max(quartos) },
		corr_preco_renda: corr(preco_total, renda),
		corr_custo_m2_renda: corr(custo_m2, renda),
		mix_quartos: mixQuartos,
	};
}

function build() {
	if (!fs.existsSync(EXCEL_PATH)) {
		throw new Error('Arquivo não encontrado: ' + EXCEL_PATH);
	}
	const wb = xlsx.readFile(EXCEL_PATH);
	const ws = wb.Sheets[wb.SheetNames[0]];
	const rows = xlsx.utils.sheet_to_json(ws, { defval: null });

	const byCidade = Object.fromEntries([...groupBy(rows, 'cidade')].map(([k, v]) => [k, calcSummary(v)]));
	const byEstado = Object.fromEntries([...groupBy(rows, 'estado')].map(([k, v]) => [k, calcSummary(v)]));
	const byRegiao = Object.fromEntries([...groupBy(rows, 'regiao')].map(([k, v]) => [k, calcSummary(v)]));
	const total = calcSummary(rows);

	// Top cities by custo_m2 avg and by ticket médio
	function topN(obj, path, n=10) {
		return Object.entries(obj)
			.map(([k, s]) => ({ key: k, value: path.reduce((a,p)=>a && a[p], s) }))
			.filter(r => Number.isFinite(r.value))
			.sort((a,b)=>b.value - a.value)
			.slice(0, n);
	}
	const topCidadesCustoM2 = topN(byCidade, ['custo_m2_brl', 'avg'], 15);
	const topCidadesTicket = topN(byCidade, ['preco_total_brl', 'avg'], 15);

	// Insights automatizados
	const insights = [];
	if (Number.isFinite(total.corr_custo_m2_renda)) {
		insights.push(`Correlação custo/m² vs renda: ${total.corr_custo_m2_renda.toFixed(2)} (nível de associação do preço com renda local)`);
	}
	if (Number.isFinite(total.corr_preco_renda)) {
		insights.push(`Correlação ticket vs renda: ${total.corr_preco_renda.toFixed(2)} (capacidade de pagamento)`);
	}
	const maiorMixQuarto = Object.entries(total.mix_quartos).sort((a,b)=>b[1]-a[1])[0];
	if (maiorMixQuarto) {
		insights.push(`Maior demanda por nº de quartos: ${maiorMixQuarto[0]} quartos (${(maiorMixQuarto[1]*100).toFixed(1)}% do portfólio)`);
	}
	if (topCidadesCustoM2.length) {
		insights.push(`Capitais com maior custo/m²: ${topCidadesCustoM2.slice(0,5).map(x=>x.key).join(', ')}`);
	}
	if (topCidadesTicket.length) {
		insights.push(`Capitais com maior ticket médio: ${topCidadesTicket.slice(0,5).map(x=>x.key).join(', ')}`);
	}

	const out = {
		updatedAt: new Date().toISOString(),
		total,
		byCidade,
		byEstado,
		byRegiao,
		lists: {
			topCidadesCustoM2,
			topCidadesTicket
		},
		insights
	};

	const outDir = path.resolve(__dirname, 'data');
	fs.mkdirSync(outDir, { recursive: true });
	fs.writeFileSync(path.join(outDir, 'kpis.json'), JSON.stringify(out, null, 2));
	console.log('Gerado:', path.join(outDir, 'kpis.json'));
}

build();