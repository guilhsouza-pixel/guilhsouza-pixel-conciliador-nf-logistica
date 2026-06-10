import { useMemo, useState } from "react";
import { read, utils, writeFile } from "xlsx";
import { AlertTriangle, CheckCircle2, Download, FileSpreadsheet, RefreshCcw, Search, UploadCloud } from "lucide-react";

type Origem = "Enviadas" | "Recebidas";
type RawRow = Record<string, any>;
type ColMap = { nf?: string; serie?: string; fornecedor?: string; cnpj?: string; data?: string; valor?: string; quantidade?: string; item?: string };
type BaseRow = RawRow & { __origem: Origem; __linha: number; __nf: string; __serie: string; __fornecedor: string; __cnpj: string; __data: string; __valor: number; __quantidade: number; __item: string; __chave: string };
type Resultado = { conciliadas: RawRow[]; enviadasNaoRecebidas: RawRow[]; recebidasNaoEnviadas: RawRow[]; divergencias: RawRow[]; duplicidades: RawRow[]; resumo: Record<string, number | string>; enviadasTratadas: BaseRow[]; recebidasTratadas: BaseRow[]; colunasEnviadas: ColMap; colunasRecebidas: ColMap };

const aliases: Record<keyof ColMap, string[]> = {
  nf: ["nf", "nota", "nota fiscal", "nº nf", "numero nf", "número nf", "num nf", "documento", "doc", "nfe"],
  serie: ["serie", "série", "ser", "nf serie", "nf série", "series"],
  fornecedor: ["fornecedor", "nome fornecedor", "razao social", "razão social", "emitente", "cliente", "parceiro", "nome"],
  cnpj: ["cnpj", "cnpj fornecedor", "cnpj emitente", "cpf/cnpj", "cpf cnpj"],
  data: ["data", "data emissao", "data emissão", "data nf", "emissão", "dt emissao", "dt emissão", "data documento"],
  valor: ["valor", "valor nf", "valor total", "vlr", "vl total", "total", "montante", "valor documento"],
  quantidade: ["qtd", "qtde", "quantidade", "qde", "qty", "pecas", "peças", "volume"],
  item: ["item", "codigo", "código", "codigo material", "código material", "codigo embalagem", "código embalagem", "material", "embalagem", "produto"],
};

const normalize = (value: any) => String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim().replace(/\s+/g, " ");
const cleanText = (value: any) => String(value ?? "").trim().replace(/\s+/g, " ");
const onlyDigits = (value: any) => cleanText(value).replace(/\D/g, "");
const toNumber = (value: any) => {
  if (typeof value === "number") return value;
  const raw = cleanText(value);
  if (!raw) return 0;
  const n = Number(raw.replace(/R\$/gi, "").replace(/\./g, "").replace(",", ".").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
};
const toDateBR = (value: any) => {
  const raw = cleanText(value);
  if (!raw) return "";
  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) return parsed.toLocaleDateString("pt-BR");
  return raw;
};

function detectColumns(headers: string[]): ColMap {
  const output: ColMap = {};
  const hs = headers.map((h) => ({ original: h, normalized: normalize(h) }));
  (Object.keys(aliases) as (keyof ColMap)[]).forEach((field) => {
    const found = hs.find((h) => aliases[field].some((a) => h.normalized === normalize(a) || h.normalized.includes(normalize(a))));
    if (found) output[field] = found.original;
  });
  return output;
}

async function readExcel(file: File) {
  const buffer = await file.arrayBuffer();
  const workbook = read(buffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("A planilha não possui abas válidas.");
  const rows = utils.sheet_to_json<RawRow>(sheet, { defval: "", raw: false });
  return { rows, headers: rows.length ? Object.keys(rows[0]) : [] };
}

function treatRows(rows: RawRow[], cols: ColMap, origem: Origem): BaseRow[] {
  return rows.map((row, index) => {
    const nf = onlyDigits(row[cols.nf || ""]);
    const serie = onlyDigits(row[cols.serie || ""]);
    const fornecedor = normalize(row[cols.fornecedor || ""]);
    const cnpj = onlyDigits(row[cols.cnpj || ""]);
    const base = cnpj || fornecedor;
    return { ...row, __origem: origem, __linha: index + 2, __nf: nf, __serie: serie, __fornecedor: fornecedor, __cnpj: cnpj, __data: toDateBR(row[cols.data || ""]), __valor: toNumber(row[cols.valor || ""]), __quantidade: toNumber(row[cols.quantidade || ""]), __item: normalize(row[cols.item || ""]), __chave: [base, nf, serie].filter(Boolean).join("|") };
  }).filter((r) => r.__nf && r.__serie && (r.__cnpj || r.__fornecedor));
}

function rowOut(r: BaseRow, status: string) {
  return { Status: status, Chave: r.__chave, NF: r.__nf, Série: r.__serie, CNPJ: r.__cnpj, Fornecedor: r.__fornecedor, Data: r.__data, Valor: r.__valor, Quantidade: r.__quantidade, Item: r.__item, Linha: r.__linha };
}

function findDuplicates(rows: BaseRow[]) {
  const count = new Map<string, number>();
  rows.forEach((r) => count.set(r.__chave, (count.get(r.__chave) || 0) + 1));
  return rows.filter((r) => count.get(r.__chave)! > 1).map((r) => rowOut(r, `Duplicidade em ${r.__origem}`));
}

function compare(enviadas: BaseRow[], recebidas: BaseRow[], colunasEnviadas: ColMap, colunasRecebidas: ColMap): Resultado {
  const recMap = new Map(recebidas.map((r) => [r.__chave, r]));
  const envMap = new Map(enviadas.map((r) => [r.__chave, r]));
  const conciliadas: RawRow[] = [];
  const divergencias: RawRow[] = [];
  const enviadasNaoRecebidas: RawRow[] = [];
  const recebidasNaoEnviadas: RawRow[] = [];

  enviadas.forEach((e) => {
    const r = recMap.get(e.__chave);
    if (!r) { enviadasNaoRecebidas.push(rowOut(e, "Enviada não recebida")); return; }
    const divs: string[] = [];
    if (Math.abs(e.__valor - r.__valor) > 0.01) divs.push("Valor");
    if (Math.abs(e.__quantidade - r.__quantidade) > 0.01) divs.push("Quantidade");
    if (e.__data && r.__data && e.__data !== r.__data) divs.push("Data");
    if (e.__item && r.__item && e.__item !== r.__item) divs.push("Código/item");
    const out = { Chave: e.__chave, NF: e.__nf, Série: e.__serie, CNPJ: e.__cnpj || r.__cnpj, Fornecedor_Enviado: e.__fornecedor, Fornecedor_Recebido: r.__fornecedor, Valor_Enviado: e.__valor, Valor_Recebido: r.__valor, Quantidade_Enviada: e.__quantidade, Quantidade_Recebida: r.__quantidade, Data_Enviada: e.__data, Data_Recebida: r.__data, Item_Enviado: e.__item, Item_Recebido: r.__item };
    if (divs.length) divergencias.push({ Tipo: divs.join(" / "), ...out }); else conciliadas.push({ Status: "Conciliada", ...out });
  });

  recebidas.forEach((r) => { if (!envMap.has(r.__chave)) recebidasNaoEnviadas.push(rowOut(r, "Recebida não enviada")); });
  const duplicidades = [...findDuplicates(enviadas), ...findDuplicates(recebidas)];
  const total = enviadas.length || 1;
  return { conciliadas, enviadasNaoRecebidas, recebidasNaoEnviadas, divergencias, duplicidades, resumo: { totalEnviadas: enviadas.length, totalRecebidas: recebidas.length, conciliadas: conciliadas.length, enviadasNaoRecebidas: enviadasNaoRecebidas.length, recebidasNaoEnviadas: recebidasNaoEnviadas.length, divergencias: divergencias.length, duplicidades: duplicidades.length, percentualConciliacao: `${((conciliadas.length / total) * 100).toFixed(1)}%` }, enviadasTratadas: enviadas, recebidasTratadas: recebidas, colunasEnviadas, colunasRecebidas };
}

function exportExcel(res: Resultado) {
  const wb = utils.book_new();
  const tabs: [string, RawRow[]][] = [["Resumo", Object.entries(res.resumo).map(([Indicador, Valor]) => ({ Indicador, Valor }))], ["Conciliadas", res.conciliadas], ["Env nao recebidas", res.enviadasNaoRecebidas], ["Rec nao enviadas", res.recebidasNaoEnviadas], ["Divergencias", res.divergencias], ["Duplicidades", res.duplicidades], ["Base enviada", res.enviadasTratadas], ["Base recebida", res.recebidasTratadas]];
  tabs.forEach(([name, data]) => utils.book_append_sheet(wb, utils.json_to_sheet(data), name));
  writeFile(wb, `relatorio_conciliacao_nf_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function Metric({ label, value }: { label: string; value: number | string }) {
  return <div className="metric"><span>{label}</span><strong>{value}</strong></div>;
}

export default function App() {
  const [envFile, setEnvFile] = useState<File | null>(null);
  const [recFile, setRecFile] = useState<File | null>(null);
  const [resultado, setResultado] = useState<Resultado | null>(null);
  const [loading, setLoading] = useState(false);
  const [filtro, setFiltro] = useState("");
  const [erro, setErro] = useState("");

  async function conciliar() {
    setErro("");
    if (!envFile || !recFile) { setErro("Envie as duas planilhas para conciliar."); return; }
    setLoading(true);
    try {
      const env = await readExcel(envFile);
      const rec = await readExcel(recFile);
      if (!env.rows.length || !rec.rows.length) throw new Error("Uma das planilhas está vazia.");
      const envCols = detectColumns(env.headers);
      const recCols = detectColumns(rec.headers);
      if (!envCols.nf || !envCols.serie || (!envCols.cnpj && !envCols.fornecedor)) throw new Error("Não localizei NF, série e CNPJ/fornecedor na planilha de enviadas.");
      if (!recCols.nf || !recCols.serie || (!recCols.cnpj && !recCols.fornecedor)) throw new Error("Não localizei NF, série e CNPJ/fornecedor na planilha de recebidas.");
      const envTrat = treatRows(env.rows, envCols, "Enviadas");
      const recTrat = treatRows(rec.rows, recCols, "Recebidas");
      if (!envTrat.length || !recTrat.length) throw new Error("Não encontrei linhas válidas após tratar NF, série e fornecedor/CNPJ.");
      setResultado(compare(envTrat, recTrat, envCols, recCols));
    } catch (error: any) {
      setErro(error?.message || "Erro ao processar as planilhas.");
    } finally { setLoading(false); }
  }

  const tableRows = useMemo(() => {
    if (!resultado) return [];
    const all = [...resultado.enviadasNaoRecebidas, ...resultado.recebidasNaoEnviadas, ...resultado.divergencias];
    return all.filter((r) => JSON.stringify(r).toLowerCase().includes(filtro.toLowerCase()));
  }, [resultado, filtro]);

  return (
    <div className="page">
      <header className="header">
        <div className="brand"><FileSpreadsheet size={28} /><div><h1>Conciliação de Notas Fiscais</h1><p>Notas enviadas x notas recebidas para logística e embalagens</p></div></div>
      </header>
      <main className="container">
        <section className="card hero">
          <div><h2>Suba as duas planilhas e gere o relatório</h2><p>O app compara CNPJ + NF + Série. Se não houver CNPJ, usa Fornecedor + NF + Série.</p></div>
          <div className="grid2">
            <label className="upload"><UploadCloud /><strong>1. Notas Enviadas</strong><span>{envFile ? envFile.name : "Selecionar .xlsx ou .xls"}</span><input type="file" accept=".xlsx,.xls" onChange={(e) => setEnvFile(e.target.files?.[0] || null)} /></label>
            <label className="upload"><UploadCloud /><strong>2. Notas Recebidas</strong><span>{recFile ? recFile.name : "Selecionar .xlsx ou .xls"}</span><input type="file" accept=".xlsx,.xls" onChange={(e) => setRecFile(e.target.files?.[0] || null)} /></label>
          </div>
          {erro && <div className="error"><AlertTriangle size={18} />{erro}</div>}
          <div className="actions"><button className="primary" onClick={conciliar} disabled={loading}><Search size={18} />{loading ? "Conciliando..." : "Conciliar Notas"}</button><button className="secondary" onClick={() => { setResultado(null); setEnvFile(null); setRecFile(null); setErro(""); }}><RefreshCcw size={18} />Limpar</button></div>
        </section>

        {resultado && <>
          <section className="metrics"><Metric label="Notas enviadas" value={resultado.resumo.totalEnviadas} /><Metric label="Notas recebidas" value={resultado.resumo.totalRecebidas} /><Metric label="Conciliadas" value={resultado.resumo.conciliadas} /><Metric label="% Conciliação" value={resultado.resumo.percentualConciliacao} /><Metric label="Enviadas não recebidas" value={resultado.resumo.enviadasNaoRecebidas} /><Metric label="Recebidas não enviadas" value={resultado.resumo.recebidasNaoEnviadas} /><Metric label="Divergências" value={resultado.resumo.divergencias} /><Metric label="Duplicidades" value={resultado.resumo.duplicidades} /></section>
          <section className="card"><div className="tableHeader"><div><h2>Relatório de pendências</h2><p>Exibe divergências e notas não encontradas.</p></div><div className="rightActions"><input placeholder="Filtrar por NF, fornecedor, CNPJ..." value={filtro} onChange={(e) => setFiltro(e.target.value)} /><button className="primary" onClick={() => exportExcel(resultado)}><Download size={18} />Baixar Excel</button></div></div><div className="tableWrap"><table><thead><tr>{["Status/Tipo", "NF", "Série", "CNPJ", "Fornecedor", "Valor", "Quantidade", "Data"].map((h) => <th key={h}>{h}</th>)}</tr></thead><tbody>{tableRows.slice(0, 500).map((r, i) => <tr key={i}><td><b>{String(r.Status || r.Tipo || "")}</b></td><td>{String(r.NF || "")}</td><td>{String(r.Série || "")}</td><td>{String(r.CNPJ || "")}</td><td>{String(r.Fornecedor || r.Fornecedor_Enviado || "")}</td><td>{String(r.Valor || r.Valor_Enviado || "")}</td><td>{String(r.Quantidade || r.Quantidade_Enviada || "")}</td><td>{String(r.Data || r.Data_Enviada || "")}</td></tr>)}</tbody></table></div><p className="hint"><CheckCircle2 size={16} /> O Excel exportado contém resumo, conciliadas, divergências, duplicidades e bases tratadas.</p></section>
        </>}
      </main>
    </div>
  );
}
