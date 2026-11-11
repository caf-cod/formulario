"use client";

import { useMemo, useState } from "react";
import { jsPDF } from "jspdf";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { BookOpen, Check, CheckCircle2, ChevronLeft, ChevronRight, Download, ExternalLink, FileCheck2, FileSpreadsheet, FileText, Mail, Search, ShieldCheck, X } from "lucide-react";
import data from "./data/caf-data.json";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Progress } from "@/components/ui/progress";
import { Tabs, TabsList, TabsTrigger } from "@/components/ui/tabs";

type Kind = "pf" | "pj";
type Variable = { variable: string; description: string; values: string[] };
type Section = { name: string; items: Variable[] };
type Handler = { id: number; name: string; role: string; cpf: string; email: string };
type FormData = { institution: string; address: string; city: string; uf: string; process: string; representative: string; role: string; nationality: string; cpf: string; email: string; phone: string; purpose: string; place: string; date: string; territory: string; period: string; audience: string };
type DataClassification = "personal" | "sensitive" | "contextual" | "nonpersonal";

const steps = ["Dicionário", "Variáveis", "Solicitante", "Documentos"];
const localDateValue = () => {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};
const initialForm: FormData = { institution: "", address: "", city: "", uf: "", process: "", representative: "", role: "", nationality: "Brasileira", cpf: "", email: "", phone: "", purpose: "", place: "", date: localDateValue(), territory: "", period: "", audience: "" };
const destinationEmail = "caf@mda.gov.br";
const classificationMeta: Record<DataClassification, { label: string; short: string; color: string; className: string; description: string }> = {
  personal: { label: "Dado pessoal", short: "Pessoal", color: "FFE600", className: "border-yellow-400 bg-yellow-100 text-yellow-900", description: "Identifica ou pode identificar diretamente uma pessoa natural." },
  sensitive: { label: "Dado pessoal sensível", short: "Sensível", color: "FF0000", className: "border-red-600 bg-red-600 text-white", description: "Revela categoria protegida pelo art. 5º, II, da LGPD." },
  contextual: { label: "Depende do contexto", short: "Contextual", color: "F4B183", className: "border-orange-400 bg-orange-100 text-orange-900", description: "Pode identificar uma pessoa quando combinado com outros dados." },
  nonpersonal: { label: "Dado não pessoal", short: "Não pessoal", color: "E7E6E6", className: "border-slate-300 bg-slate-100 text-slate-700", description: "Não se relaciona, isoladamente, a pessoa natural identificada ou identificável." },
};
const sensitiveVariables = new Set(["nm_tipo_etnia"]);
const personalVariables = new Set(["nr_cpf", "nm_usuario", "nm_pessoa_fisica", "dt_nascimento", "nm_sexo", "nm_mae", "nm_social", "nm_tipo_estado_civil", "nm_nacionalidade", "nr_documento", "nm_emissor_orgao", "sg_uf_rg", "nm_tipo_membro_familiar", "ds_tipo_escolaridade", "ds_email", "nr_telefone", "nr_cpf_t", "nm_pessoa_fisica_t"]);
const contextualVariables = new Set(["nr_caf", "ds_tipo_enquadramento_renda", "ds_situacao_unidade_familiar", "st_trabalha_ufpr", "nm_tipo_terreno_ufpr", "nm_caracterizacao_area", "nr_cep", "cd_municipio", "sg_uf", "nm_municipio", "ds_logradouro", "nr_logradouro", "ds_referencia", "ds_complemento", "ds_tipo_area", "ds_tipo_unidade_medida", "nr_area", "ds_tipo_localizacao_area", "ds_condicao_dominio", "st_imovel_principal", "nr_latitude", "nr_longitude", "QUANTIDADE_MAO_DE_OBRA_FAMILIAR", "QUANTIDADE_MAO_DE_OBRA_CONTRATADA", "categoria_renda", "ds_tipo_renda", "ds_produto", "vl_renda_auferida", "vl_renda_estimada"]);
function classifyVariable(variable: string, kind: Kind): DataClassification {
  if (sensitiveVariables.has(variable)) return "sensitive";
  if (personalVariables.has(variable)) return "personal";
  if (contextualVariables.has(variable)) return "contextual";
  if (kind === "pj") return "nonpersonal";
  return "contextual";
}
const normalize = (value: string) => value.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
const digits = (value: string) => value.replace(/\D/g, "");
const formatCpf = (value: string) => digits(value).slice(0, 11).replace(/(\d{3})(\d)/, "$1.$2").replace(/(\d{3})(\d)/, "$1.$2").replace(/(\d{3})(\d{1,2})$/, "$1-$2");
const formatSeiProcess = (value: string) => {
  const number = digits(value).slice(0, 17);
  return number
    .replace(/^(\d{5})(\d)/, "$1.$2")
    .replace(/^(\d{5}\.\d{6})(\d)/, "$1/$2")
    .replace(/^(\d{5}\.\d{6}\/\d{4})(\d)/, "$1-$2");
};
const formatPhone = (value: string) => {
  const number = digits(value).slice(0, 11);
  if (!number) return "";
  if (number.length <= 2) return `(${number}`;
  const subscriber = number.slice(2);
  const split = subscriber.length > 8 ? 5 : 4;
  return `(${number.slice(0, 2)}) ${subscriber.slice(0, split)}${subscriber.length > split ? `-${subscriber.slice(split)}` : ""}`;
};
const formatDate = (value: string) => value ? new Intl.DateTimeFormat("pt-BR").format(new Date(`${value}T12:00:00`)) : "";

function validateCpf(value: string) {
  const cpf = digits(value);
  if (cpf.length !== 11 || /^(\d)\1+$/.test(cpf)) return false;
  const calc = (length: number) => { let total = 0; for (let i = 0; i < length; i++) total += Number(cpf[i]) * (length + 1 - i); const rest = (total * 10) % 11; return rest === 10 ? 0 : rest; };
  return calc(9) === Number(cpf[9]) && calc(10) === Number(cpf[10]);
}

function pdfLines(doc: jsPDF, text: string, y: number, options?: { bold?: boolean; center?: boolean; size?: number }) {
  doc.setFont("helvetica", options?.bold ? "bold" : "normal"); doc.setFontSize(options?.size ?? 10.5);
  const lines = doc.splitTextToSize(text, 172); if (y + lines.length * 5.2 > 278) { doc.addPage(); y = 22; }
  doc.text(lines, options?.center ? 105 : 19, y, { align: options?.center ? "center" : "left" }); return y + lines.length * 5.2 + 2.5;
}

function loadPdfHeader() {
  return new Promise<string>((resolve, reject) => {
    const image = new Image();
    image.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = 792; canvas.height = 145;
      const context = canvas.getContext("2d");
      if (!context) return reject(new Error("Cabeçalho indisponível"));
      context.fillStyle = "#ffffff"; context.fillRect(0, 0, canvas.width, canvas.height);
      context.drawImage(image, 4, 35, 792, 145, 0, 0, 792, 145);
      resolve(canvas.toDataURL("image/png"));
    };
    image.onerror = () => reject(new Error("Cabeçalho indisponível"));
    image.src = `${import.meta.env.BASE_URL}assets/brasao-cabecalho.png`;
  });
}

const cloneFormat = <T,>(value: T): T => value ? JSON.parse(JSON.stringify(value)) : value;

function createCell(value: string, source?: XLSX.CellObject): XLSX.CellObject {
  const cell: XLSX.CellObject = { t: "s", v: value };
  if (source?.s) cell.s = cloneFormat(source.s);
  return cell;
}

function buildSelectionSheet(source: XLSX.WorkSheet, items: Array<Variable & { section: string; key: string }>, justifications: Record<string, string>, kind: Kind) {
  const sheet: XLSX.WorkSheet = cloneFormat(source);
  const sourceRows = XLSX.utils.sheet_to_json<(string | null)[]>(source, { header: 1, raw: false }) as (string | null)[][];
  let section = "";
  for (let row = 0; row < sourceRows.length; row++) {
    const first = String(sourceRows[row]?.[0] || "").trim();
    const second = String(sourceRows[row]?.[1] || "").trim();
    const third = String(sourceRows[row]?.[2] || "").trim();
    if (first && !second && !third && String(sourceRows[row + 1]?.[0] || "").trim() === "VARIÁVEL") {
      section = first;
      continue;
    }
    if (first === "VARIÁVEL") {
      const address = XLSX.utils.encode_cell({ r: row, c: 1 });
      sheet[address] = createCell("SELEÇÃO", source[address]);
      continue;
    }
    if (!first) continue;
    const item = items.find((candidate) => candidate.section === section && candidate.variable === first);
    const selectionAddress = XLSX.utils.encode_cell({ r: row, c: 1 });
    const justificationAddress = XLSX.utils.encode_cell({ r: row, c: 2 });
    sheet[selectionAddress] = createCell(item ? "SIM" : "NÃO", source[selectionAddress]);
    sheet[justificationAddress] = createCell(item ? justifications[item.key].trim() : "-", source[justificationAddress]);
    const classification = classifyVariable(first, kind);
    for (let col = 0; col < 3; col++) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      if (!sheet[address]) sheet[address] = createCell("", source[address]);
      sheet[address].s = { ...(sheet[address].s || {}), fill: { patternType: "solid", fgColor: { rgb: classificationMeta[classification].color } } };
    }
  }
  return sheet;
}

function buildRequestDataSheet(source: XLSX.WorkSheet, form: FormData, handlers: Handler[]) {
  const sheet = XLSX.utils.aoa_to_sheet([]);
  const sectionStyle = [source.A4, source.B4, source.C4];
  const labelStyle = source.A5;
  const valueStyle = source.B5;
  sheet["!cols"] = cloneFormat(source["!cols"] || [{ wch: 30 }, { wch: 42 }, { wch: 42 }]);
  sheet["!rows"] = cloneFormat(source["!rows"] || []);
  sheet["!margins"] = cloneFormat(source["!margins"]);
  sheet["!merges"] = [];
  let row = 2;
  const mergedRow = (text: string, styles = sectionStyle) => {
    for (let col = 0; col < 3; col++) sheet[XLSX.utils.encode_cell({ r: row, c: col })] = createCell(col === 0 ? text : "", styles[col]);
    sheet["!merges"]!.push({ s: { r: row, c: 0 }, e: { r: row, c: 2 } });
    row += 1;
  };
  const field = (label: string, value: string) => {
    sheet[XLSX.utils.encode_cell({ r: row, c: 0 })] = createCell(label, labelStyle);
    sheet[XLSX.utils.encode_cell({ r: row, c: 1 })] = createCell(value || "Não informado", valueStyle);
    sheet[XLSX.utils.encode_cell({ r: row, c: 2 })] = createCell("", valueStyle);
    sheet["!merges"]!.push({ s: { r: row, c: 1 }, e: { r: row, c: 2 } });
    row += 1;
  };

  mergedRow("DADOS DA SOLICITAÇÃO", [source.A3, source.B3, source.C3]);
  row += 1;
  mergedRow("INSTITUIÇÃO SOLICITANTE");
  field("Instituição ou órgão", form.institution);
  field("Endereço completo", form.address);
  field("Cidade/UF", `${form.city}/${form.uf}`);
  field("Processo SEI", form.process);
  field("Finalidade do uso", form.purpose);
  field("Recorte territorial", form.territory);
  field("Referência temporal", form.period);
  field("Recorte de público", form.audience);
  field("Local e data", `${form.place}, ${formatDate(form.date)}`);
  row += 1;
  mergedRow("RESPONSÁVEL PELA INSTITUIÇÃO E PELOS DADOS");
  field("Nome", form.representative);
  field("Cargo ou função", form.role);
  field("Nacionalidade", form.nationality);
  field("CPF", form.cpf);
  field("E-mail", form.email);
  field("Telefone", form.phone);
  row += 1;
  mergedRow("TRATADORES DE DADOS");
  handlers.forEach((person, index) => {
    mergedRow(`TRATADOR(A) ${index + 1}`);
    field("Nome", person.name);
    field("Cargo ou função", person.role);
    field("CPF", person.cpf);
    field("E-mail", person.email);
    if (index < handlers.length - 1) row += 1;
  });
  sheet["!ref"] = `A1:C${row}`;
  return sheet;
}

const spreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

function rowCell(row: Element, column: string) {
  return Array.from(row.getElementsByTagNameNS(spreadsheetNamespace, "c")).find((cell) => cell.getAttribute("r")?.replace(/\d/g, "") === column);
}

function cellText(row: Element, column: string) {
  return rowCell(row, column)?.getElementsByTagNameNS(spreadsheetNamespace, "t")[0]?.textContent || "";
}

function setRowCell(document: XMLDocument, row: Element, column: string, value: string) {
  const rowNumber = row.getAttribute("r") || "1";
  let cell = rowCell(row, column);
  if (!cell) {
    cell = document.createElementNS(spreadsheetNamespace, "c");
    cell.setAttribute("r", `${column}${rowNumber}`);
    row.appendChild(cell);
  }
  cell.setAttribute("t", "inlineStr");
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  const inline = document.createElementNS(spreadsheetNamespace, "is");
  const text = document.createElementNS(spreadsheetNamespace, "t");
  text.textContent = value;
  inline.appendChild(text); cell.appendChild(inline);
}

function renumberRow(row: Element, rowNumber: number) {
  row.setAttribute("r", String(rowNumber));
  Array.from(row.getElementsByTagNameNS(spreadsheetNamespace, "c")).forEach((cell) => {
    const reference = cell.getAttribute("r") || "A1";
    cell.setAttribute("r", `${reference.replace(/\d/g, "")}${rowNumber}`);
  });
}

async function buildExcelFromTemplate(template: ArrayBuffer, form: FormData, handlers: Handler[], selected: Record<string, string>) {
  const zip = await JSZip.loadAsync(template);
  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  const requestPath = "xl/worksheets/sheet1.xml";
  const requestFile = zip.file(requestPath);
  if (!requestFile) throw new Error("Aba de dados não encontrada");
  const requestDocument = parser.parseFromString(await requestFile.async("string"), "application/xml");
  const requestRows = Array.from(requestDocument.getElementsByTagNameNS(spreadsheetNamespace, "row"));
  const requestRow = (number: number) => requestRows.find((row) => row.getAttribute("r") === String(number));
  const fixedValues: Record<number, string> = {
    5: form.institution, 6: form.address, 7: `${form.city}/${form.uf}`, 8: form.process,
    9: form.purpose, 10: form.territory, 11: form.period, 12: form.audience,
    13: `${form.place}, ${formatDate(form.date)}`, 16: form.representative, 17: form.role,
    18: form.nationality, 19: form.cpf, 20: form.email, 21: form.phone,
  };
  Object.entries(fixedValues).forEach(([number, value]) => {
    const row = requestRow(Number(number)); if (row) setRowCell(requestDocument, row, "B", value || "Não informado");
  });

  const sheetData = requestDocument.getElementsByTagNameNS(spreadsheetNamespace, "sheetData")[0];
  const templates = [24, 25, 26, 27, 28, 29].map((number) => requestRow(number)?.cloneNode(true) as Element);
  Array.from(sheetData.getElementsByTagNameNS(spreadsheetNamespace, "row")).forEach((row) => {
    if (Number(row.getAttribute("r")) >= 24) row.remove();
  });
  let currentRow = 24;
  handlers.forEach((person, index) => {
    const values = [`TRATADOR(A) ${index + 1}`, person.name, person.role, person.cpf, person.email];
    for (let offset = 0; offset < 5; offset++) {
      const row = templates[offset].cloneNode(true) as Element;
      renumberRow(row, currentRow + offset);
      if (offset === 0) setRowCell(requestDocument, row, "A", values[offset]);
      else setRowCell(requestDocument, row, "B", values[offset]);
      sheetData.appendChild(row);
    }
    currentRow += 5;
    if (index < handlers.length - 1) {
      const spacer = templates[5].cloneNode(true) as Element;
      renumberRow(spacer, currentRow); sheetData.appendChild(spacer); currentRow += 1;
    }
  });
  const lastRow = currentRow - 1;
  const mergeCells = requestDocument.getElementsByTagNameNS(spreadsheetNamespace, "mergeCells")[0];
  Array.from(mergeCells.getElementsByTagNameNS(spreadsheetNamespace, "mergeCell")).forEach((merge) => {
    const numbers = (merge.getAttribute("ref") || "").match(/\d+/g)?.map(Number) || [];
    if (numbers.some((number) => number >= 24)) merge.remove();
  });
  let handlerRow = 24;
  handlers.forEach((_, index) => {
    [`A${handlerRow}:C${handlerRow}`, ...[1, 2, 3, 4].map((offset) => `B${handlerRow + offset}:C${handlerRow + offset}`)].forEach((reference) => {
      const merge = requestDocument.createElementNS(spreadsheetNamespace, "mergeCell"); merge.setAttribute("ref", reference); mergeCells.appendChild(merge);
    });
    handlerRow += index < handlers.length - 1 ? 6 : 5;
  });
  mergeCells.setAttribute("count", String(mergeCells.getElementsByTagNameNS(spreadsheetNamespace, "mergeCell").length));
  requestDocument.getElementsByTagNameNS(spreadsheetNamespace, "dimension")[0]?.setAttribute("ref", `A1:C${lastRow}`);
  zip.file(requestPath, serializer.serializeToString(requestDocument));

  for (const [sheetIndex, kind] of [[2, "pf"], [3, "pj"]] as const) {
    const path = `xl/worksheets/sheet${sheetIndex}.xml`;
    const file = zip.file(path); if (!file) throw new Error("Aba de variáveis não encontrada");
    const document = parser.parseFromString(await file.async("string"), "application/xml");
    const rows = Array.from(document.getElementsByTagNameNS(spreadsheetNamespace, "row"));
    let section = "";
    rows.forEach((row, index) => {
      const first = cellText(row, "A").trim();
      const next = rows[index + 1] ? cellText(rows[index + 1], "A").trim() : "";
      if (first && next === "VARIÁVEL") { section = first; return; }
      if (first === "VARIÁVEL") { setRowCell(document, row, "B", "SELEÇÃO"); return; }
      if (!first || !section) return;
      const key = `${kind}|${section}|${first}`;
      const chosen = Object.prototype.hasOwnProperty.call(selected, key);
      setRowCell(document, row, "B", chosen ? "SIM" : "NÃO");
      setRowCell(document, row, "C", chosen ? selected[key].trim() : "-");
    });
    zip.file(path, serializer.serializeToString(document));
  }
  return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url; link.download = filename; document.body.appendChild(link); link.click(); link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export default function Home() {
  const [step, setStep] = useState(0); const [kind, setKind] = useState<Kind>("pf"); const [search, setSearch] = useState("");
  const [selected, setSelected] = useState<Record<string, string>>({}); const [form, setForm] = useState<FormData>(initialForm); const [notice, setNotice] = useState("");
  const [noticeError, setNoticeError] = useState(false); const [showFormErrors, setShowFormErrors] = useState(false);
  const [handlers, setHandlers] = useState<Handler[]>([{ id: 1, name: "", role: "", cpf: "", email: "" }]);
  const [nextHandlerId, setNextHandlerId] = useState(2);
  const municipalities = data.municipalities as Record<string, string[]>;
  const sections = data[kind] as Section[];
  const variables = useMemo(() => sections.flatMap((section) => section.items.map((item) => ({ ...item, section: section.name, kind, key: `${kind}|${section.name}|${item.variable}` }))), [sections]);
  const shown = useMemo(() => { const q = normalize(search); return variables.filter((item) => !q || normalize(`${item.variable} ${item.description} ${item.section} ${item.values.join(" ")}`).includes(q)); }, [variables, search]);
  const allVariables = useMemo(() => (["pf", "pj"] as Kind[]).flatMap((modality) => (data[modality] as Section[]).flatMap((section) => section.items.map((item) => ({ ...item, kind: modality, section: section.name, key: `${modality}|${section.name}|${item.variable}` })))), []);
  const selectedItems = allVariables.filter((item) => selected[item.key] !== undefined);
  const missingJustifications = selectedItems.filter((item) => !selected[item.key]?.trim());
  const modalities = (["pf", "pj"] as Kind[]).filter((modality) => selectedItems.some((item) => item.kind === modality));
  const modalityLabel = modalities.map((modality) => modality.toUpperCase()).join(" e "); const cpfOkay = !form.cpf || validateCpf(form.cpf);
  const validEmail = (value: string) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.trim());
  const emailOkay = !form.email || validEmail(form.email);
  const formReady = Boolean(form.institution.trim() && form.address.trim() && form.city.trim() && form.uf && form.representative.trim() && form.role.trim() && form.nationality.trim() && form.cpf && cpfOkay && form.email.trim() && emailOkay && form.purpose.trim() && form.territory.trim() && form.period.trim() && form.audience.trim() && form.date && handlers.length && handlers.every((person) => person.name.trim() && person.role.trim() && validateCpf(person.cpf) && validEmail(person.email)));
  const changeKind = (value: Kind) => { setKind(value); setSearch(""); };
  const updateForm = (field: keyof FormData, value: string) => setForm((old) => ({ ...old, [field]: value }));
  const updateHandler = (id: number, field: keyof Omit<Handler, "id">, value: string) => setHandlers((old) => old.map((person) => person.id === id ? { ...person, [field]: value } : person));
  const addHandler = () => { setHandlers((old) => [...old, { id: nextHandlerId, name: "", role: "", cpf: "", email: "" }]); setNextHandlerId((old) => old + 1); };
  const toggle = (key: string) => setSelected((old) => { const next = { ...old }; if (key in next) delete next[key]; else next[key] = ""; return next; });

  function validateVariables() {
    if (!selectedItems.length) { setNoticeError(true); setNotice("Selecione ao menos uma variável para continuar."); setStep(1); return false; }
    if (missingJustifications.length) {
      const first = missingJustifications[0];
      setNoticeError(true);
      setNotice(`Justifique o uso de todas as variáveis selecionadas. Faltam ${missingJustifications.length} justificativas; confira ${first.variable} (${first.kind.toUpperCase()}).`);
      setKind(first.kind); setSearch(""); setStep(1); return false;
    }
    const missingCuts = [!form.territory.trim() && "recorte territorial", !form.period.trim() && "referência temporal", !form.audience.trim() && "recorte de público"].filter(Boolean);
    if (missingCuts.length) {
      setShowFormErrors(true); setNoticeError(true);
      setNotice(`Preencha os campos obrigatórios dos recortes da solicitação: ${missingCuts.join(", ")}.`);
      setStep(1); window.scrollTo({ top: 0, behavior: "smooth" }); return false;
    }
    return true;
  }

  async function generatePdf(type: "responsibility" | "confidentiality", person?: Handler) {
    if (!validateVariables()) return;
    if (!formReady) { setShowFormErrors(true); setNoticeError(true); setNotice("Há campos obrigatórios não preenchidos ou preenchidos incorretamente. Confira as mensagens em vermelho."); setStep(2); return; }
    const doc = new jsPDF({ unit: "mm", format: "a4" });
    try { doc.addImage(await loadPdfHeader(), "PNG", 10, 7, 190, 34.8); }
    catch { setNoticeError(true); setNotice("Não foi possível carregar o brasão para gerar o termo. Atualize a página e tente novamente."); return; }
    let y = 49;
    const title = type === "responsibility" ? "TERMO DE RESPONSABILIDADE - MDA" : "TERMO DE COMPROMISSO DE MANUTENÇÃO DE SIGILO - MDA";
    y = pdfLines(doc, title, y, { bold: true, center: true, size: 13 }); y = pdfLines(doc, `Processo nº ${form.process || "não informado"}`, y, { center: true }); y += 4;
    if (type === "responsibility") {
      y = pdfLines(doc, "Termo de Responsabilidade pela utilização de dados de identificação do Cadastro Nacional da Agricultura Familiar (CAF), regulamentado pelo Decreto nº 9.064, de 31 de maio de 2017.", y);
      y = pdfLines(doc, `O/A ${form.institution}, com sede estabelecida no/a ${form.address}, localizada(o) em ${form.city}, ${form.uf}, Brasil, doravante denominado(a) SIGNATÁRIO(A), neste ato representado(a) por ${form.representative}, ${form.role}, ${form.nationality}, CPF nº ${form.cpf}, em virtude do PROCESSO SEI nº ${form.process || "não informado"}, que trata do compartilhamento de informações do Cadastro Nacional da Agricultura Familiar - CAF, firma o presente TERMO DE RESPONSABILIDADE, que disciplina a utilização de dados de identificação do CAF mediante as cláusulas e condições descritas a seguir.`, y);
      y = pdfLines(doc, "CLÁUSULA PRIMEIRA – DO OBJETO", y, { bold: true }); y = pdfLines(doc, "O presente termo estabelece as regras que regulam a utilização dos dados de identificação do Cadastro Nacional da Agricultura Familiar (CAF), pelo(a) SIGNATÁRIO(A), sem prejuízo dos parâmetros legais vigentes.", y);
      y = pdfLines(doc, "CLÁUSULA SEGUNDA – DAS OBRIGAÇÕES DO SIGNATÁRIO", y, { bold: true }); y = pdfLines(doc, `O(A) SIGNATÁRIO(A) compromete-se a utilizar os dados de identificação do CAF exclusivamente para a finalidade de ${form.purpose}.`, y);
      y = pdfLines(doc, "O(A) SIGNATÁRIO(A) poderá permitir o acesso somente aos servidores e técnicos do órgão, mediante assinatura de Termo de Compromisso de Manutenção de Sigilo, que ficará sob guarda do signatário e será apresentado sempre que solicitado ao MDA.", y);
      ["Não disponibilizar ou ceder os dados a terceiros;", "Informar ao Ministério sobre a substituição do responsável por este Termo e pelos Termos de Compromisso de Manutenção de Sigilo;", "Utilizar medidas de segurança técnicas e administrativas aptas a proteger os dados pessoais de acessos não autorizados e de situações acidentais ou ilícitas;", "Adotar medidas para prevenir danos em virtude do tratamento dos dados;", "Eliminar os dados após a finalidade ser alcançada ou quando deixarem de ser necessários, ressalvadas as hipóteses do art. 16 da Lei nº 13.709/2018; e", "Comunicar ao Ministério a ocorrência de incidente de segurança que possa acarretar risco ou dano relevante às pessoas ou famílias inscritas no CAF."].forEach((line, i) => { y = pdfLines(doc, `${String.fromCharCode(97 + i)}) ${line}`, y); });
      y = pdfLines(doc, "CLÁUSULA TERCEIRA – DAS PENALIDADES", y, { bold: true }); y = pdfLines(doc, "O(A) SIGNATÁRIO(A), bem como os servidores e técnicos signatários do Termo de Compromisso de Manutenção de Sigilo, responderão civil e criminalmente pela utilização dos dados de identificação do CAF para fins diversos do previsto neste Termo e por quaisquer danos causados pela divulgação inadequada de informações.", y);
      y = pdfLines(doc, "CLÁUSULA QUARTA – DO PRAZO", y, { bold: true }); y = pdfLines(doc, "O tratamento deverá ocorrer por até 12 (doze) meses a contar da disponibilização. Após esse prazo, o tratamento deverá ser encerrado e os dados eliminados pelo demandante, conforme os artigos 15 e 16 da Lei nº 13.709/2018. Havendo necessidade de prorrogação, o(a) SIGNATÁRIO(A) deverá oficializar a solicitação ao Ministério em até 30 (trinta) dias antes do encerramento do prazo.", y);
    } else {
      if (!person) return;
      y = pdfLines(doc, `Eu, ${person.name}, ${person.role}, CPF nº ${person.cpf}, declaro estar ciente da habilitação que me foi conferida para tratamento de dados de identificação do Cadastro Nacional da Agricultura Familiar (CAF) no âmbito do Termo de Responsabilidade assinado pelo(a) ${form.institution}.`, y); y = pdfLines(doc, "No que se refere às atribuições a mim conferidas, comprometo-me a:", y);
      [`Manusear as bases de dados apenas por necessidade de serviço e para fins exclusivamente relacionados a ${form.purpose};`, "Proteger os dados de acessos não autorizados, garantindo que sejam manipulados em ambiente seguro e controlado;", "Manter absoluta cautela na exibição de dados em tela, impressora ou meios eletrônicos;", "Não me ausentar do terminal sem encerrar a sessão de uso das bases;", "Manter sigilo dos dados ou informações obtidos por força de minhas atribuições, abstendo-me de revelá-los a terceiros;", "Eliminar os dados após a finalidade ser alcançada ou quando deixarem de ser necessários; e", "Observar as disposições e normas da Política de Controle de Acesso aos dados do CAF."].forEach((line, i) => { y = pdfLines(doc, `${String.fromCharCode(97 + i)}) ${line}`, y); });
    }
    y += 5; y = pdfLines(doc, "E, por estar de pleno acordo, firma o presente Termo.", y); y = pdfLines(doc, `${form.place}, ${formatDate(form.date)}.`, y); y += 9;
    y = pdfLines(doc, "(assinado eletronicamente)", y, { center: true }); y = pdfLines(doc, "______________________________________________", y, { center: true }); y = pdfLines(doc, type === "responsibility" ? form.representative : person!.name, y, { center: true, bold: true }); pdfLines(doc, `CPF nº ${type === "responsibility" ? form.cpf : person!.cpf}`, y, { center: true });
    doc.save(type === "responsibility" ? "Termo_de_Responsabilidade_CAF.pdf" : `Termo_de_Compromisso_de_Sigilo_CAF_${digits(person!.cpf)}.pdf`); setNoticeError(false); setNotice("PDF gerado. Guarde o arquivo para assinatura eletrônica no gov.br.");
  }

  async function generateVariablesSpreadsheet() {
    if (!validateVariables()) return;
    try {
      const response = await fetch(`${import.meta.env.BASE_URL}templates/SOLICITACAO_DE_DADOS_MODELO.xlsx`);
      if (!response.ok) throw new Error("Modelo indisponível");
      const templateBuffer = await response.arrayBuffer();
      const file = await buildExcelFromTemplate(templateBuffer, form, handlers, selected);
      downloadBlob(file, `Solicitacao_de_Dados_CAF_${modalities.map((modality) => modality.toUpperCase()).join("_")}.xlsx`);
      setNoticeError(false); setNotice("Planilha Excel gerada com as três abas, logotipo CICAF, formatação do modelo e indicação SIM/NÃO para todas as variáveis.");
    } catch {
      setNoticeError(true); setNotice("Não foi possível gerar a planilha. Atualize a página e tente novamente.");
    }
  }

  function openEmail() {
    if (!validateVariables()) return;
    if (!formReady) { setShowFormErrors(true); setNoticeError(true); setNotice("Há campos obrigatórios não preenchidos ou preenchidos incorretamente. Confira as mensagens em vermelho."); setStep(2); return; }
    const body = `Prezados,\n\nEncaminho a solicitação de dados do CAF, o termo de responsabilidade assinado, os ${handlers.length} termos de compromisso e manutenção de sigilo assinados e a relação das ${selectedItems.length} variáveis selecionadas.\n\nInstituição: ${form.institution}\nResponsável: ${form.representative}\nTratadores: ${handlers.map((person) => person.name).join(", ")}\nProcesso SEI: ${form.process || "não informado"}\n\nAtenciosamente,\n${form.representative}`;
    window.location.href = `mailto:${destinationEmail}?subject=${encodeURIComponent(`Solicitação de Compartilhamento de dados da/o ${form.institution.trim()}`)}&body=${encodeURIComponent(body)}`;
  }

  return <main className="min-h-screen bg-[var(--page)] text-[var(--ink)]">
    <header className="border-b border-white/10 bg-[#073a31] text-white"><div className="mx-auto flex max-w-[1480px] items-center justify-between gap-5 px-5 py-4 lg:px-10"><div className="flex items-center gap-3"><div className="grid h-11 w-11 place-items-center rounded-xl bg-[#f3c845] text-[#073a31]"><FileCheck2 size={24}/></div><div><p className="text-xs font-semibold uppercase tracking-[.18em] text-emerald-100">Cadastro Nacional da Agricultura Familiar</p><h1 className="text-lg font-bold">Solicitação de Dados do CAF</h1></div></div><div className="hidden items-center gap-2 text-sm text-emerald-50 md:flex"><ShieldCheck size={18}/> Dados preenchidos permanecem no seu navegador</div></div></header>
    <div className="mx-auto max-w-[1480px] px-4 py-5 lg:px-10 lg:py-8">
      <div className="mb-6 grid gap-3 rounded-2xl border border-[#d8dfdc] bg-white p-3 shadow-sm md:grid-cols-4">{steps.map((label, index) => <button key={label} onClick={() => { if (index > Math.max(step, 1)) return; if (index >= 2 && !validateVariables()) return; setStep(index); }} className={`step ${index === step ? "active" : ""} ${index < step ? "done" : ""}`}><span>{index < step ? <Check size={16}/> : index + 1}</span><div><small>Etapa {index + 1}</small><strong>{label}</strong></div></button>)}</div>
      <div className="mb-5 flex flex-wrap items-center justify-between gap-3"><div><p className="eyebrow">Etapa {step + 1} de 4</p><h2 className="text-2xl font-bold tracking-tight">{steps[step]}</h2></div><div className="w-full max-w-xs"><Progress value={(step + 1) * 25} className="h-2"/></div></div>
      {notice && <div role={noticeError ? "alert" : "status"} className={`mb-5 flex items-center justify-between rounded-xl border px-4 py-3 text-sm font-medium ${noticeError ? "border-red-700 bg-red-600 text-white" : "border-[#b9d9cd] bg-[#edf8f4] text-[#07523f]"}`}><span className="flex items-center gap-2">{noticeError ? <ShieldCheck size={18}/> : <CheckCircle2 size={18}/>} {notice}</span><button aria-label="Fechar aviso" onClick={() => setNotice("")}><X size={18}/></button></div>}

      {step <= 1 && <>
        <Tabs value={kind} onValueChange={(value) => changeKind(value as Kind)} className="mb-4">
          <TabsList className="caf-tabs-list grid w-full grid-cols-2 rounded-xl border border-[#d8dfdc] bg-white p-1.5 shadow-sm">
            <TabsTrigger value="pf" className="caf-tab rounded-lg px-4 text-base font-bold">
              Pessoa Física <span className="rounded-full bg-black/5 px-2 py-0.5 text-xs data-[state=active]:bg-white/15">PF</span>
            </TabsTrigger>
            <TabsTrigger value="pj" className="caf-tab rounded-lg px-4 text-base font-bold">
              Pessoa Jurídica <span className="rounded-full bg-black/5 px-2 py-0.5 text-xs data-[state=active]:bg-white/15">PJ</span>
            </TabsTrigger>
          </TabsList>
        </Tabs>
        <div className="mb-4 rounded-xl border border-[#d8dfdc] bg-white p-4 shadow-sm"><p className="mb-3 text-sm font-bold text-[#073a31]">Classificação dos dados conforme a LGPD</p><div className="grid gap-2 sm:grid-cols-2 xl:grid-cols-4">{(Object.keys(classificationMeta) as DataClassification[]).map((classification) => { const meta = classificationMeta[classification]; return <div key={classification} className="flex items-start gap-2"><span className={`mt-0.5 shrink-0 rounded-full border px-2.5 py-1 text-xs font-bold ${meta.className}`}>{meta.label}</span><span className="text-xs leading-5 text-slate-600">{meta.description}</span></div>; })}</div><p className="mt-3 text-xs leading-5 text-slate-500">Classificação de apoio. A análise definitiva deve considerar a finalidade, a combinação entre variáveis e o risco de reidentificação.</p></div>
        {step === 1 && <section className="panel mb-4 p-5 md:p-6"><div className="mb-4"><h3 className="font-bold">Recortes da solicitação</h3><p className="mt-1 text-sm text-slate-500">Informe os recortes que devem ser considerados na extração das variáveis selecionadas. Os campos com * são obrigatórios.</p></div><div className="form-grid"><TextField label="Recorte territorial *" value={form.territory} onChange={(v) => updateForm("territory", v)} placeholder="Ex.: Estado de Goiás; municípios específicos..." error={showFormErrors && !form.territory.trim() ? "Campo obrigatório." : undefined}/><TextField label="Referência temporal *" value={form.period} onChange={(v) => updateForm("period", v)} placeholder="Ex.: janeiro de 2024 a dezembro de 2026" error={showFormErrors && !form.period.trim() ? "Campo obrigatório." : undefined}/><TextField label="Recorte de público *" value={form.audience} onChange={(v) => updateForm("audience", v)} placeholder="Ex.: famílias indígenas, faixa de renda ou idade" error={showFormErrors && !form.audience.trim() ? "Campo obrigatório." : undefined} wide/></div></section>}
        <section className="workspace-grid">
        <aside className="panel h-fit p-5 lg:sticky lg:top-5"><div><Label htmlFor="search">Buscar no dicionário</Label><div className="relative mt-2"><Search className="absolute left-3 top-2.5 text-slate-400" size={18}/><Input id="search" className="pl-10" placeholder="Variável, descrição ou valor" value={search} onChange={(e) => setSearch(e.target.value)}/></div></div><div className="mt-6 rounded-xl bg-[#f4f7f6] p-4"><div className="flex items-end justify-between"><div><p className="text-sm text-slate-500">{step === 0 ? "Modo de consulta" : "Selecionadas"}</p><p className={`${step === 0 ? "text-lg" : "text-3xl"} font-bold text-[#07523f]`}>{step === 0 ? "Somente leitura" : selectedItems.length}</p></div><BookOpen className="text-[#d0a819]"/></div><p className="mt-2 text-xs leading-5 text-slate-500">{step === 0 ? "Consulte nomes, descrições e valores possíveis. A seleção será feita na próxima etapa." : "Cada variável deve ter uma justificativa ligada à finalidade informada."}</p></div><div className="mt-5 space-y-2">{sections.map((section) => <a key={section.name} href={`#${section.name.replace(/\W/g, "-")}`} className="flex justify-between rounded-lg px-3 py-2 text-sm hover:bg-[#f1f5f3]"><span>{section.name}</span><span className="text-slate-400">{section.items.length}</span></a>)}</div></aside>
        <div className="panel min-w-0 overflow-hidden"><div className="border-b border-[#dfe5e2] px-5 py-4"><h3 className="font-bold">{step === 0 ? `Dicionário de dados — ${kind.toUpperCase()}` : "Selecione e justifique as variáveis"}</h3><p className="mt-1 text-sm text-slate-500">{shown.length} variáveis encontradas. {step === 0 ? "Consulte as informações; nenhuma seleção é realizada nesta etapa." : "Todas as variáveis do dicionário estão disponíveis para solicitação."}</p></div><div className="max-h-[68vh] overflow-auto">{sections.map((section) => { const items = shown.filter((item) => item.section === section.name); if (!items.length) return null; return <div key={section.name} id={section.name.replace(/\W/g, "-")} className="scroll-mt-6"><div className="sticky top-0 z-10 border-y border-[#dfe5e2] bg-[#edf3f0] px-5 py-2 text-xs font-bold uppercase tracking-[.12em] text-[#07523f]">{section.name}</div>{items.map((item) => { const classification = classifyVariable(item.variable, item.kind); const meta = classificationMeta[classification]; return <article key={item.key} className={`variable-row ${step === 0 ? "readonly" : ""} ${step === 1 && selected[item.key] !== undefined ? "selected" : ""}`}>{step === 1 && <div className="pt-1"><Checkbox checked={selected[item.key] !== undefined} onCheckedChange={() => toggle(item.key)} aria-label={`Selecionar ${item.variable}`}/></div>}<div className="min-w-0"><div className="flex min-w-max items-center gap-3"><code className="shrink-0">{item.variable}</code><p className="whitespace-nowrap text-sm leading-6 text-slate-700">{item.description}</p><span className={`shrink-0 rounded-full border px-2.5 py-1 text-xs font-bold ${meta.className}`} title={meta.description}>{meta.short}</span></div>{item.values.length > 0 && <details className="mt-2"><summary className="cursor-pointer text-xs font-semibold text-[#0b6b54]">Ver valores possíveis ({item.values.length})</summary><div className="mt-2 flex flex-wrap gap-1.5">{item.values.map((value) => <span key={value} className="value-chip">{value}</span>)}</div></details>}{step === 1 && selected[item.key] !== undefined && <div><Label className="mt-3 block" htmlFor={`justification-${item.key}`}>Justificativa do uso *</Label><Textarea id={`justification-${item.key}`} required aria-invalid={!selected[item.key]?.trim()} aria-label={`Justificativa para ${item.variable}`} className={`mt-2 min-h-20 ${!selected[item.key]?.trim() ? "border-red-600 bg-red-50" : ""}`} placeholder="Justifique por que esta variável é necessária..." value={selected[item.key]} onChange={(e) => setSelected((old) => ({ ...old, [item.key]: e.target.value }))}/>{!selected[item.key]?.trim() && <p className="mt-2 rounded-md bg-red-600 px-3 py-2 text-sm font-medium text-white">Campo obrigatório: informe a justificativa para continuar.</p>}</div>}</div></article>; })}</div>; })}</div></div>
        <aside className="panel h-fit p-5 lg:sticky lg:top-5">{step === 0 ? <><p className="eyebrow">Consulta ao dicionário</p><h3 className="mt-1 text-lg font-bold">Conheça os dados disponíveis</h3><p className="mt-3 text-sm leading-6 text-slate-500">Use a busca, alterne entre PF e PJ e consulte os valores possíveis. Na Etapa 2 você poderá escolher as variáveis necessárias.</p></> : <><p className="eyebrow">Resumo da solicitação</p><h3 className="mt-1 text-lg font-bold">{selectedItems.length ? `${selectedItems.length} variáveis` : "Nenhuma variável"}</h3><div className="mt-4 max-h-72 space-y-2 overflow-auto">{selectedItems.map((item) => <div key={item.key} className="rounded-lg border border-[#dfe5e2] p-3 text-sm"><div className="flex items-start justify-between gap-2"><code>{item.variable}</code><button onClick={() => toggle(item.key)} aria-label="Remover variável"><X size={15}/></button></div><p className="mt-1 text-xs text-slate-500">{item.kind.toUpperCase()} · {item.section}{!selected[item.key]?.trim() && " — justificativa pendente"}</p></div>)}</div><div className="mt-5 rounded-xl border border-[#ecdca6] bg-[#fff9e8] p-4 text-xs leading-5 text-[#725d18]">Selecione somente os dados estritamente necessários, conforme o princípio da necessidade da LGPD.</div></>}</aside>
        </section>
      </>}

      {step === 2 && <section className="panel p-5 md:p-8"><div className="mb-7 max-w-3xl"><h3 className="text-xl font-bold">Dados da solicitação</h3><p className="mt-2 text-sm leading-6 text-slate-500">Preencha os três grupos abaixo. Os campos com * são obrigatórios.</p></div>
        <div className="space-y-8">
          <div><h4 className="mb-4 border-b pb-2 text-lg font-bold">1. Dados da instituição solicitante</h4><div className="form-grid"><Field label="Instituição ou órgão *" value={form.institution} onChange={(v) => updateForm("institution", v)} error={showFormErrors && !form.institution.trim() ? "Campo obrigatório." : undefined} wide/><Field label="Endereço completo *" value={form.address} onChange={(v) => updateForm("address", v)} error={showFormErrors && !form.address.trim() ? "Campo obrigatório." : undefined} wide/><SelectField label="UF *" value={form.uf} onChange={(v) => setForm((old) => ({ ...old, uf: v, city: "", place: "" }))} options={Object.keys(municipalities)} error={showFormErrors && !form.uf ? "Campo obrigatório." : undefined}/><SelectField label="Município *" value={form.city} onChange={(v) => setForm((old) => ({ ...old, city: v, place: v ? `${v}/${old.uf}` : "" }))} options={form.uf ? municipalities[form.uf] || [] : []} disabled={!form.uf} placeholder={form.uf ? "Selecione o município" : "Selecione primeiro a UF"} error={showFormErrors && !form.city ? "Campo obrigatório." : undefined}/><Field label="Processo SEI" value={form.process} onChange={(v) => updateForm("process", formatSeiProcess(v))} placeholder="00000.000000/0000-00" inputMode="numeric" maxLength={20}/><Field label="Local de assinatura" value={form.place} onChange={() => {}} placeholder="Preenchido após selecionar Município e UF" readOnly/><Field label="Data do preenchimento" value={formatDate(form.date)} onChange={() => {}} readOnly/><TextField label="Finalidade do uso dos dados *" value={form.purpose} onChange={(v) => updateForm("purpose", v)} error={showFormErrors && !form.purpose.trim() ? "Campo obrigatório." : undefined} placeholder="Descreva de forma específica o estudo, política pública ou atividade..." wide/></div><p className="mt-3 text-xs leading-5 text-slate-500">O local é preenchido automaticamente com o Município/UF escolhidos, e a data corresponde ao dia do preenchimento do formulário.</p></div>
          <div><h4 className="mb-2 border-b pb-2 text-lg font-bold">2. Responsável pela instituição e pelos dados</h4><p className="mb-4 text-sm text-slate-500">Esta pessoa assinará somente o Termo de Responsabilidade.</p><div className="form-grid"><Field label="Nome do responsável *" value={form.representative} onChange={(v) => updateForm("representative", v)} error={showFormErrors && !form.representative.trim() ? "Campo obrigatório." : undefined}/><Field label="Cargo ou função *" value={form.role} onChange={(v) => updateForm("role", v)} error={showFormErrors && !form.role.trim() ? "Campo obrigatório." : undefined}/><Field label="Nacionalidade *" value={form.nationality} onChange={(v) => updateForm("nationality", v)} error={showFormErrors && !form.nationality.trim() ? "Campo obrigatório." : undefined}/><Field label="CPF do responsável *" value={form.cpf} onChange={(v) => updateForm("cpf", formatCpf(v))} error={showFormErrors && !form.cpf ? "Campo obrigatório." : !cpfOkay ? "CPF preenchido incorretamente." : undefined}/><Field label="E-mail do responsável *" value={form.email} onChange={(v) => updateForm("email", v)} error={showFormErrors && !form.email.trim() ? "Campo obrigatório." : !emailOkay ? "E-mail preenchido incorretamente." : undefined} type="email"/><Field label="Telefone" value={form.phone} onChange={(v) => updateForm("phone", formatPhone(v))} type="tel" placeholder="(61) 99999-9999"/></div></div>
          <div><h4 className="mb-2 border-b pb-2 text-lg font-bold">3. Pessoas que vão tratar os dados</h4><p className="mb-4 text-sm text-slate-500">Será gerado um Termo de Compromisso e Manutenção de Sigilo para cada pessoa.</p><div className="space-y-4">{handlers.map((person, index) => <div key={person.id} className="rounded-xl border border-[#d9e1de] bg-[#f8faf9] p-4"><div className="mb-4 flex items-center justify-between"><h5 className="font-bold">Tratador(a) {index + 1}</h5>{handlers.length > 1 && <Button variant="outline" size="sm" onClick={() => setHandlers((old) => old.filter((item) => item.id !== person.id))}>Remover</Button>}</div><div className="form-grid"><Field label={`Nome do tratador ${index + 1} *`} value={person.name} onChange={(v) => updateHandler(person.id, "name", v)} error={showFormErrors && !person.name.trim() ? "Campo obrigatório." : undefined}/><Field label={`Cargo ou função do tratador ${index + 1} *`} value={person.role} onChange={(v) => updateHandler(person.id, "role", v)} error={showFormErrors && !person.role.trim() ? "Campo obrigatório." : undefined}/><Field label={`CPF do tratador ${index + 1} *`} value={person.cpf} onChange={(v) => updateHandler(person.id, "cpf", formatCpf(v))} error={showFormErrors && !person.cpf ? "Campo obrigatório." : person.cpf && !validateCpf(person.cpf) ? "CPF preenchido incorretamente." : undefined}/><Field label={`E-mail do tratador ${index + 1} *`} value={person.email} onChange={(v) => updateHandler(person.id, "email", v)} error={showFormErrors && !person.email.trim() ? "Campo obrigatório." : person.email && !validEmail(person.email) ? "E-mail preenchido incorretamente." : undefined} type="email"/></div></div>)}</div><Button variant="outline" className="mt-4" onClick={addHandler}>Adicionar outra pessoa</Button></div>
        </div></section>}

      {step === 3 && <section className="grid gap-5 lg:grid-cols-[1.25fr_.75fr]"><div className="panel p-5 md:p-8"><p className="eyebrow">Documentos prontos</p><h3 className="mt-1 text-xl font-bold">Baixe, assine e encaminhe</h3><div className="mt-6 grid gap-4 md:grid-cols-2"><DocumentCard title="Termo de Responsabilidade" description={`Para ${form.representative}, responsável pela instituição.`} onDownload={() => generatePdf("responsibility")}/>{handlers.map((person) => <DocumentCard key={person.id} title="Termo de Compromisso e Manutenção de Sigilo" description={`Para ${person.name} — CPF ${person.cpf}.`} onDownload={() => generatePdf("confidentiality", person)}/>) }<SpreadsheetCard onDownload={generateVariablesSpreadsheet}/></div><div className="mt-6 flex flex-col gap-4 rounded-xl border border-[#c9d8d2] bg-[#f2f7f5] p-4 text-sm leading-6 text-slate-600 sm:flex-row sm:items-center sm:justify-between"><p><strong className="text-[#07523f]">Assinatura pelo gov.br:</strong> o responsável assina o Termo de Responsabilidade. Cada tratador assina seu próprio Termo de Compromisso e Manutenção de Sigilo. Anexe todos os PDFs assinados e a planilha Excel ao e-mail.</p><a href="https://assinador.iti.br/" target="_blank" rel="noopener noreferrer" className="inline-flex min-h-10 shrink-0 items-center justify-center gap-2 rounded-lg bg-[#07523f] px-4 py-2 font-semibold text-white shadow-sm transition-colors hover:bg-[#063f32] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#16866b] focus-visible:ring-offset-2">Acessar assinatura GOV.BR <ExternalLink size={16}/></a></div></div><aside className="panel h-fit p-5 md:p-7"><div className="grid h-11 w-11 place-items-center rounded-xl bg-[#e5f3ee] text-[#07523f]"><Mail size={22}/></div><h3 className="mt-4 text-lg font-bold">Encaminhar por e-mail</h3><p className="mt-2 text-sm leading-6 text-slate-500">O botão abre seu aplicativo de e-mail com destinatário, assunto e mensagem preenchidos. Anexe os termos assinados e a planilha.</p><div className="mt-5"><Label htmlFor="destination">E-mail da Coordenação do CAF</Label><Input id="destination" className="mt-2" type="email" value={destinationEmail} readOnly aria-readonly="true"/></div><Button className="mt-4 w-full" onClick={openEmail}><Mail size={17}/> Preparar e-mail</Button><div className="mt-6 border-t border-[#dfe5e2] pt-5 text-sm"><div className="flex justify-between"><span className="text-slate-500">Modalidade</span><strong>{modalityLabel}</strong></div><div className="mt-2 flex justify-between"><span className="text-slate-500">Tratadores</span><strong>{handlers.length}</strong></div><div className="mt-2 flex justify-between"><span className="text-slate-500">Variáveis</span><strong>{selectedItems.length}</strong></div></div></aside></section>}

      <footer className="mt-6 flex flex-col-reverse items-stretch justify-between gap-3 sm:flex-row sm:items-center"><Button variant="outline" disabled={step === 0} onClick={() => setStep((old) => Math.max(0, old - 1))}><ChevronLeft size={17}/> Voltar</Button>{step < 3 && <Button onClick={() => { if (step >= 1 && !validateVariables()) return; if (step === 2 && !formReady) { setShowFormErrors(true); setNoticeError(true); setNotice("Há campos obrigatórios não preenchidos ou preenchidos incorretamente. Confira as mensagens em vermelho."); window.scrollTo({ top: 0, behavior: "smooth" }); return; } setShowFormErrors(false); setNoticeError(false); setNotice(""); setStep((old) => old + 1); }}>{step === 0 ? "Selecionar variáveis" : step === 1 ? "Preencher dados" : "Gerar documentos"}<ChevronRight size={17}/></Button>}</footer>
    </div>
  </main>;
}

function Field({ label, value, onChange, wide, error, type = "text", placeholder, readOnly = false, inputMode, maxLength }: { label: string; value: string; onChange: (v: string) => void; wide?: boolean; error?: string; type?: string; placeholder?: string; readOnly?: boolean; inputMode?: "numeric" | "text" | "email" | "tel" | "decimal" | "search" | "url" | "none"; maxLength?: number }) { const id = label.replace(/\W/g, "-"); return <div className={wide ? "md:col-span-2" : ""}><Label htmlFor={id}>{label}</Label><Input id={id} aria-invalid={Boolean(error)} aria-readonly={readOnly} readOnly={readOnly} inputMode={inputMode} maxLength={maxLength} className={`mt-2 ${readOnly ? "cursor-not-allowed bg-slate-100 text-slate-700" : ""} ${error ? "border-red-600 bg-red-50" : ""}`} type={type} placeholder={placeholder} value={value} onChange={(e) => onChange(e.target.value)}/>{error && <p role="alert" className="mt-2 rounded-md bg-red-600 px-3 py-2 text-xs font-semibold text-white">{error}</p>}</div>; }
function SelectField({ label, value, onChange, options, error, disabled = false, placeholder = "Selecione" }: { label: string; value: string; onChange: (v: string) => void; options: string[]; error?: string; disabled?: boolean; placeholder?: string }) { return <div><Label>{label}</Label><select aria-invalid={Boolean(error)} disabled={disabled} className={`mt-2 h-9 w-full rounded-md border px-3 text-sm disabled:cursor-not-allowed disabled:bg-slate-100 disabled:text-slate-500 ${error ? "border-red-600 bg-red-50" : "border-input bg-transparent"}`} value={value} onChange={(e) => onChange(e.target.value)}><option value="">{placeholder}</option>{options.map((option) => <option key={option}>{option}</option>)}</select>{error && <p role="alert" className="mt-2 rounded-md bg-red-600 px-3 py-2 text-xs font-semibold text-white">{error}</p>}</div>; }
function TextField({ label, value, onChange, placeholder, wide, error }: { label: string; value: string; onChange: (v: string) => void; placeholder?: string; wide?: boolean; error?: string }) { return <div className={wide ? "md:col-span-2" : ""}><Label>{label}</Label><Textarea aria-invalid={Boolean(error)} className={`mt-2 min-h-28 ${error ? "border-red-600 bg-red-50" : ""}`} placeholder={placeholder} value={value} onChange={(e) => onChange(e.target.value)}/>{error && <p role="alert" className="mt-2 rounded-md bg-red-600 px-3 py-2 text-xs font-semibold text-white">{error}</p>}</div>; }
function DocumentCard({ title, description, onDownload }: { title: string; description: string; onDownload: () => void }) { return <article className="rounded-2xl border border-[#d9e1de] bg-white p-5"><div className="flex items-start justify-between"><div className="grid h-11 w-11 place-items-center rounded-xl bg-[#f4ead0] text-[#8c6d0b]"><FileText size={22}/></div><span className="rounded-full bg-[#e7f5ef] px-2.5 py-1 text-xs font-semibold text-[#07523f]">PDF</span></div><h4 className="mt-5 font-bold">{title}</h4><p className="mt-2 min-h-12 text-sm leading-6 text-slate-500">{description}</p><Button variant="outline" className="mt-4 w-full" onClick={onDownload}><Download size={17}/> Baixar PDF</Button></article>; }
function SpreadsheetCard({ onDownload }: { onDownload: () => void }) { return <article className="rounded-2xl border border-[#d9e1de] bg-white p-5"><div className="flex items-start justify-between"><div className="grid h-11 w-11 place-items-center rounded-xl bg-[#e5f3ee] text-[#07523f]"><FileSpreadsheet size={22}/></div><span className="rounded-full bg-[#e7f5ef] px-2.5 py-1 text-xs font-semibold text-[#07523f]">XLSX</span></div><h4 className="mt-5 font-bold">Planilha de solicitação</h4><p className="mt-2 min-h-12 text-sm leading-6 text-slate-500">No padrão do CAF, com os dados da solicitação, tratadores, variáveis escolhidas e justificativas.</p><Button variant="outline" className="mt-4 w-full" onClick={onDownload}><Download size={17}/> Baixar Excel</Button></article>; }
