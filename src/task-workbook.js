const WORKBOOK_FILE_NAME = "TASK_CONTROL_Master.xlsx";
const WORKBOOK_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const WORKBOOK_FORMAT_VERSION = "task-control-v1";
const WORKBOOK_TABLE_NAME = "TaskMasterTable";
const TEMPLATE_URL = new URL("./TASK_CONTROL_Master.xlsx", document.baseURI).href;

const HEADERS = [
  "TaskID", "ParentID", "RootID", "SortOrder", "分類", "タスク名", "状態", "優先度",
  "担当", "開始日", "期日", "進捗率", "次の一手", "阻害要因", "メモ", "リンクURL",
  "削除", "作成日時", "更新日時", "更新元", "版", "想定所要時間（h）",
  "次に触る日", "今日への配置", "作業分割", "最小作業単位（h）"
];

function requireJsZip() {
  if (!globalThis.JSZip) throw new Error("Excel同期部品を読み込めませんでした");
  return globalThis.JSZip;
}

function escapeXml(value = "") {
  return String(value).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&apos;"}[c]));
}

function columnName(index) {
  let result = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + ((n - 1) % 26)) + result;
  return result;
}

function cell(reference, value) {
  const text = String(value ?? "");
  const preserve = /^\s|\s$|\n/.test(text);
  return `<c r="${reference}" t="inlineStr"><is><t${preserve ? ' xml:space="preserve"' : ""}>${escapeXml(text)}</t></is></c>`;
}

function normalizeTask(task = {}) {
  const progress = Math.max(0, Math.min(100, Number(task.progress) || 0));
  return {
    id: String(task.id || "").trim(), parentId: String(task.parentId || "").trim(), rootId: String(task.rootId || task.id || "").trim(),
    sortOrder: Number(task.sortOrder) || 0, category: String(task.category || "その他"), title: String(task.title || ""),
    status: String(task.status || "未着手"), priority: String(task.priority || "🟠WARM"), owner: String(task.owner || ""),
    startDate: String(task.startDate || "").slice(0, 10), dueDate: String(task.dueDate || "").slice(0, 10), progress,
    expectedMinutes: Math.max(0, Math.min(600000, Math.round((Number(task.expectedMinutes) || 0) / 15) * 15)),
    nextTouchDate: String(task.nextTouchDate || "").slice(0, 10),
    todayPlacement: ["自動判定","必ず入れる","入れない"].includes(task.todayPlacement) ? task.todayPlacement : "自動判定",
    workSplit: ["分割可","分割不可"].includes(task.workSplit) ? task.workSplit : "分割不可",
    minimumBlockMinutes: Math.max(15, Math.min(600000, Math.round((Number(task.minimumBlockMinutes) || 30) / 15) * 15)),
    nextAction: String(task.nextAction || ""), blocker: String(task.blocker || ""), notes: String(task.notes || ""), linkUrl: String(task.linkUrl || ""),
    deleted: Boolean(task.deleted), createdAt: String(task.createdAt || ""), updatedAt: String(task.updatedAt || ""),
    updatedBy: String(task.updatedBy || ""), revision: Math.max(1, Number.parseInt(task.revision, 10) || 1)
  };
}

function sortedTasks(tasks) {
  return [...tasks].map(normalizeTask).filter(t => t.id).sort((a,b) => a.rootId.localeCompare(b.rootId) || a.sortOrder-b.sortOrder || a.createdAt.localeCompare(b.createdAt));
}

function rowValues(t) {
  return [t.id,t.parentId,t.rootId,t.sortOrder,t.category,t.title,t.status,t.priority,t.owner,t.startDate,t.dueDate,t.progress,
    t.nextAction,t.blocker,t.notes,t.linkUrl,t.deleted ? "TRUE" : "FALSE",t.createdAt,t.updatedAt,t.updatedBy,t.revision,t.expectedMinutes / 60,
    t.nextTouchDate,t.todayPlacement,t.workSplit,t.minimumBlockMinutes / 60];
}

function sheetXml(tasks, tableRelationshipId) {
  const records = sortedTasks(tasks);
  const rows = [`<row r="1" ht="28" customHeight="1">${HEADERS.map((h,i)=>cell(`${columnName(i)}1`,h)).join("")}</row>`];
  records.forEach((task,index) => {
    const row = index + 2;
    rows.push(`<row r="${row}" ht="34" customHeight="1">${rowValues(task).map((v,i)=>cell(`${columnName(i)}${row}`,v)).join("")}</row>`);
  });
  if (!records.length) rows.push(`<row r="2" ht="24" customHeight="1">${HEADERS.map((_,i)=>cell(`${columnName(i)}2`,"")).join("")}</row>`);
  const lastRow = Math.max(2, records.length + 1);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:Z${lastRow}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="3" width="24" customWidth="1"/><col min="4" max="4" width="10" customWidth="1"/><col min="5" max="5" width="14" customWidth="1"/><col min="6" max="6" width="36" customWidth="1"/><col min="7" max="9" width="14" customWidth="1"/><col min="10" max="12" width="12" customWidth="1"/><col min="13" max="16" width="34" customWidth="1"/><col min="17" max="17" width="8" customWidth="1"/><col min="18" max="20" width="24" customWidth="1"/><col min="21" max="21" width="8" customWidth="1"/><col min="22" max="22" width="18" customWidth="1"/><col min="23" max="23" width="14" customWidth="1"/><col min="24" max="25" width="16" customWidth="1"/><col min="26" max="26" width="18" customWidth="1"/></cols>
<sheetData>${rows.join("")}</sheetData><pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.2" footer="0.2"/><tableParts count="1"><tablePart r:id="${escapeXml(tableRelationshipId)}"/></tableParts></worksheet>`;
}

function tableXml(count) {
  const lastRow = Math.max(2, count + 1);
  const cols = HEADERS.map((h,i)=>`<tableColumn id="${i+1}" name="${escapeXml(h)}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="${WORKBOOK_TABLE_NAME}" displayName="${WORKBOOK_TABLE_NAME}" ref="A1:Z${lastRow}" headerRowCount="1" totalsRowCount="0" totalsRowShown="0"><autoFilter ref="A1:Z${lastRow}"/><tableColumns count="${HEADERS.length}">${cols}</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
}

function tableRelationshipId(xml) {
  const tag = String(xml || "").match(/<Relationship\b[^>]*\bType=["'][^"']*\/table["'][^>]*>/i)?.[0] || "";
  const id = tag.match(/\bId=["']([^"']+)["']/i)?.[1] || "";
  if (!id) throw new Error("Excelテンプレートのテーブル関連付けが見つかりません");
  return id;
}

export async function buildTaskWorkbook(tasks, templateOverride = null) {
  const response = templateOverride ? null : await fetch(TEMPLATE_URL);
  if (response && !response.ok) throw new Error(`Excelテンプレートを読み込めませんでした（HTTP ${response.status}）`);
  const bytes = templateOverride || await response.arrayBuffer();
  const zip = await requireJsZip().loadAsync(bytes);
  const relFile = zip.file("xl/worksheets/_rels/sheet1.xml.rels");
  if (!relFile) throw new Error("Excelテンプレートの構成が不正です");
  const relationshipId = tableRelationshipId(await relFile.async("text"));
  const records = sortedTasks(tasks);
  zip.file("xl/worksheets/sheet1.xml", sheetXml(records, relationshipId));
  zip.file("xl/tables/table1.xml", tableXml(records.length));
  return zip.generateAsync({type:"uint8array",compression:"DEFLATE",compressionOptions:{level:6}});
}

function parseXml(text,label) {
  const doc = new DOMParser().parseFromString(text,"application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error(`${label}のXMLを解析できません`);
  return doc;
}

function sharedStrings(doc) {
  return doc ? [...doc.getElementsByTagName("si")].map(n=>[...n.getElementsByTagName("t")].map(t=>t.textContent||"").join("")) : [];
}

function cellValue(c,strings) {
  const type = c.getAttribute("t") || "";
  if (type === "inlineStr") return [...c.getElementsByTagName("t")].map(n=>n.textContent||"").join("");
  const raw = c.getElementsByTagName("v")[0]?.textContent || "";
  return type === "s" ? (strings[Number.parseInt(raw,10)] || "") : raw;
}

export async function parseTaskWorkbook(bytes) {
  const zip = await requireJsZip().loadAsync(bytes);
  const sheetFile = zip.file("xl/worksheets/sheet1.xml");
  if (!sheetFile) throw new Error("TASKSシートが見つかりません");
  const doc = parseXml(await sheetFile.async("text"),"TASKS");
  const sharedFile = zip.file("xl/sharedStrings.xml");
  const strings = sharedStrings(sharedFile ? parseXml(await sharedFile.async("text"),"sharedStrings") : null);
  const cells = [...doc.getElementsByTagName("c")];
  const headerValue = reference => {
    const headerCell = cells.find(c => c.getAttribute("r") === reference);
    return headerCell ? cellValue(headerCell, strings) : "";
  };
  const expectedStoredAsHours = headerValue("V1") === "想定所要時間（h）";
  const minimumBlockStoredAsHours = headerValue("Z1") === "最小作業単位（h）";
  const tasks = [];
  [...doc.getElementsByTagName("row")].forEach(row => {
    if (Number(row.getAttribute("r") || 0) <= 1) return;
    const values = Array(HEADERS.length).fill("");
    [...row.getElementsByTagName("c")].forEach(c => {
      const match = (c.getAttribute("r") || "").match(/^([A-Z]+)/);
      if (!match) return;
      let index = 0;
      for (const ch of match[1]) index = index * 26 + ch.charCodeAt(0) - 64;
      if (index > 0 && index <= values.length) values[index-1] = cellValue(c,strings);
    });
    if (!values[0]) return;
    tasks.push(normalizeTask({id:values[0],parentId:values[1],rootId:values[2],sortOrder:values[3],category:values[4],title:values[5],status:values[6],priority:values[7],owner:values[8],startDate:values[9],dueDate:values[10],progress:values[11],nextAction:values[12],blocker:values[13],notes:values[14],linkUrl:values[15],deleted:/^(TRUE|1)$/i.test(values[16]),createdAt:values[17],updatedAt:values[18],updatedBy:values[19],revision:values[20],expectedMinutes:(Number(values[21])||0)*(expectedStoredAsHours?60:1),nextTouchDate:values[22],todayPlacement:values[23],workSplit:values[24],minimumBlockMinutes:(Number(values[25])||0)*(minimumBlockStoredAsHours?60:1)}));
  });
  const tableFile = zip.file("xl/tables/table1.xml");
  const tableText = tableFile ? await tableFile.async("text") : "";
  Object.defineProperty(tasks,"workbookFormatVersion",{value:new RegExp(`\\b(?:name|displayName)=[\"']${WORKBOOK_TABLE_NAME}[\"']`).test(tableText)?WORKBOOK_FORMAT_VERSION:"legacy",enumerable:false});
  return tasks;
}

export { WORKBOOK_FILE_NAME, WORKBOOK_MIME, WORKBOOK_FORMAT_VERSION };
