import fs from 'node:fs/promises';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const outDir = 'D:/项目二和三/100kw-commercial-pv-system-design/calculations';
const workbook = Workbook.create();
const inputs = workbook.worksheets.add('Inputs');
const products = workbook.worksheets.add('Products');
const strings = workbook.worksheets.add('String_Check');
const schemes = workbook.worksheets.add('Schemes');
for (const s of [inputs, products, strings, schemes]) s.showGridLines = false;

const navy = '#17324D';
const blue = '#D9EAF7';
const green = '#E2F0D9';
const orange = '#FCE4D6';
const grey = '#F2F4F7';
const border = '#C9D2DC';
const title = (sheet, text, range) => {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[text]];
  sheet.getRange(range).format = { fill: navy, font: { bold: true, color: '#FFFFFF', size: 14 }, horizontalAlignment: 'left', verticalAlignment: 'center' };
};
const header = (sheet, range) => {
  sheet.getRange(range).format = { fill: blue, font: { bold: true, color: navy }, borders: { preset: 'all', style: 'thin', color: border }, wrapText: true, verticalAlignment: 'center' };
};
const body = (sheet, range) => {
  sheet.getRange(range).format = { borders: { preset: 'insideHorizontal', style: 'thin', color: border }, verticalAlignment: 'center', wrapText: true };
};
const note = (sheet, range) => sheet.getRange(range).format = { fill: grey, font: { italic: true, color: '#44546A', size: 10 }, wrapText: true };

title(inputs, '100 kW 工商业屋顶光伏：输入与边界', 'A1:F1');
inputs.getRange('A3:B13').values = [
  ['输入项','取值'],['地点','青岛'],['屋顶可用面积 (m²)',800],['交流侧目标 (kW)',100],['并网电压 (V)',400],['最低电池片温度 (°C)',-10],['最高电池片温度 (°C)',70],['屋顶面积利用系数',1.25],['系统类型','并网、无储能'],['项目性质','个人虚拟方案分析'],['数据访问日期','2026-07-30']
];
header(inputs,'A3:B3'); body(inputs,'A4:B13');
inputs.getRange('B5:B8').format.numberFormat = [['#,##0.0'],['#,##0.0'],['#,##0'],['#,##0']];
inputs.getRange('D3:F9').values = [
  ['边界提醒','本表做初步选型，不替代施工图',''],
  ['包含','组件/逆变器参数、组串约束、容量比、面积初核',''],
  ['不包含','结构承载、线缆截面、保护整定、并网审批、报价',''],
  ['温度口径','电池片温度假设，不等于环境温度',''],
  ['推荐阅读','先看 String_Check，再看 Schemes',''],
  ['来源','见 Products 表 Source 列',''],
  ['SAM','未填虚构年发电量，按 sam_method.md 后续运行','']
];
header(inputs,'D3:F3'); body(inputs,'D4:F9'); note(inputs,'D4:F9');
inputs.getRange('A:A').format.columnWidth = 25; inputs.getRange('B:B').format.columnWidth = 22; inputs.getRange('D:D').format.columnWidth = 18; inputs.getRange('E:F').format.columnWidth = 38; inputs.freezePanes.freezeRows(3);

title(products, '组件与逆变器参数（公开资料摘录）', 'A1:N1');
products.getRange('A3:N7').values = [
  ['类别','厂家','型号','Pmax / Pac','Vmp / MPPT min','Voc / Vdc max','Imp / Imax','Isc / Isc max','beta Voc','beta Vmp','MPPT数','尺寸/备注','来源','访问日期'],
  ['组件','Jinko','JKM585N-72HL4-V',585,43.53,52.47,13.44,14.07,-0.0025,-0.0029,'','2278×1134×35 mm','https://www.jinkosolar.com/uploads/JKM565-585N-72HL4-%28V%29-F4-EN.pdf','2026-07-30'],
  ['组件','Trina','TSM-580DE19R.W',580,39.0,46.3,14.86,15.94,-0.0025,-0.0034,'','2384×1134×35 mm','https://static.trinasolar.com/sites/default/files/Datasheet_Vertex_DE19R.W_EU_2022A.pdf','2026-07-30'],
  ['逆变器','Sungrow','SG110CX',110,200,1100,26,40, '', '',9,'每路最多2串；MPPT 200–1000 V','https://en.sungrowpower.com/upload/documentFile/DS_SG110CX%20Datasheet_V14_EN.pdf.pdf','2026-07-30'],
  ['逆变器','Huawei','SUN2000-100KTL-M2',100,200,1100,30,40,'','','10','每路最多2路输入；MPPT 200–1000 V','https://solar.huawei.com/admin/asset/v1/pro/view/0d3b86d0da694d88b8a9871688cda48d.pdf','2026-07-30']
];
header(products,'A3:N3'); body(products,'A4:N7');
products.getRange('D4:H7').format.numberFormat = [['0.0','0.0','0.0','0.0','0.0'],['0.0','0.0','0.0','0.0','0.0'],['0.0','0.0','0.0','0.0','0.0'],['0.0','0.0','0.0','0.0','0.0']];
products.getRange('I4:J5').format.numberFormat = [['0.00%','0.00%'],['0.00%','0.00%']];
for (const [col,w] of [['A',10],['B',12],['C',24],['D',13],['E',15],['F',14],['G',14],['H',14],['I',12],['J',12],['K',10],['L',32],['M',48],['N',14]]) products.getRange(`${col}:${col}`).format.columnWidth=w;
products.freezePanes.freezeRows(3);

title(strings, '组串电压、电流与面积校核', 'A1:O1');
strings.getRange('A3:O7').values = [
  ['组件','逆变器','串联块数','并联串数/MPPT','低温 Voc/块','低温组串 Voc','高温 Vmp/块','高温组串 Vmp','工作电流/MPPT','短路电流/MPPT','直流容量(kWp)','容量比','所需面积(m²)','硬约束结果','说明'],
  ['Jinko 585','Sungrow SG110CX',19,1,null,null,null,null,null,null,null,null,null,'',''],
  ['Jinko 585','Sungrow SG110CX',18,1,null,null,null,null,null,null,null,null,null,'',''],
  ['Trina 580','Huawei 100KTL-M2',21,1,null,null,null,null,null,null,null,null,null,'',''],
  ['Trina 580','Huawei 100KTL-M2',20,1,null,null,null,null,null,null,null,null,'','']
];
header(strings,'A3:O3'); body(strings,'A4:O7');
strings.getRange('E4:O7').formulas = [
  ["='Products'!F4*(1+'Products'!I4*('Inputs'!B8-25))",'=C4*E4',"='Products'!E4*(1+'Products'!J4*('Inputs'!B9-25))",'=C4*G4','=D4*\'Products\'!G4','=D4*\'Products\'!H4','=C4*D4*\'Products\'!D4/1000','=K4/\'Inputs\'!B6','=C4*D4*2.58*\'Inputs\'!B10','=IF(AND(F4<=1100,H4>=200,H4<=1000,I4<=26,J4<=40),"通过","需复核")','="Jinko参数"'],
  ["='Products'!F4*(1+'Products'!I4*('Inputs'!B8-25))",'=C5*E5',"='Products'!E4*(1+'Products'!J4*('Inputs'!B9-25))",'=C5*G5','=D5*\'Products\'!G4','=D5*\'Products\'!H4','=C5*D5*\'Products\'!D4/1000','=K5/\'Inputs\'!B6','=C5*D5*2.58*\'Inputs\'!B10','=IF(AND(F5<=1100,H5>=200,H5<=1000,I5<=26,J5<=40),"通过","需复核")','="Jinko参数"'],
  ["='Products'!F5*(1+'Products'!I5*('Inputs'!B8-25))",'=C6*E6',"='Products'!E5*(1+'Products'!J5*('Inputs'!B9-25))",'=C6*G6','=D6*\'Products\'!G5','=D6*\'Products\'!H5','=C6*D6*\'Products\'!D5/1000','=K6/\'Inputs\'!B6','=C6*D6*2.58*\'Inputs\'!B10','=IF(AND(F6<=1100,H6>=200,H6<=1000,I6<=30,J6<=40),"通过","需复核")','="Trina参数"'],
  ["='Products'!F5*(1+'Products'!I5*('Inputs'!B8-25))",'=C7*E7',"='Products'!E5*(1+'Products'!J5*('Inputs'!B9-25))",'=C7*G7','=D7*\'Products\'!G5','=D7*\'Products\'!H5','=C7*D7*\'Products\'!D5/1000','=K7/\'Inputs\'!B6','=C7*D7*2.58*\'Inputs\'!B10','=IF(AND(F7<=1100,H7>=200,H7<=1000,I7<=30,J7<=40),"通过","需复核")','="Trina参数"']
];
strings.getRange('E4:M7').format.numberFormat = [['0.0','0.0','0.0','0.0','0.0','0.0','0.0','0.00','0.0'],['0.0','0.0','0.0','0.0','0.0','0.0','0.0','0.00','0.0'],['0.0','0.0','0.0','0.0','0.0','0.0','0.0','0.00','0.0'],['0.0','0.0','0.0','0.0','0.0','0.0','0.0','0.00','0.0']];
strings.getRange('N4:N7').conditionalFormats.add('cellIs',{operator:'equalTo',formula:'="通过"',format:{fill:green,font:{color:'#006100',bold:true}}});
strings.getRange('N4:N7').conditionalFormats.add('cellIs',{operator:'equalTo',formula:'="需复核"',format:{fill:orange,font:{color:'#9C0006',bold:true}}});
for (const [col,w] of [['A',14],['B',20],['C',12],['D',15],['E',13],['F',14],['G',13],['H',14],['I',15],['J',15],['K',14],['L',10],['M',14],['N',12],['O',20]]) strings.getRange(`${col}:${col}`).format.columnWidth=w;
strings.freezePanes.freezeRows(3);

title(schemes, '三套方案比较（组件数量按整串取整）', 'A1:J1');
schemes.getRange('A3:J7').values = [
  ['方案','组件型号','逆变器','串联块数','总组串数','总组件数','直流容量(kWp)','容量比','面积初核(m²)','定位'],
  ['A 保守','Jinko 585','Sungrow SG110CX',19,9,null,null,null,null,'约100 kWp，面积和直流侧压力较小'],
  ['B 平衡','Jinko 585','Sungrow SG110CX',19,10,null,null,null,null,'约111 kWp，推荐候选'],
  ['C 偏高','Trina 580','Huawei 100KTL-M2',20,11,null,null,null,null,'约128 kWp，观察削峰和屋顶压力'],
  ['说明','三套方案仅作方案级比较','','','','','','','','SAM 年发电量待实际运行']
];
header(schemes,'A3:J3'); body(schemes,'A4:J7');
schemes.getRange('F4:I6').formulas = [
  ["=D4*E4", "=F4*'Products'!D4/1000", "=G4/'Inputs'!B6", "=F4*2.58*'Inputs'!B10"],
  ["=D5*E5", "=F5*'Products'!D4/1000", "=G5/'Inputs'!B6", "=F5*2.58*'Inputs'!B10"],
  ["=D6*E6", "=F6*'Products'!D5/1000", "=G6/'Inputs'!B6", "=F6*2.58*'Inputs'!B10"]
];
schemes.getRange('G4:H6').format.numberFormat = [['0.0','0.00'],['0.0','0.00'],['0.0','0.00']]; schemes.getRange('I4:I6').format.numberFormat = [['0.0'],['0.0'],['0.0']];
schemes.getRange('A9:B12').values = [['图表数据','容量比'],['A 保守',null],['B 平衡',null],['C 偏高',null]];
schemes.getRange('B10:B12').formulas = [['=H4'],['=H5'],['=H6']];
header(schemes,'A9:B9'); body(schemes,'A10:B12');
const chart = schemes.charts.add('bar', schemes.getRange('A9:B12')); chart.title = '三套方案容量比'; chart.hasLegend = false; chart.setPosition('L3','T18');
for (const [col,w] of [['A',14],['B',20],['C',22],['D',12],['E',15],['F',12],['G',14],['H',10],['I',14],['J',26]]) schemes.getRange(`${col}:${col}`).format.columnWidth=w;
schemes.freezePanes.freezeRows(3);

await fs.mkdir(outDir, { recursive: true });
const errors = await workbook.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A',options:{useRegex:true,maxResults:50},summary:'formula scan'});
console.log(errors.ndjson);
const inspect = await workbook.inspect({kind:'table',sheetId:'String_Check',range:'A3:O7',include:'values,formulas',tableMaxRows:8,tableMaxCols:16,maxChars:6000});
console.log(inspect.ndjson);
for (const [name,range] of [['Inputs','A1:F13'],['Products','A1:N7'],['String_Check','A1:O7'],['Schemes','A1:T18']]) {
  const blob = await workbook.render({sheetName:name,range,scale:1,format:'png'});
  await fs.writeFile(`${outDir}/preview_${name}.png`, new Uint8Array(await blob.arrayBuffer()));
}
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outDir}/100kw_pv_sizing.xlsx`);
