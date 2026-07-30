import fs from 'node:fs/promises';
import { Presentation, PresentationFile } from '@oai/artifact-tool';

const outDir = 'D:/项目二和三/100kw-commercial-pv-system-design/deliverables';
const p = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const C = { bg:'#F7F9FC', navy:'#17324D', teal:'#1F6F8B', pale:'#D9EAF7', text:'#263645', muted:'#5E7184', white:'#FFFFFF', green:'#E2F0D9', orange:'#FCE4D6' };
const writeBlob = async (path, blob) => fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
function box(slide, text, x, y, w, h, style={}) {
  const s = slide.shapes.add({geometry:'textbox', position:{left:x,top:y,width:w,height:h}, fill:style.fill??'none', line:{style:'solid',fill:style.line??'none',width:style.lineWidth??0}});
  s.text = text; s.text.style = {fontSize:style.fontSize??20, color:style.color??C.text, bold:style.bold??false, italic:style.italic??false, align:style.align??'left', valign:style.valign??'top'}; return s;
}
function rect(slide,x,y,w,h,fill,line='none') { return slide.shapes.add({geometry:'roundRect',position:{left:x,top:y,width:w,height:h},fill,line:{style:'solid',fill:line,width:line==='none'?0:1},borderRadius:'rounded-md'}); }
function base(title, kicker='PROJECT 3 / PERSONAL CASE') {
  const slide=p.slides.add(); slide.background.fill=C.bg; box(slide,kicker,64,38,360,24,{fontSize:14,bold:true,color:C.teal}); box(slide,title,64,76,1140,62,{fontSize:36,bold:true,color:C.navy}); return slide;
}
function notes(slide, text) { slide.speakerNotes.textFrame.setText(`[Sources]\n${text}`); slide.speakerNotes.setVisible(true); }

let s=p.slides.add(); s.background.fill=C.navy; box(s,'100 kW 工商业屋顶光伏',78,112,800,90,{fontSize:50,bold:true,color:C.white}); box(s,'方案设计与设备选型分析',82,220,700,52,{fontSize:30,color:'#B9D9E8'}); box(s,'个人虚拟案例｜青岛｜用于秋招展示',84,322,560,34,{fontSize:20,color:'#D6E4EE'}); rect(s,84,424,490,6,C.teal); box(s,'核心方法：需求 → 参数 → 约束 → 方案取舍',84,462,700,34,{fontSize:22,color:C.white}); notes(s,'产品参数：厂商数据手册，详见 references/source_register.md。');

s=base('先把问题说清楚：一个假设的工商业屋顶');
rect(s,64,180,350,340,C.white,'#D8E2EC'); box(s,'场景输入',92,210,200,30,{fontSize:22,bold:true,color:C.navy}); box(s,'青岛\n800 m² 可用屋顶\n380/400 V 三相并网\n交流侧目标 100 kW\n无储能',92,266,260,180,{fontSize:24,color:C.text});
rect(s,472,180,690,340,C.pale); box(s,'我没有把它包装成真实工程',510,214,580,34,{fontSize:22,bold:true,color:C.navy}); box(s,'包含：设备参数、组串匹配、容量比、面积初核\n不包含：结构、施工图、保护整定、并网审批、报价与收益承诺',510,282,560,138,{fontSize:25,color:C.text}); box(s,'这样项目重点放在“能解释的选型逻辑”上。',510,452,560,30,{fontSize:20,bold:true,color:C.teal}); notes(s,'场景与边界来自 requirements/project_requirements.md。');

s=base('设备选型：先看电气参数，不先看品牌');
box(s,'组件样本',70,170,240,30,{fontSize:22,bold:true,color:C.teal}); box(s,'Jinko 585 W\nVmp 43.53 V｜Voc 52.47 V\nImp 13.44 A｜Isc 14.07 A',70,216,470,160,{fontSize:26,color:C.text});
box(s,'逆变器样本',680,170,240,30,{fontSize:22,bold:true,color:C.teal}); box(s,'Sungrow SG110CX\n1100 V max｜200–1000 V MPPT\n9 MPPT｜26 A / MPPT',680,216,500,160,{fontSize:26,color:C.text});
rect(s,70,448,1110,100,C.white,'#D8E2EC'); box(s,'比较逻辑',96,468,150,28,{fontSize:20,bold:true,color:C.navy}); box(s,'工作电压更高、组串更容易落在 MPPT 区间；输入电流不做满，给配置和维护留余量。',278,468,860,40,{fontSize:22,color:C.text}); notes(s,'Jinko、Sungrow 参数来自厂商数据手册，见 references/source_register.md。');

s=base('组串校核：我最需要讲清楚的一页');
rect(s,64,174,1110,80,C.pale); box(s,'低温看 Voc 上限｜高温看 Vmp 下限｜并联看 MPPT 输入电流',92,198,1000,32,{fontSize:27,bold:true,color:C.navy});
box(s,'低温 Voc：52.47 × [1 + (-0.0025) × (-10 - 25)] = 57.1 V/块\n19 块/串：57.1 × 19 ≈ 1084 V\n\n高温 Vmp：43.53 × [1 + (-0.0029) × (70 - 25)] = 37.8 V/块\n19 块/串：37.8 × 19 ≈ 719 V',96,292,1050,238,{fontSize:24,color:C.text});
rect(s,64,560,1110,58,C.green); box(s,'结论：19 块/串通过本项目假设下的电压校核；每个 MPPT 先按 1 串接入，避免电流余量被用满。',88,578,1040,24,{fontSize:20,bold:true,color:'#21613A'}); notes(s,'计算公式与结果来自 calculations/100kw_pv_sizing.xlsx 的 String_Check 工作表。');

s=base('三套方案：不是“功率越大越好”');
const scheme=[['A 保守','1.00','100.0 kWp',C.pale],['B 平衡','1.11','111.2 kWp',C.green],['C 偏高','1.28','127.6 kWp',C.orange]]; let x=72; for(const [name,ratio,dc,fill] of scheme){rect(s,x,184,330,300,fill); box(s,name,x+28,216,220,30,{fontSize:25,bold:true,color:C.navy}); box(s,ratio,x+28,280,180,64,{fontSize:46,bold:true,color:C.teal}); box(s,'容量比',x+216,302,90,24,{fontSize:17,color:C.muted}); box(s,dc,x+28,390,220,32,{fontSize:24,bold:true,color:C.text}); box(s,name==='B 平衡'?'发电量、面积、复杂度更均衡':name==='A 保守'?'面积压力较小':'组件更多，需关注削峰和屋顶压力',x+28,450,270,42,{fontSize:17,color:C.text}); x+=370; }
box(s,'推荐候选：B。理由是综合取舍更平衡，不是因为每一项都最高。',72,548,1040,34,{fontSize:23,bold:true,color:C.navy}); notes(s,'方案容量比来自 calculations/100kw_pv_sizing.xlsx 的 Schemes 工作表；年发电量暂未填入，避免虚构 SAM 结果。');

s=base('交付与不足：把边界讲出来反而更可靠');
rect(s,70,174,520,360,C.white,'#D8E2EC'); box(s,'已交付',100,208,200,30,{fontSize:24,bold:true,color:C.teal}); box(s,'✓ 参数来源登记\n✓ 输入与公式计算表\n✓ 组串和面积初核\n✓ 三套方案比较\n✓ 技术方案书\n✓ SAM 后续运行方法',100,270,400,230,{fontSize:24,color:C.text});
rect(s,660,174,520,360,C.pale); box(s,'后续必须确认',690,208,240,30,{fontSize:24,bold:true,color:C.navy}); box(s,'结构承载\n真实排布与通道\n线缆与保护\n并网要求\n实际 SAM 导出\n商务报价',690,270,400,230,{fontSize:24,color:C.text});
box(s,'面试表达：我完成的是方案级分析，不冒充做过最终工程设计。',72,584,1100,34,{fontSize:24,bold:true,color:C.teal}); notes(s,'项目边界与待确认项来自 requirements/project_requirements.md 和 sam/sam_method.md。');

await fs.mkdir(outDir,{recursive:true});
for (const [i,slide] of p.slides.items.entries()) { const png=await p.export({slide,format:'png',scale:1}); await writeBlob(`${outDir}/preview_slide_${i+1}.png`,png); }
const pptx=await PresentationFile.exportPptx(p); await pptx.save(`${outDir}/presentation.pptx`);
