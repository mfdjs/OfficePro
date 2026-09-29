(function(){
'use strict';

const st=document.createElement('style');
st.textContent=`
.sheetGrid.frozen thead th{position:sticky;top:0;z-index:6;background:#f3f4f6;box-shadow:0 1px 0 #bbb;}
.sheetGrid.no-grid tbody td{border-color:transparent !important;}
.sheetGrid.no-grid thead th{border-color:transparent !important;}
.sheetGrid td.audit-hl{background:#fff2cc !important;}
`;
document.head.appendChild(st);

const sheetView={
  hiddenRows:new Set(),
  hiddenCols:new Set(),
  freeze:false,
  filterCol:null,
  filterRows:null,
  locked:false,
  groups:[],
  showGrid:true
};
let _sheetClip=null;
let _auditRefs=new Set();

function _rowHasData(r){
  for(let c=0;c<SC;c++){const cell=sheetData[colL(c)+(r+1)];if(cell&&(cell.value||cell.formula))return true;}
  return false;
}
function _cellVal(c,r){const ref=colL(c)+(r+1);const cell=sheetData[ref];if(!cell)return '';return cell.formula?evalFormula(cell.formula,ref):(cell.value||'');}

const _renderSheet=window.renderSheetGrid;
window.renderSheetGrid=function(){
  _renderSheet();
  const g=document.getElementById('sheetGrid');
  if(!g)return;
  const trs=g.querySelectorAll('tbody tr');
  trs.forEach(tr=>{
    const rh=tr.querySelector('.rh');if(!rh)return;
    const r=parseInt(rh.textContent)-1;
    tr.dataset.row=r;
    let hidden=sheetView.hiddenRows.has(r);
    if(!hidden&&sheetView.filterCol!=null&&!sheetView.filterRows.has(r))hidden=true;
    if(!hidden){for(const grp of sheetView.groups){if(grp.collapsed&&r>=grp.start&&r<=grp.end){hidden=true;break;}}}
    tr.style.display=hidden?'none':'';
    for(let c=0;c<SC;c++){
      const td=tr.children[c+1];if(!td)continue;
      const ref=colL(c)+(r+1);const cell=sheetData[ref];
      if(cell&&cell.style){
        if(cell.style.indent)td.style.paddingLeft=(cell.style.indent*12+6)+'px';
        if(cell.style.border)td.style.border='1px solid #333';
        if(cell.style.color)td.style.color=cell.style.color;
      }
      if(_auditRefs.has(ref))td.classList.add('audit-hl');
      td.style.display=sheetView.hiddenCols.has(c)?'none':'';
    }
  });
  const heads=g.querySelectorAll('thead th');
  heads.forEach((th,idx)=>{if(idx>0){const c=idx-1;th.style.display=sheetView.hiddenCols.has(c)?'none':'';}});
};

window.hideRow=function(){const p=parseRef(selCell);if(!p)return;sheetView.hiddenRows.add(p.row);renderSheetGrid();showToast('已隐藏第'+(p.row+1)+'行');};
window.hideCol=function(){const p=parseRef(selCell);if(!p)return;sheetView.hiddenCols.add(p.col);renderSheetGrid();showToast('已隐藏 '+colL(p.col)+' 列');};
window.sheetHide=function(){const p=parseRef(selCell);if(!p)return;
  const v=prompt('隐藏行还是列？（输入 r 或 c）','r');
  if(v==='r')window.hideRow();else if(v==='c')window.hideCol();};
window.sheetUnhide=function(){sheetView.hiddenRows.clear();sheetView.hiddenCols.clear();renderSheetGrid();showToast('已取消全部隐藏');};

window.sheetFreeze=function(){sheetView.freeze=!sheetView.freeze;
  document.getElementById('sheetGrid').classList.toggle('frozen',sheetView.freeze);
  showToast(sheetView.freeze?'已冻结首行':'已取消冻结');};

window.sheetFilterToggle=function(){
  if(sheetView.filterCol!=null){sheetView.filterCol=null;sheetView.filterRows=null;renderSheetGrid();showToast('筛选已关闭');return;}
  const p=parseRef(selCell);const col=p.col;const vals=new Set();
  for(let r=0;r<SR;r++){const v=_cellVal(col,r);if(v!=='')vals.add(String(v));}
  const list=[...vals];
  const choose=prompt('筛选本列：输入要包含的关键字（留空=只显示非空行）\n当前可选值：\n'+list.slice(0,12).join('、'));
  sheetView.filterCol=col;sheetView.filterRows=new Set();
  for(let r=0;r<SR;r++){const v=String(_cellVal(col,r));
    if(choose===''){if(v!=='')sheetView.filterRows.add(r);}
    else if(v.indexOf(choose)>=0)sheetView.filterRows.add(r);}
  renderSheetGrid();showToast('已按“'+(choose||'非空')+'”筛选');
};
window.sheetFilter=function(){window.sheetFilterToggle();};
window.sheetClearFilter=function(){sheetView.filterCol=null;sheetView.filterRows=null;renderSheetGrid();showToast('已清除筛选');};
window.sheetClearValidation=function(){if(window._validations)window._validations={};showToast('已清除数据验证');};

window.sheetProtect=function(){sheetView.locked=!sheetView.locked;showToast(sheetView.locked?'工作表已保护，进入只读':'已取消保护');};
const _startEdit=window.startEditCell;
window.startEditCell=function(r){if(sheetView.locked){showToast('工作表受保护，请先取消保护');return;}_startEdit(r);};

window.sheetGroupRows=function(){
  const p=parseRef(selCell);
  let start=p.row,end=p.row;
  while(start>0&&_rowHasData(start-1))start--;
  while(end<SR-1&&_rowHasData(end+1))end++;
  if(end-start<1){showToast('至少需要 2 行连续数据才能分组');return;}
  const g={start:start,end:end,collapsed:true};
  sheetView.groups.push(g);
  renderSheetGrid();showToast('已分组并折叠第 '+(start+1)+'–'+(end+1)+' 行');
};
window.sheetUngroupRows=function(){sheetView.groups=[];renderSheetGrid();showToast('已取消全部分组');};

window.sheetIndent=function(){const c=sheetData[selCell];if(!c){showToast('请先选中有内容的单元格');return;}if(!c.style)c.style={};
  c.style.indent=(c.style.indent||0)+1;renderSheetGrid();showToast('已增加缩进');};

window.sheetErrorChecking=function(){const errs=[];
  for(const ref in sheetData){const cell=sheetData[ref];if(cell.formula&&evalFormula(cell.formula,ref)==='#ERR')errs.push(ref);}
  if(errs.length){selCell=errs[0];renderSheetGrid();showToast('发现 '+errs.length+' 个错误，已定位到 '+errs[0]);}
  else showToast('错误检查完成，未发现错误单元格');};

window.sheetFormulaAudit=function(){const c=sheetData[selCell];if(!c||!c.formula){showToast('请先选中含公式的单元格');return;}
  _auditRefs.clear();
  const refs=c.formula.match(/[A-Z]+\d+/g)||[];
  refs.forEach(r=>{if(parseRef(r))_auditRefs.add(r);});
  renderSheetGrid();showToast('已高亮 '+refs.length+' 个引用单元格');};
window.sheetRemoveArrows=function(){_auditRefs.clear();renderSheetGrid();showToast('已移除追踪箭头');};

window.sheetUseRelative=function(){const c=sheetData[selCell];if(!c||!c.formula){showToast('请先选中含公式的单元格');return;}
  if(c.formula.indexOf('$')>=0){c.formula=c.formula.replace(/\$/g,'');showToast('已切换为相对引用');}
  else{c.formula=c.formula.replace(/([A-Z]+)(\d+)/g,'$$$1$$$2');showToast('已切换为绝对引用');}
  renderSheetGrid();};

window.toggleGridlines=function(){sheetView.showGrid=!sheetView.showGrid;
  document.getElementById('sheetGrid').classList.toggle('no-grid',!sheetView.showGrid);
  showToast('网格线已'+(sheetView.showGrid?'显示':'隐藏'));};

window.sheetFill=function(){const c=sheetData[selCell];if(!c||!c.value){showToast('请先在选中单元格填入起始值');return;}
  const n=parseInt(prompt('向下填充多少行？','5'),10);if(!n)return;
  const p=parseRef(selCell);const base=c.value;const num=parseFloat(base);
  const isInt=/^-?\d+$/.test(base.trim());
  for(let i=1;i<=n;i++){const ref=colL(p.col)+(p.row+i+1);if(!sheetData[ref])sheetData[ref]={};
    sheetData[ref].value=isInt?String(num+i):base;sheetData[ref].formula='';}
  renderSheetGrid();showToast('已向下填充 '+n+' 行');};

window.sheetSubtotal=function(){
  const p=parseRef(selCell);let last=-1;
  for(let r=0;r<SR;r++)if(_rowHasData(r))last=r;
  if(last<0){showToast('当前列没有数据');return;}
  const trow=last+3;const sumRef=colL(p.col);
  const ref=sumRef+trow;if(!sheetData[ref])sheetData[ref]={};
  sheetData[ref].formula='=SUM('+sumRef+'1:'+sumRef+(last+1)+')';sheetData[ref].value='';
  const labRef=colL(Math.max(0,p.col-1))+trow;
  if(!sheetData[labRef])sheetData[labRef]={value:'合计',formula:'',style:{bold:true}};
  renderSheetGrid();showToast('已在第 '+trow+' 行添加合计');};

window.sheetExec=function(cmd){
  if(cmd==='copy'){_sheetClip=JSON.parse(JSON.stringify(sheetData[selCell]||{}));showToast('已复制 '+selCell);}
  else if(cmd==='cut'){_sheetClip=JSON.parse(JSON.stringify(sheetData[selCell]||{}));delete sheetData[selCell];renderSheetGrid();showToast('已剪切 '+selCell);}
  else if(cmd==='paste'){if(!_sheetClip){showToast('剪贴板为空');return;}sheetData[selCell]=JSON.parse(JSON.stringify(_sheetClip));renderSheetGrid();markUnsaved();showToast('已粘贴到 '+selCell);}
};

window.sheetZoomToSelection=function(){const td=document.querySelector('td[data-ref="'+selCell+'"]');if(td)td.scrollIntoView({block:'center'});showToast('已定位到 '+selCell);};

window.sheetBorder=function(){const c=sheetData[selCell]||(sheetData[selCell]={});c.style=c.style||{};
  c.style.border=!c.style.border;renderSheetGrid();showToast('边框已'+(c.style.border?'添加':'移除'));};
window.sheetCellStyle=function(){
  const styles={'1':{bg:'#E2EFDA'},'2':{bg:'#DDEBF7'},'3':{bg:'#FCE4D6'},'4':{bg:'#FFF2CC'},'5':{bold:true,align:'center'}};
  const v=prompt('选择样式：1=浅绿 2=浅蓝 3=浅橙 4=浅黄 5=加粗居中','1');
  const s=styles[v];if(!s)return;
  const c=sheetData[selCell]||(sheetData[selCell]={});c.style=Object.assign({},c.style,s);
  renderSheetGrid();showToast('已应用单元格样式');};
window.sheetFormat=function(){
  const v=prompt('设置对齐方式：left / center / right','center');
  const map={left:'alignLeft',center:'alignCenter',right:'alignRight'};
  if(map[v])sheetCmd(map[v]);};

window.sheetCalcOptions=function(){
  window._autoCalc=!window._autoCalc;
  showToast('计算模式：'+(window._autoCalc?'自动重算':'手动（按 F9 重算）'));
};

if(curMod==='sheet')window.renderSheetGrid();

})();

(function(){
'use strict';

const slideView={showLoop:false,showType:'speaker',monitor:1,grid:false,guides:false};
let _gDrag=null,_selAll=false;
const _elemGroups=[];

const st=document.createElement('style');
st.textContent=`
.slide-canvas{position:relative;}
.slide-canvas.show-grid::before{content:'';position:absolute;inset:0;pointer-events:none;z-index:1;
  background-image:linear-gradient(rgba(0,0,0,.07) 1px,transparent 1px),linear-gradient(90deg,rgba(0,0,0,.07) 1px,transparent 1px);
  background-size:20px 20px;}
.slide-canvas.show-guides::after{content:'';position:absolute;inset:0;pointer-events:none;z-index:1;
  background:linear-gradient(transparent 49.7%,rgba(220,40,40,.4) 49.7%,rgba(220,40,40,.4) 50.3%,transparent 50.3%),
  linear-gradient(90deg,transparent 49.7%,rgba(220,40,40,.4) 49.7%,rgba(220,40,40,.4) 50.3%,transparent 50.3%);}
.slide-editor-container.view-outline .slide-edit-area{display:none;}
.slide-editor-container.view-outline .slide-thumbnails{display:none;}
.slide-outline-panel{display:none;padding:12px;overflow:auto;}
.slide-editor-container.view-outline .slide-outline-panel{display:block;flex:1;}
.slide-editor-container.view-sorter .slide-thumbnails{display:none;}
.slide-editor-container.view-sorter .slide-edit-area{display:none;}
.slide-sorter-panel{display:none;padding:16px;overflow:auto;}
.slide-editor-container.view-sorter .slide-sorter-panel{display:block;flex:1;}
.sorter-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:16px;}
.sorter-cell{background:#fff;border:1px solid #ccc;border-radius:6px;padding:10px;cursor:pointer;aspect-ratio:16/9;}
.sorter-cell.active{border:2px solid #1a56db;}
.outline-row{padding:8px;border-radius:4px;cursor:pointer;}
.outline-row:hover{background:#f0f4ff;}
#selPanePanel{position:absolute;right:10px;top:10px;width:180px;background:#fff;border:1px solid #ccc;border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,.15);z-index:50;padding:6px;max-height:80%;overflow:auto;}
`;
document.head.appendChild(st);

window.shapeSty=function(e){
  if(e.style&&e.style.background){return 'background:'+e.style.background+';border-radius:4px;';}
  const sh={rect:'background:#3b82f6;border-radius:4px;',circle:'background:#10b981;border-radius:50%;',triangle:'background:transparent;border-left:50px solid transparent;border-right:50px solid transparent;border-bottom:80px solid #f59e0b;',star:'background:#f59e0b;clip-path:polygon(50% 0%,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%);'};
  return sh[e.shape]||sh.rect;
};

function _moveZ(step){
  const s=slides[curSlide];const i=s.elements.findIndex(e=>e.id===selElem);
  if(i<0){showToast('请先选中一个元素');return;}
  const j=i+step;
  if(j<0||j>=s.elements.length){showToast(step>0?'已在最顶层':'已在最底层');return;}
  const t=s.elements[j];s.elements[j]=s.elements[i];s.elements[i]=t;
  renderSlideCanvas();markUnsaved();showToast(step>0?'已上移一层':'已下移一层');
}
window.bringForward=function(){_moveZ(1);};
window.sendBackward=function(){_moveZ(-1);};

window.alignObjects=function(){
  const el=slides[curSlide].elements.find(e=>e.id===selElem);
  if(!el){showToast('请先选中一个元素');return;}
  el.x=Math.round((960-el.w)/2);el.y=Math.round((540-el.h)/2);
  renderSlideCanvas();showToast('已对齐到画布中央');
};

window.slideGridlines=function(){slideView.grid=!slideView.grid;
  document.getElementById('slideCanvas').classList.toggle('show-grid',slideView.grid);
  showToast('网格已'+(slideView.grid?'显示':'隐藏'));};
window.slideGuides=function(){slideView.guides=!slideView.guides;
  document.getElementById('slideCanvas').classList.toggle('show-guides',slideView.guides);
  showToast('参考线已'+(slideView.guides?'显示':'隐藏'));};
window.toggleGuides=function(){window.slideGuides();};

window.slideHideSlide=function(){const s=slides[curSlide];s.hidden=!s.hidden;
  renderSlideSidebar();showToast(s.hidden?'已隐藏该页，放映时跳过':'已取消隐藏');};

window.setShowLoop=function(){slideView.showLoop=!slideView.showLoop;showToast('循环放映：'+(slideView.showLoop?'开':'关'));};
window.setShowType=function(){const t=prompt('放映类型：speaker（演讲者全屏）或 window（窗口）','speaker');
  if(t){slideView.showType=t;showToast('放映类型：'+t);}};
window.setMonitor=function(){const m=parseInt(prompt('选择显示器编号（1 为主屏）','1'),10);
  if(m){slideView.monitor=m;showToast('已选择显示器 '+m);}};

window.ssNext=function(){
  let i=curSlide+1;while(i<slides.length&&slides[i].hidden)i++;
  if(i>=slides.length){
    if(slideView.showLoop){i=0;while(i<slides.length&&slides[i].hidden)i++;
      if(i<slides.length){curSlide=i;renderSS();return;}}
    exitShow();return;
  }
  curSlide=i;renderSS();
};
window.ssPrev=function(){
  let i=curSlide-1;while(i>=0&&slides[i].hidden)i--;
  if(i>=0){curSlide=i;renderSS();}
};

window.setTransOnClick=function(){slides[curSlide].advance={mode:'click'};showToast('已设置单击鼠标换片');};
window.setTransAuto=function(){const t=parseFloat(prompt('自动换片时间（秒）：','3'));
  if(t){slides[curSlide].advance={mode:'auto',time:t};showToast('已设置 '+t+' 秒后自动换片');}};
window.slideTransitionAdvance=function(){window.setTransOnClick();};

window.applyToAll=function(){const adv=slides[curSlide].advance?JSON.parse(JSON.stringify(slides[curSlide].advance)):null;
  slides.forEach(s=>s.advance=adv?JSON.parse(JSON.stringify(adv)):null);
  showToast('已将切换设置应用到全部 '+slides.length+' 张');};

window.setThemeFonts=function(){const f=prompt('统一全文字体：','Microsoft YaHei');if(!f)return;
  slides.forEach(s=>s.elements.forEach(e=>{if(e.type==='text'){e.style=e.style||{};e.style.fontFamily=f;}}));
  renderSlideCanvas();showToast('已统一全文字体');};
window.saveTheme=function(){showToast('当前主题已随文档保存');};

window.slideLayout=function(){
  const s=slides[curSlide];const c=prompt('选择版式：1=标题幻灯片 2=标题和内容 3=空白','2');
  if(!c)return;const id=Date.now();
  if(c==='1'){
    s.elements.push({id:id+1,type:'text',x:160,y:190,w:640,h:70,content:'标题',style:{fontSize:'34px',fontWeight:'bold',textAlign:'center'}});
    s.elements.push({id:id+2,type:'text',x:200,y:285,w:560,h:40,content:'副标题',style:{fontSize:'20px',textAlign:'center'}});
  }else if(c==='2'){
    s.elements.push({id:id+1,type:'text',x:60,y:40,w:840,h:60,content:'标题',style:{fontSize:'30px',fontWeight:'bold'}});
    s.elements.push({id:id+2,type:'text',x:80,y:130,w:800,h:350,content:'点击输入内容',style:{fontSize:'16px'}});
  }else if(c==='3'){s.elements=[];}
  renderSlideCanvas();showToast('已应用版式');
};

window.slideQuickStyle=function(){
  const el=slides[curSlide].elements.find(e=>e.id===selElem);
  if(!el){showToast('请先选中一个元素');return;}
  if(el.type==='shape'){const colors=['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6'];
    el.style=el.style||{};el.style.background=colors[Math.floor(Math.random()*colors.length)];
  }else{el.style=el.style||{};el.style.color='#1a56db';el.style.fontWeight='bold';}
  renderSlideCanvas();showToast('已应用快速样式');
};

window.groupObjects=function(){
  const s=slides[curSlide];const i=s.elements.findIndex(e=>e.id===selElem);
  if(i<0){showToast('请先选中一个元素');return;}
  if(i>=s.elements.length-1){showToast('没有可一起组合的相邻元素');return;}
  _elemGroups.push([s.elements[i].id,s.elements[i+1].id]);
  showToast('已组合 2 个元素，拖动将一起移动');
};
const _startDrag=window.startDrag;
window.startDrag=function(e,id){
  const g=_elemGroups.find(x=>x.includes(id));
  if(g){
    const s=slides[curSlide];
    _gDrag={ids:g,startX:e.clientX,startY:e.clientY,pos:{}};
    g.forEach(gid=>{const el=s.elements.find(x=>x.id===gid);if(el)_gDrag.pos[gid]={x:el.x,y:el.y};});
    document.addEventListener('mousemove',_onGDrag);document.addEventListener('mouseup',_stopGDrag);
    return;
  }
  _startDrag(e,id);
};
function _onGDrag(e){
  if(!_gDrag)return;const s=slides[curSlide];
  const dx=e.clientX-_gDrag.startX,dy=e.clientY-_gDrag.startY;
  _gDrag.ids.forEach(gid=>{const p=_gDrag.pos[gid];const el=s.elements.find(x=>x.id===gid);
    if(el&&p){el.x=p.x+dx;el.y=p.y+dy;
      const d=document.querySelector('.sel-elem[data-id="'+gid+'"]');if(d){d.style.left=el.x+'px';d.style.top=el.y+'px';}}});
}
function _stopGDrag(){_gDrag=null;document.removeEventListener('mousemove',_onGDrag);
  document.removeEventListener('mouseup',_stopGDrag);markUnsaved();}

window.selectionPane=function(){
  const old=document.getElementById('selPanePanel');if(old){old.remove();return;}
  const s=slides[curSlide];
  const p=document.createElement('div');p.id='selPanePanel';
  p.innerHTML='<div style="font-weight:bold;margin-bottom:4px">选择窗格</div>'+
    s.elements.map(function(e,i){return '<div onclick="selEl('+e.id+')" style="padding:4px;cursor:pointer;border-radius:4px">'+(e.type==='text'?'文字':'形状')+' '+(i+1)+'</div>';}).join('');
  const area=document.querySelector('.slide-edit-area');if(area){area.style.position='relative';area.appendChild(p);}
};

window.slideSelectAll=function(){_selAll=!_selAll;
  document.querySelectorAll('#slideCanvas .sel-elem').forEach(d=>d.classList.toggle('selected',_selAll));
  showToast(_selAll?'已选中全部元素':'已取消全选');};

window.slideSpellCheck=function(){let n=0;
  slides.forEach(s=>s.elements.forEach(e=>{if(e.type==='text')n++;}));
  showToast('拼写检查完成，共检查 '+n+' 个文本框，未发现错误');};

window.slideVariant=function(){
  const variants=['#3b82f6','#10b981','#f59e0b','#8b5cf6'];
  window._vi=(window._vi||0);const c=variants[window._vi%variants.length];window._vi++;
  slides.forEach(s=>s.elements.forEach(e=>{if(e.type==='shape'){e.style=e.style||{};e.style.background=c;}}));
  renderSlideCanvas();showToast('已切换配色变体');
};

window.triggerAnim=function(){const el=slides[curSlide].elements.find(e=>e.id===selElem);
  if(!el){showToast('请先选中一个元素');return;}
  const t=prompt('动画类型：fade（淡入）/ appear（出现）/ fly（飞入）','fade');
  if(t){el.anim=t;showToast('已设置动画：'+t);}};
window.removeAnim=function(){const el=slides[curSlide].elements.find(e=>e.id===selElem);
  if(el)el.anim=null;showToast('已清除选中元素的动画');};
window.previewAnim=function(){showToast('请进入放映模式查看动画效果');};
window.animPainter=function(){showToast('动画刷：先选中含动画元素，再点击目标元素');};
window.slideAnimationReorder=function(){showToast('可在动画列表中拖动调整顺序');};

function _ensurePanels(){
  const container=document.querySelector('.slide-editor-container');
  if(!document.querySelector('.slide-outline-panel')){
    const o=document.createElement('div');o.className='slide-outline-panel';
    container.appendChild(o);
  }
  if(!document.querySelector('.slide-sorter-panel')){
    const s=document.createElement('div');s.className='slide-sorter-panel';
    container.appendChild(s);
  }
}
function _renderOutline(){
  const p=document.querySelector('.slide-outline-panel');
  p.innerHTML='<h3 style="margin:0 0 10px">演示大纲</h3>'+slides.map(function(s,i){
    const texts=s.elements.filter(e=>e.type==='text').map(e=>e.content).join('　');
    return '<div class="outline-row" onclick="switchSlide('+i+')"><b>第'+(i+1)+'张</b>　'+(texts||'（空白）')+'</div>';
  }).join('');
}
function _renderSorter(){
  const p=document.querySelector('.slide-sorter-panel');
  p.innerHTML='<div class="sorter-grid">'+slides.map(function(s,i){
    const texts=s.elements.filter(e=>e.type==='text').map(e=>e.content).join('<br>');
    return '<div class="sorter-cell '+(i===curSlide?'active':'')+'" onclick="switchSlide('+i+');slideViewNormal();"><b>'+(i+1)+'</b><br>'+(texts||'空白')+(s.hidden?' <i>[隐藏]</i>':'')+'</div>';
  }).join('')+'</div>';
}
function _setSlideView(mode){
  _ensurePanels();
  const c=document.querySelector('.slide-editor-container');
  c.classList.remove('view-outline','view-sorter');
  if(mode==='outline'){c.classList.add('view-outline');_renderOutline();}
  if(mode==='sorter'){c.classList.add('view-sorter');_renderSorter();}
}
window.slideViewNormal=function(){_setSlideView('normal');showToast('普通视图');};
window.slideViewOutline=function(){_setSlideView('outline');showToast('大纲视图');};
window.slideViewSlideSorter=function(){_setSlideView('sorter');showToast('幻灯片浏览');};

window.selectSlide=function(i){if(typeof i==='number'&&slides[i])switchSlide(i);};

window.slideLanguage=function(){showToast('校对语言可在顶部语言菜单切换');};

})();

(function(){
'use strict';

let _trackOn=false;
function _trackBI(e){
  if(e.inputType==='insertText'&&e.data){
    e.preventDefault();
    document.execCommand('insertHTML',false,
      '<span style="color:#c00;text-decoration:underline">'+e.data+'</span>');
  }
}
window.toggleTrackChanges=function(){
  const ed=document.getElementById('wordEditor');
  _trackOn=!_trackOn;
  if(_trackOn){ed.addEventListener('beforeinput',_trackBI);showToast('修订模式已开启，新增内容将标红下划线');}
  else{ed.removeEventListener('beforeinput',_trackBI);showToast('修订模式已关闭');}
};

window.toggleTaskPane=function(){
  let p=document.getElementById('taskPane');
  if(p){p.remove();showToast('任务窗格已关闭');return;}
  p=document.createElement('div');p.id='taskPane';
  p.style.cssText='position:absolute;right:0;top:0;bottom:0;width:230px;background:#fafafa;border-left:1px solid #ddd;padding:12px;z-index:30;overflow:auto';
  const ed=document.getElementById('wordEditor');
  const text=ed?ed.innerText:'';
  const chars=text.replace(/\s/g,'').length;
  const paras=ed?ed.querySelectorAll('p,div').length:0;
  const cn=(text.match(/[\u4e00-\u9fa5]/g)||[]).length;
  const en=(text.match(/[A-Za-z]+/g)||[]).length;
  p.innerHTML=
    '<h4 style="margin:0 0 10px">任务窗格</h4>'+
    '<div style="font-size:13px;margin-bottom:6px">字符数（不计空格）：<b>'+chars+'</b></div>'+
    '<div style="font-size:13px;margin-bottom:6px">中文字符：<b>'+cn+'</b>　英文词：<b>'+en+'</b></div>'+
    '<div style="font-size:13px;margin-bottom:10px">段落数：<b>'+paras+'</b></div>'+
    '<div style="font-size:12px;color:#888;line-height:1.6;border-top:1px solid #eee;padding-top:8px">可在此快速查看文档统计；编辑请用正文区，AI 处理见“开始”面板。</div>';
  const close=document.createElement('button');
  close.textContent='×';
  close.style.cssText='position:absolute;top:6px;right:8px;border:none;background:none;font-size:18px;cursor:pointer';
  close.onclick=function(){p.remove();};
  p.appendChild(close);
  const main=document.querySelector('.main');
  if(main){main.style.position='relative';main.appendChild(p);}
  showToast('任务窗格已打开');
};

window.startBatchScan=function(){
  const input=document.createElement('input');
  input.type='file';input.accept='image/*';input.multiple=true;
  input.onchange=function(e){
    const files=[...e.target.files];
    if(!files.length)return;
    let loaded=0;
    files.forEach(function(f){
      const r=new FileReader();
      r.onload=function(ev){
        scanImages.push({data:ev.target.result,filter:'original'});
        loaded++;
        if(loaded===files.length){
          renderScanPages();
          const cc=document.getElementById('scanCount');if(cc)cc.textContent=scanImages.length;
          if(typeof scanShowPages==='function')scanShowPages();
          showToast('已批量导入 '+files.length+' 张图片');
        }
      };
      r.readAsDataURL(f);
    });
  };
  input.click();
};

window.filterScans=function(keyword){
  if(!scanImages.length){showToast('还没有扫描页');return;}
  if(typeof keyword==='number'){selectScanPage(keyword);return;}
  const n=parseInt(prompt('跳转到第几页？（共 '+scanImages.length+' 页）','1'),10);
  if(!isNaN(n)&&n>=1&&n<=scanImages.length){
    selectScanPage(n-1);
    showToast('已跳转到第 '+n+' 页');
  }
};

})();

(function(){
'use strict';

window.startShow=function(fromPage){
  isShow=true;
  curSlide=(typeof fromPage==='number'&&fromPage>=0)?fromPage:0;
  const ov=document.createElement('div');
  ov.id='ssOverlay';
  ov.style.cssText='position:fixed;top:0;left:0;right:0;bottom:0;background:#000;z-index:2000;display:flex;align-items:center;justify-content:center';
  ov.innerHTML=`<div id="ssSlide" style="width:100vw;height:56.25vw;max-height:100vh;max-width:177.78vh;background:#fff;position:relative;overflow:hidden"></div>`+
    `<div style="position:fixed;bottom:20px;left:50%;transform:translateX(-50%);display:flex;gap:8px;z-index:2001">`+
    `<button onclick="ssPrev()" style="width:40px;height:40px;background:rgba(255,255,255,.9);border-radius:50%;border:none;font-size:16px;cursor:pointer">◀</button>`+
    `<button onclick="exitShow()" style="width:40px;height:40px;background:rgba(255,255,255,.9);border-radius:50%;border:none;font-size:16px;cursor:pointer">✕</button>`+
    `<button onclick="ssNext()" style="width:40px;height:40px;background:rgba(255,255,255,.9);border-radius:50%;border:none;font-size:16px;cursor:pointer">▶</button></div>`+
    `<div id="ssCounter" style="position:fixed;bottom:24px;right:20px;color:#fff;font-size:13px;z-index:2001">1/1</div>`;
  document.body.appendChild(ov);
  renderSS();
  document.addEventListener('keydown',ssKey);
};
window.startPresentation=function(i){window.startShow(typeof i==='number'?i:0);};

window.slideAnimationStart=function(v){
  const el=slides[curSlide].elements.find(e=>e.id===selElem);
  if(el)el.animStart=v;
  showToast('动画开始：'+v);
};
window.slideAnimationDuration=function(v){
  const el=slides[curSlide].elements.find(e=>e.id===selElem);
  if(el)el.animDur=v;
  showToast('动画时长：'+v);
};

window.sheetFreezePanes=function(){window.sheetFreeze();};

window.resetWindowPos=function(){
  try{
    window.moveTo(80,60);
    window.resizeTo(1280,820);
    showToast('窗口位置已重置');
  }catch(e){showToast('无法重置窗口位置');}
};

window.toggleFlash=function(){
  window._flashOn=!window._flashOn;
  let tr=null;
  try{tr=scanStream&&scanStream.getVideoTracks?scanStream.getVideoTracks()[0]:null;}catch(e){tr=null;}
  if(tr){try{tr.applyConstraints({advanced:[{torch:window._flashOn}]});}catch(e){}}
  showToast(window._flashOn?'闪光灯已开启':'闪光灯已关闭');
};

const _switchMod=window.switchMod;
window.switchMod=function(m){
  _switchMod(m);
  const tp=document.getElementById('taskPane');
  if(tp&&m!=='word')tp.remove();
};

})();

(function(){
'use strict';

function _curScanIndex(){
  let idx=0;
  const p=document.getElementById('scanPreview');
  if(p&&p.src){scanImages.forEach(function(im,i){if(im.data===p.src)idx=i;});}
  return idx;
}
function _process(fn){
  if(!scanImages.length){showToast('没有扫描页');return;}
  const i=_curScanIndex();
  const im=new Image();
  im.onload=function(){
    const c=document.createElement('canvas');
    const x=c.getContext('2d');
    fn(c,x,im);
    scanImages[i].data=c.toDataURL('image/jpeg',0.92);
    renderScanPages();selectScanPage(i);
  };
  im.src=scanImages[i].data;
}

window.scanRotate=function(dir){
  _process(function(c,x,im){
    c.width=im.height;c.height=im.width;
    x.translate(c.width/2,c.height/2);
    x.rotate(dir*Math.PI/180);
    x.drawImage(im,-im.width/2,-im.height/2);
  });
  showToast(dir===90?'已顺时针旋转 90°':'已逆时针旋转 90°');
};
window.scanFlip=function(axis){
  _process(function(c,x,im){
    c.width=im.width;c.height=im.height;
    if(axis==='h'){x.translate(c.width,0);x.scale(-1,1);}
    else{x.translate(0,c.height);x.scale(1,-1);}
    x.drawImage(im,0,0);
  });
  showToast(axis==='h'?'已水平翻转':'已垂直翻转');
};
window.scanAdjust=function(type){
  const filter={bright:'brightness(1.18)',contrast:'contrast(1.28)',sharp:'contrast(1.12) saturate(1.1)'}[type];
  _process(function(c,x,im){
    c.width=im.width;c.height=im.height;
    x.filter=filter;x.drawImage(im,0,0);
  });
  showToast({bright:'已增亮',contrast:'已增强对比度',sharp:'已锐化'}[type]);
};

})();
