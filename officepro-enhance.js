

(function(){
  'use strict';

  function importDocument(){
    if(typeof openExternalFile==='function'){ openExternalFile(); return; }
    if(window.officepro && window.officepro.openFileDialog){
      window.officepro.openFileDialog().then(function(result){
        if(result && result.success){
          loadOpenedFile(result);
        } else if(result && result.error && result.error !== '用户取消选择'){
          showToast('打开失败: ' + result.error);
        }
      }).catch(function(e){ showToast('打开失败: ' + e.message); });
    } else {
      var input=document.createElement('input');
      input.type='file';
      input.accept='.txt,.md,.csv,.json,.xml,.html,.htm,.rtf,.log,.doc,.docx,.odt,.xls,.xlsx,.xlsm,.xlsb,.ods,.ppt,.pptx,.odp,.pdf,.epub,.mobi,.jpg,.jpeg,.png,.bmp,.gif,.webp,.tiff,.tif';
      input.onchange=function(e){
        var f=e.target.files[0];
        if(!f)return;
        var reader=new FileReader();
        var ext=f.name.split('.').pop().toLowerCase();
        if(['jpg','jpeg','png','bmp','gif','webp','tiff','tif'].indexOf(ext)>=0){
          reader.onload=function(ev){ loadOpenedFile({type:'image',content:ev.target.result,fileName:f.name,fileSize:f.size}); };
          reader.readAsDataURL(f);
        } else {
          reader.onload=function(ev){ loadOpenedFile({type:'text',content:ev.target.result,fileName:f.name,fileSize:f.size}); };
          reader.readAsText(f);
        }
      };
      input.click();
    }
  }
  window.importDocument = importDocument;

  function loadOpenedFile(data){
    if(typeof loadExternalFile==='function'){ loadExternalFile(data); return; }
    if(!data || !data.type) return;
    var type=data.type, content=data.content||'', fileName=data.fileName||'未命名';
    showToast('正在打开: ' + fileName);
    try{
      if(type==='word' || type==='text' || type==='epub'){
        switchMod('word');
        setTimeout(function(){
          var editor=document.getElementById('wordEditor');
          if(editor){
            if(type==='text' && !content.startsWith('<')){
              editor.innerHTML='<pre style="white-space:pre-wrap;font-family:inherit">'+content.replace(/</g,'&lt;').replace(/>/g,'&gt;')+'</pre>';
            } else {
              editor.innerHTML=content;
            }
            if(typeof markUnsaved==='function') markUnsaved();
            if(typeof updateWordStatus==='function') updateWordStatus();
          }
          var fnEl=document.getElementById('currentFileName');
          if(fnEl)fnEl.textContent=fileName;
        },100);
      }
      else if(type==='sheet'){
        switchMod('sheet');
        setTimeout(function(){
          loadSheetFromCSV(content);
          var fnEl=document.getElementById('currentFileName');
          if(fnEl)fnEl.textContent=fileName;
        },100);
      }
      else if(type==='slide'){
        switchMod('slide');
        setTimeout(function(){
          loadSlidesFromData(content, data.slides);
          var fnEl=document.getElementById('currentFileName');
          if(fnEl)fnEl.textContent=fileName;
        },100);
      }
      else if(type==='pdf'){
        switchMod('pdf');
        setTimeout(function(){
          var viewer=document.querySelector('#page-pdf .pdf-viewer') || document.getElementById('pdfViewer');
          if(viewer){
            var pages=content.split(/\f/);
            var html='';
            pages.forEach(function(p){
              if(p.trim()){
                html+='<div class="pdf-page" style="padding:25mm;margin-bottom:12px;white-space:pre-wrap;font-size:14px;line-height:1.8">'+p.replace(/</g,'&lt;')+'</div>';
              }
            });
            viewer.innerHTML=html || '<div class="pdf-empty"><div class="pdf-empty-icon">📄</div><div class="pdf-empty-text">PDF内容为空</div></div>';
          }
          var fnEl=document.getElementById('currentFileName');
          if(fnEl)fnEl.textContent=fileName;
        },100);
      }
      else if(type==='image'){
        switchMod('scan');
        setTimeout(function(){
          if(typeof loadScanImage === 'function'){
            loadScanImage(content, fileName);
          } else {
            var scanArea=document.querySelector('#page-scan .scan-video-wrap') || document.getElementById('scanPreview');
            if(scanArea){
              scanArea.innerHTML='<img src="'+content+'" style="max-width:100%;max-height:100%;object-fit:contain">';
            }
          }
          var fnEl=document.getElementById('currentFileName');
          if(fnEl)fnEl.textContent=fileName;
        },100);
      }
      else if(type==='error'){
        showToast('文件解析失败: ' + (data.error||'未知错误'));
      }
      else {
        showToast('不支持的文件格式');
      }
    }catch(e){
      showToast('加载失败: ' + e.message);
      console.error('loadOpenedFile error', e);
    }
  }
  window.loadOpenedFile = loadOpenedFile;

  function loadSheetFromCSV(csvText){
    try{
      var delimiter='\t';
      if(csvText.indexOf('\t')<0 && csvText.indexOf(',')>=0) delimiter=',';
      var lines=csvText.split(/\r?\n/);
      var grid=document.getElementById('excelGrid') || document.querySelector('#page-sheet .sheet-grid');
      if(!grid) return;
      var tbody=grid.querySelector('tbody') || grid;
      var rows=tbody.querySelectorAll('tr');
      var maxRows=Math.min(lines.length, rows.length||100);
      for(var r=0;r<maxRows;r++){
        if(!lines[r]) continue;
        var cells=lines[r].split(delimiter);
        var rowCells=rows[r]?rows[r].querySelectorAll('td'):[];
        for(var c=0;c<Math.min(cells.length,rowCells.length);c++){
          var cell=rowCells[c];
          if(cell){
            var val=cells[c].replace(/^"|"$/g,'').replace(/""/g,'"');
            cell.textContent=val;
            cell.removeAttribute('data-formula');
          }
        }
      }
      if(typeof calculateSheet==='function') calculateSheet();
      if(typeof updateSheetStatus==='function') updateSheetStatus();
      showToast('表格已加载 (' + maxRows + ' 行)');
    }catch(e){
      showToast('表格加载失败: ' + e.message);
    }
  }

  function loadSlidesFromData(content, slidesData){
    try{
      var slideTexts=[];
      if(slidesData && Array.isArray(slidesData)){
        slideTexts=slidesData;
      } else if(content){
        try{ slideTexts=JSON.parse(content); }catch(e){ slideTexts=[content]; }
      }
      if(!slideTexts.length){ showToast('演示内容为空'); return; }
      if(typeof slides !== 'undefined'){
        slides=[];
        slideTexts.forEach(function(text,i){
          var lines=text.split('\n').filter(function(l){return l.trim();});
          var title=lines[0]||('幻灯片 '+(i+1));
          var body=lines.slice(1).join('\n');
          slides.push({
            bg:'#ffffff',
            elements:[
              {type:'text',x:60,y:60,w:840,h:80,text:title,size:36,bold:true,color:'#1a202c'},
              {type:'text',x:60,y:160,w:840,h:320,text:body,size:18,color:'#333'}
            ]
          });
        });
        curSlide=0;
        if(typeof renderSlideSidebar==='function') renderSlideSidebar();
        if(typeof renderSlideCanvas==='function') renderSlideCanvas();
        showToast('演示文稿已加载 (' + slides.length + ' 页)');
      }
    }catch(e){
      showToast('演示加载失败: ' + e.message);
    }
  }

  function overrideOCR(){
    if(typeof startOCR !== 'function') return;
    var _origStartOCR = startOCR;
    startOCR = function(){
      if(window.officepro && window.officepro.ocrImage){
        var input=document.getElementById('ocrInput');
        if(!input||!input.files||input.files.length===0){showToast('请先选择图片');return;}
        var area=document.getElementById('ocrResultArea'),box=document.getElementById('ocrResult');
        if(area)area.style.display='block';
        if(box)box.innerHTML='<p style="font-size:13px;color:#666;">正在使用Tesseract引擎识别文字（中文+英文）...</p><div class="batch-progress"><div class="batch-progress-bar" id="tesseractProgress" style="width:0%"></div></div>';
        showToast('正在识别...');
        var files=Array.prototype.slice.call(input.files);
        var allResults=[], allImages=[];
        var idx=0;
        function processNext(){
          if(idx>=files.length){
            window._jxcOCRText=allResults.join('\n\n');
            var h='';
            allImages.forEach(function(u){h+='<img src="'+u+'" style="max-width:100%;border:1px solid #e2e8f0;border-radius:8px;margin:6px 0;">';});
            h+='<pre style="white-space:pre-wrap;font-size:12.5px;line-height:1.7;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px;margin-top:6px;">'+window._jxcOCRText.replace(/</g,'&lt;')+'</pre>';
            h+='<div style="margin-top:8px;display:flex;gap:8px;"><button class="btn btn-primary" onclick="insertOCRResult()">插入到文档</button><button class="btn btn-secondary" onclick="copyOCRResult()">复制结果</button></div>';
            if(box)box.innerHTML=h;
            showToast('识别完成 ('+allResults.length+' 张)');
            return;
          }
          var f=files[idx];
          var reader=new FileReader();
          reader.onload=function(ev){
            var dataUrl=ev.target.result;
            allImages.push(dataUrl);
            window.officepro.ocrImage(dataUrl).then(function(res){
              if(res && res.success){
                allResults.push(res.text);
              } else {
                allResults.push('[识别失败: '+(res&&res.error?res.error:'未知')+']');
              }
              idx++;
              var prog=document.getElementById('tesseractProgress');
              if(prog)prog.style.width=Math.round(idx/files.length*100)+'%';
              processNext();
            }).catch(function(e){
              allResults.push('[识别错误: '+e.message+']');
              idx++;
              processNext();
            });
          };
          reader.readAsDataURL(f);
        }
        processNext();
      } else {
        _origStartOCR();
      }
    };

    if(typeof insertOCRResult === 'function'){
      var _origInsertOCR=insertOCRResult;
      insertOCRResult=function(){
        var editor=document.getElementById('wordEditor');
        if(editor && window._jxcOCRText){
          switchMod('word');
          setTimeout(function(){
            editor.innerHTML+='<p>'+window._jxcOCRText.replace(/\n/g,'</p><p>')+'</p>';
            if(typeof markUnsaved==='function') markUnsaved();
            showToast('OCR结果已插入文档');
            closeModal('commonModal');
          },150);
        } else {
          _origInsertOCR();
        }
      };
    }
  }

  function setupFileAssociation(){
    if(window.officepro && window.officepro.onFileOpened){
      window.officepro.onFileOpened(function(data){
        if(data && data.filePath){
          loadOpenedFile(data);
        }
      });
    }
  }

  function isMobile(){
    return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent) || window.innerWidth < 768;
  }

  function applyMobileAdaptation(){
    if(!isMobile()) return;
    if(document.getElementById('mobileAdaptStyle')) return;
    var style=document.createElement('style');
    style.id='mobileAdaptStyle';
    style.textContent=[
      '@media(max-width:768px){',
      'body{overflow:hidden!important;margin:0!important;padding:0!important}',
      '#splashScreen{display:none!important;position:fixed!important;z-index:999999!important}',
      '#ribbonTabs{display:none!important}',
      '#toolbar{display:none!important}',
      '.main{overflow:hidden!important;position:relative!important;flex:1!important}',
      '.page{position:relative!important;display:none!important;width:100%!important;height:100%!important;overflow:hidden!important}',
      '.page.active{display:flex!important}',
      '.doc-tabs-bar{flex-shrink:0!important}',
      '.modtabs{display:flex!important;flex-wrap:nowrap!important;overflow-x:auto!important;overflow-y:hidden!important;-webkit-overflow-scrolling:touch!important;flex-shrink:0!important}',
      '.modtab{padding:8px 12px!important;font-size:12px!important;white-space:nowrap!important;flex-shrink:0!important}',
      '.wps-home-sidebar{width:48px!important}',
      '.wps-home-app-name{font-size:9px!important}',
      '.wps-blank-grid{grid-template-columns:repeat(2,1fr)!important}',
      '.wps-template-grid{grid-template-columns:repeat(2,1fr)!important}',
      '.wps-home-tools{grid-template-columns:repeat(2,1fr)!important}',
      '.word-page{padding:10mm!important;width:100%!important;max-width:100%!important;box-sizing:border-box!important}',
      '#wordEditor{width:100%!important;max-width:100%!important;box-sizing:border-box!important;overflow-wrap:break-word!important}',
      '.slide-sidebar{width:80px!important;min-width:80px!important}',
      '#slideCanvas{width:100%!important;height:auto!important;aspect-ratio:16/9;max-width:100%!important}',
      '.scan-right{width:100px!important;min-width:100px!important}',
      '.scan-controls{height:50px!important;gap:8px!important}',
      '.scan-btn{width:36px!important;height:36px!important;font-size:14px!important}',
      '.scan-btn.capture{width:44px!important;height:44px!important;font-size:18px!important}',
      '.toolbar{padding:4px!important;gap:2px!important;flex-wrap:wrap!important}',
      '.tool-btn{width:28px!important;height:28px!important;font-size:12px!important}',
      '.page .ribbon-tabs{display:flex!important;flex-wrap:nowrap!important;overflow-x:auto!important;overflow-y:hidden!important;-webkit-overflow-scrolling:touch!important;flex-shrink:0!important}',
      '.ribbon-tab{padding:6px 10px!important;font-size:11px!important;white-space:nowrap!important;flex-shrink:0!important}',
      '.ribbon-content{overflow-x:auto!important;overflow-y:hidden!important;-webkit-overflow-scrolling:touch!important;flex-shrink:0!important}',
      '.ribbon-panel{display:none!important;flex-wrap:wrap!important;gap:4px!important;padding:6px!important}',
      '.ribbon-panel.active{display:flex!important}',
      '.ribbon-group{flex-shrink:0!important;min-width:0!important}',
      '.ribbon-buttons{display:flex!important;flex-wrap:wrap!important;gap:2px!important}',
      '.ribbon-btn{padding:4px 6px!important;font-size:10px!important;min-width:0!important}',
      '.ribbon-select{max-width:80px!important;font-size:11px!important}',
      '.modal{min-width:92vw!important;max-width:96vw!important;max-height:85vh!important;overflow-y:auto!important}',
      '.modal-body{padding:12px!important;overflow-y:auto!important}',
      '.print-preview{width:100%!important;padding:8mm!important;box-sizing:border-box!important}',
      '.titlebar{height:40px!important;flex-shrink:0!important}',
      '.wps-statusbar{height:22px!important;font-size:10px!important;flex-shrink:0!important}',
      'input,select,textarea{font-size:14px!important;max-width:100%!important;box-sizing:border-box!important}',
      '.wps-home-search-input{font-size:14px!important}',
      '#excelGrid{display:block!important;overflow-x:auto!important;overflow-y:auto!important;max-width:100%!important;-webkit-overflow-scrolling:touch!important}',
      '#excelGrid table{min-width:600px!important;border-collapse:collapse!important}',
      '.sheet-cell{min-width:40px!important;max-width:80px!important;overflow:hidden!important;text-overflow:ellipsis!important;white-space:nowrap!important}',
      '.formula-bar{flex-wrap:wrap!important;gap:4px!important}',
      '#formulaInput{flex:1!important;min-width:100px!important}',
      '.pdf-viewer{overflow-x:hidden!important;overflow-y:auto!important;-webkit-overflow-scrolling:touch!important}',
      '.pdf-page{width:100%!important;max-width:100%!important;box-sizing:border-box!important;padding:15mm!important;overflow-wrap:break-word!important}',
      '.scan-video-wrap{width:100%!important;max-width:100%!important;overflow:hidden!important}',
      '.scan-video-wrap video,.scan-video-wrap img{max-width:100%!important;height:auto!important;object-fit:contain!important}',
      '*{max-width:100vw!important}',
      'img,video,canvas,table{max-width:100%!important;height:auto!important}',
      '}'
    ].join('');
    document.head.appendChild(style);

    var touchStyle = document.createElement('style');
    touchStyle.textContent = 'html,body,*{touch-action:manipulation!important;-webkit-user-select:none!important;user-select:none!important}input,textarea{touch-action:auto!important;-webkit-user-select:text!important;user-select:text!important}';
    document.head.appendChild(touchStyle);
  }

  function autoFitScreen(){
    try{
      var w = window.innerWidth;
      var h = window.innerHeight;
      var dpr = window.devicePixelRatio || 1;

      var baseFontSize = Math.max(12, Math.min(16, w / 40));
      document.documentElement.style.fontSize = baseFontSize + 'px';

      var meta = document.querySelector('meta[name="viewport"]');
      if(meta){
        meta.setAttribute('content', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0, user-scalable=no, viewport-fit=cover');
      }

      var ribbonContent = document.querySelectorAll('.ribbon-content');
      var maxRibbonHeight = Math.max(120, Math.min(200, h * 0.22));
      ribbonContent.forEach(function(el){
        el.style.maxHeight = maxRibbonHeight + 'px';
      });

      var titlebar = document.querySelector('.wps-titlebar');
      var statusbar = document.querySelector('.wps-statusbar');
      var modtabs = document.querySelector('.modtabs');
      var tbH = Math.max(32, Math.min(44, h * 0.045));
      var sbH = Math.max(20, Math.min(28, h * 0.03));
      var mtH = Math.max(32, Math.min(40, h * 0.04));
      if(titlebar) titlebar.style.height = tbH + 'px';
      if(statusbar) statusbar.style.height = sbH + 'px';
      if(modtabs) modtabs.style.height = mtH + 'px';

      var main = document.querySelector('.main');
      if(main){
        main.style.top = (tbH + mtH) + 'px';
        main.style.bottom = sbH + 'px';
      }

      window._screenFit = {width:w, height:h, dpr:dpr, baseFontSize:baseFontSize, tbH:tbH, sbH:sbH, mtH:mtH};
    }catch(e){
      console.warn('autoFitScreen error', e);
    }
  }

  function fitScreenNow(){
    autoFitScreen();
    if(typeof showToast === 'function') showToast('已自适应屏幕: ' + window.innerWidth + 'x' + window.innerHeight);
  }
  window.fitScreenNow = fitScreenNow;

  function setupShortcuts(){
    document.addEventListener('keydown',function(e){
      if((e.ctrlKey||e.metaKey) && e.key==='o'){
        e.preventDefault();
        importDocument();
      }
    });
  }

  function resetLayout(){
    try{
      window.scrollTo(0,0);
      document.documentElement.scrollTop = 0;
      document.body.scrollTop = 0;

      var hiddenIds = ['splashScreen','ribbonTabs','toolbar','docTabsBar','titlebar','tabbar','sidebar','statusbar','excelGrid','pptSlides','accessTables'];
      hiddenIds.forEach(function(id){
        var el = document.getElementById(id);
        if(el){
          el.style.display = 'none';
          el.style.height = '0';
          el.style.margin = '0';
          el.style.padding = '0';
          el.style.overflow = 'hidden';
        }
      });

      document.body.style.margin = '0';
      document.body.style.padding = '0';
      document.body.style.transform = 'none';

      var current = typeof curMod !== 'undefined' ? curMod : 'home';
      document.querySelectorAll('.page').forEach(function(p){
        if(p.id === 'page-' + current){
          p.style.display = 'flex';
          p.classList.add('active');
        } else {
          p.style.display = 'none';
          p.classList.remove('active');
        }
      });
    }catch(e){}
  }

  function forcePageVisibility(){
    var pages = document.querySelectorAll('.page');
    for(var i=0;i<pages.length;i++){
      var p = pages[i];
      var on = p.classList.contains('active');
      p.style.setProperty('display', on?'flex':'none', 'important');
      p.style.setProperty('visibility', on?'visible':'hidden', 'important');
      p.style.setProperty('opacity', on?'1':'0', 'important');
      p.style.setProperty('z-index', on?'5':'1', 'important');
      p.style.setProperty('pointer-events', on?'auto':'none', 'important');
    }
  }
  window.forcePageVisibility = forcePageVisibility;

  function patchSwitchMod(){
    if(typeof switchMod !== 'function') return false;
    var _origSwitchMod = switchMod;
    switchMod = function(m){
      _origSwitchMod(m);

      setTimeout(forcePageVisibility, 0);
      setTimeout(forcePageVisibility, 50);
      setTimeout(forcePageVisibility, 200);
    };

    forcePageVisibility();
    return true;
  }

  function init(){
    overrideOCR();
    setupShortcuts();

    if(!isMobile()) return;
    applyMobileAdaptation();
    resetLayout();

    autoFitScreen();
    window.addEventListener('resize', function(){
      clearTimeout(window._fitTimer);
      window._fitTimer = setTimeout(function(){
        autoFitScreen();
        forcePageVisibility();
      }, 100);
    });
    window.addEventListener('orientationchange', function(){
      setTimeout(function(){
        autoFitScreen();
        forcePageVisibility();
      }, 300);
    });

    if(!patchSwitchMod()){

      setTimeout(function(){
        if(!patchSwitchMod()){
          setTimeout(patchSwitchMod, 500);
        }
      }, 100);
    }

    setTimeout(resetLayout, 100);
    setTimeout(forcePageVisibility, 100);
    setTimeout(resetLayout, 500);
    setTimeout(forcePageVisibility, 500);
    setTimeout(resetLayout, 1000);
    setTimeout(forcePageVisibility, 1000);
    setTimeout(autoFitScreen, 200);
    setTimeout(autoFitScreen, 800);
  }

  if(document.readyState==='loading'){
    document.addEventListener('DOMContentLoaded',init);
  } else {
    init();
  }

})();
