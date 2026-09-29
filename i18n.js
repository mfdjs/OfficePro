(function(){
  const I18N={
    lang:'zh',
    dicts:{},
    rtlLangs:['ar'],
    skipSelector:'#wordEditor,#sheetGrid,#slideCanvas,#excelGrid,#pptSlides,.pdf-viewer,#accessArea,#diagramArea,.content-area,[contenteditable],.code-editor,#devTools',
    register(lang,dict){this.dicts[lang]=Object.assign(this.dicts[lang]||{},dict);},
    isRtl(){return this.rtlLangs.indexOf(this.lang)!==-1;},
    setDir(){
      try{document.documentElement.dir=this.isRtl()?'rtl':'ltr';}catch(e){}
    },
    t(text){
      if(this.lang==='zh')return text;
      const key=(text||'').trim();
      const d=this.dicts[this.lang];
      if(d&&d[key]!=null)return d[key];
      if(this.dicts.en&&this.dicts.en[key]!=null)return this.dicts.en[key];
      return text;
    },
    isSkipNode(node){
      let el=node.nodeType===3?node.parentElement:node;
      while(el){
        if(el.id||el.classList){
          if(el.matches&&el.matches(this.skipSelector))return true;
        }
        el=el.parentElement;
      }
      return false;
    },
    translateTree(root){
      if(!root)root=document.body;
      const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT,null);
      const targets=[];
      let n;
      while((n=walker.nextNode())){
        if(n.nodeValue&&n.nodeValue.trim()&&/[\u4e00-\u9fa5]/.test(n.nodeValue)){
          if(!this.isSkipNode(n))targets.push(n);
        }
      }
      targets.forEach(node=>{
        const raw=node.nodeValue;
        const lead=raw.match(/^\s*/)[0];
        const trail=raw.match(/\s*$/)[0];
        const core=raw.trim();
        const tr=this.t(core);
        if(tr!==core)node.nodeValue=lead+tr+trail;
      });
      root.querySelectorAll('[title],[placeholder]').forEach(el=>{
        if(this.isSkipNode(el))return;
        if(el.title&&/[\u4e00-\u9fa5]/.test(el.title)){
          const tr=this.t(el.title);
          if(tr!==el.title)el.title=tr;
        }
        if(el.placeholder&&/[\u4e00-\u9fa5]/.test(el.placeholder)){
          const tr=this.t(el.placeholder);
          if(tr!==el.placeholder)el.placeholder=tr;
        }
      });
    },
    apply(){
      this.setDir();
      this.translateTree(document.body);
    },
    set(lang){
      this.lang=lang;
      try{localStorage.setItem('officepro-lang',lang);}catch(e){}
      location.reload();
    },
    init(){
      (window.__DICTS||[]).forEach(pair=>this.register(pair[0],pair[1]));
      let saved='zh';
      try{saved=localStorage.getItem('officepro-lang')||'zh';}catch(e){}
      this.lang=saved;
      if(saved!=='zh'){
        this.setDir();
        const run=()=>this.apply();
        if(document.readyState==='complete')setTimeout(run,80);
        else window.addEventListener('load',()=>setTimeout(run,80));
      }
    }
  };
  window.I18N=I18N;
  window.LANG_LIST=[
    {code:'zh',label:'中文'},
    {code:'en',label:'English'},
    {code:'ja',label:'日本語'},
    {code:'ko',label:'한국어'},
    {code:'es',label:'Español'},
    {code:'fr',label:'Français'},
    {code:'de',label:'Deutsch'},
    {code:'ru',label:'Русский'},
    {code:'pt',label:'Português'},
    {code:'ar',label:'العربية'}
  ];
  window.setLanguage=function(){
    const list=window.LANG_LIST||[];
    const names=list.map((l,i)=>(i+1)+'='+l.label).join('  ');
    const v=prompt('Language / 语言（输入编号）：\n'+names,'1');
    if(!v)return;
    const idx=parseInt(v,10)-1;
    if(list[idx])I18N.set(list[idx].code);
  };
  window.slideLanguage=window.setLanguage;
  I18N.init();
})();
