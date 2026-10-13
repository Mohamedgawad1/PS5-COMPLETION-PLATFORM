"""
patch_index.py - one-shot refactor of index.html (line-addressed).

Run split_data.py first: this removes the 6.3 MB of inline datasets, adds the
lazy loader and the pager, and repoints the long tables at one page of rows at a
time. Refuses to run twice.
"""
import io
import os
import shutil
import sys

P = os.path.dirname(os.path.abspath(__file__))
F = os.path.join(P, 'index.html')

lines = io.open(F, encoding='utf-8').read().split('\n')

# (line_no, must_start_with, replacement lines)
E = []


def ed(ln, starts, *new):
    E.append((ln, starts, list(new)))


# ----------------------------------------------------------------- 1369 gsearch
ed(1369, 'function gsearch(q){',
    "var GKEYS=['punch','itr','cables','wire','cbl','cs','fg','io','trm'];\nfunction gsearch(q){")
ed(1370, ' q=String(q||\'\').trim();', """ q=String(q||'').trim();
 var box=$('gres');
 // the search covers all 16 pages, so it needs every data file
 var gmiss=[];
 for(var gi=0;gi<GKEYS.length;gi++)if(!DOK[GKEYS[gi]])gmiss.push(GKEYS[gi]);
 if(gmiss.length){
  $('gcount').textContent='loading data...';
  loadData(gmiss,function(){gsearch(q)});
  return;
 }""")

# ------------------------------------------------------------------- 1119 CSS
ed(1119, '#xlsbtn,#prnbtn,#svbtn{width:32px;', """.pgr{display:inline-flex;align-items:center;gap:6px;margin-left:10px;vertical-align:middle}
.pgr button{font-family:inherit;font-size:9.5px;font-weight:800;letter-spacing:.4px;color:#0b2239;background:#e3edf8;border:1px solid #c3d6ea;border-radius:999px;padding:3px 10px;cursor:pointer}
.pgr button:hover:not([disabled]){background:#d3e4f5}
.pgr button[disabled]{opacity:.4;cursor:default}
.pgr b{font-size:10px;color:#0b2239;min-width:52px;text-align:center;display:inline-block}
#dload{position:fixed;left:0;right:0;top:0;height:3px;background:linear-gradient(90deg,var(--acc),var(--gold));z-index:9999;opacity:0;transition:opacity .18s;pointer-events:none;box-shadow:0 0 10px rgba(237,125,49,.8)}
#xlsbtn,#prnbtn,#svbtn{width:32px;""")

# ---------------------------------------------------------------- 1099 s-fg opts
ed(1099, '(function(){const s=$(\'s-fg\');',
    """function fillFgOpts(){const s=$('s-fg');if(s.dataset.d0===undefined)s.dataset.d0=s.options[0]?s.options[0].outerHTML:'';s.innerHTML=s.dataset.d0;const v={};FG.forEach(r=>{if(r[0]&&!v[r[0]]){v[r[0]]=1;const o=document.createElement('option');o.value=r[0];o.textContent=r[0]+'  -  '+r[4];s.appendChild(o);}});}DHOOK.fg=fillFgOpts;""")

# ------------------------------------------------------------ 1088-1097 io/trm
ed(1088, '(function(){', 'function fillIoTrmOpts(){')
ed(1089, " var mi=[['s-io'", " var mi=[['s-io'")
ed(1090, ' [[mi,IO],[mt,TRM]]', '  [[mi,IO],[mt,TRM]]')
ed(1091, '  var sel=document.getElementById',
    '  var sel=document.getElementById')
ed(1092, '  c[1].forEach(', '  if(sel.dataset.d0===undefined)sel.dataset.d0=sel.options[0]?sel.options[0].outerHTML:\'\';')
ed(1093, '   if(k&&!seen[k]){',
    '  sel.innerHTML=sel.dataset.d0;')
ed(1094, ' });', '  c[1].forEach(function(r){String(fn(r)).split(\' \').forEach(function(k){')
ed(1095, " $('q-io').addEventListener",
    '   if(k&&!seen[k]){seen[k]=1;var o=document.createElement(\'option\');o.value=k;o.textContent=k;sel.appendChild(o);}});});')
ed(1096, " $('q-trm').addEventListener", '  });}')
ed(1097, '})();',
    "DHOOK.io=fillIoTrmOpts;DHOOK.trm=fillIoTrmOpts;\n$('q-io').addEventListener('input',renderIo);$('s-io').addEventListener('change',renderIo);\n$('q-trm').addEventListener('input',renderTrm);$('s-trm').addEventListener('change',renderTrm);")

# ---------------------------------------------------------------- 1086 s-cs opts
ed(1086, '(function(){const s=$(\'s-cs\');',
    """function fillCsOpts(){const s=$('s-cs');if(s.dataset.d0===undefined)s.dataset.d0=s.options[0]?s.options[0].outerHTML:'';s.innerHTML=s.dataset.d0;const v={};CS.forEach(r=>{[r[5],r[9]].forEach(k=>{if(k&&!v[k]){v[k]=1;const o=document.createElement('option');o.value=k;o.textContent=k;s.appendChild(o);}});});}DHOOK.cs=fillCsOpts;""")

# --------------------------------------------------------------- 1084 s-cbl opts
ed(1084, '(function(){const s=$(\'s-cbl\');',
    """function fillCblOpts(){const s=$('s-cbl');if(s.dataset.d0===undefined)s.dataset.d0=s.options[0]?s.options[0].outerHTML:'';s.innerHTML=s.dataset.d0;const v={};CBL.forEach(r=>{if(r[5]&&!v[r[5]]){v[r[5]]=1;const o=document.createElement('option');o.value=r[5];o.textContent=r[4]||r[5];s.appendChild(o);}});}DHOOK.cbl=fillCblOpts;""")

# ------------------------------------------------------------ 1082 s-layout opts
ed(1082, '(function(){const s=$(\'s-layout\');',
    """function fillLayoutOpts(){const s=$('s-layout');if(s.dataset.d0===undefined)s.dataset.d0=s.options[0]?s.options[0].outerHTML:'';s.innerHTML=s.dataset.d0;const seen={};WIRE.forEach(r=>{const k=r[8]||'-';if(!seen[k]){seen[k]=1;const o=document.createElement('option');o.value=k;o.textContent=k;s.appendChild(o);}});}DHOOK.wire=fillLayoutOpts;""")

# -------------------------------------------------------------- 1010 checkbox cb
ed(1010, ' if(RENDER[V])RENDER[V]();', ' renderView(V);')

# ---------------------------------------------------------------- 1001 st chips
ed(1001, " ch.classList.add('on');F.stt=ch.dataset.st;renderList('it');});",
    " ch.classList.add('on');F.stt=ch.dataset.st;renderList('it',1);});")

# ---------------------------------------------------------------- 998 cat chips
ed(998, " ch.classList.add('on');F.cat=ch.dataset.cat;renderList('pu');});",
    " ch.classList.add('on');F.cat=ch.dataset.cat;renderList('pu',1);});")

# --------------------------------------------------------------- 993-992 filters
ed(993, "$('s-pu').addEventListener('change',()=>renderList('pu'));",
    "$('s-pu').addEventListener('change',()=>renderList('pu',1));")
ed(992, "[['q-rfc',renderRfc]",
    "[['q-rfc',renderRfc],['q-blk',renderBlk],['q-pu',()=>renderList('pu',1)],['q-it',()=>renderList('it',1)]].forEach(([id,fn])=>$(id).addEventListener('input',fn));\n"
    "$('s-it').addEventListener('change',()=>renderList('it',1));")

# -------------------------------------------------------------- 988-984 show(v)
ed(988, ' if(RENDER[v])RENDER[v]();',
    """ var need=NEED[v]||['core'],miss=[];
 for(var i=0;i<need.length;i++)if(!DOK[need[i]])miss.push(need[i]);
 if(miss.length){
  busy(v,true);
  loadData(miss,function(){if(V===v)renderView(v);busy(v,false);});
 }else renderView(v);""")
ed(984, 'function show(v){', 'function renderView(v){if(RENDER[v])try{RENDER[v]()}catch(e){}}\nfunction show(v){')

# ------------------------------------------------------------------- 956 h-cs
ed(956, " $('h-cs').textContent=",
    " pgHint('h-cs','cs',CSPG,CS.length+' cables'+(sv?' | '+sv:'')+(q?' | \"'+q+'\"':''));")

# ------------------------------------------------------------ 942 renderCs rows
ed(941, " let o='';", " let o='';const CSPG=pgSlice('cs',f);")
ed(942, ' f.forEach((x,i)=>{o+=',
    ' CSPG.view.forEach((x,i)=>{o+=\'<tr><td style="text-align:center">\'+(CSPG.p*CSPG.size+i+1)+\'</td>\'')

# ------------------------------------------------------------------ 933 h-trm
ed(933, " $('h-trm').textContent=",
    " pgHint('h-trm','trm',TPG,TRM.length+' termination sheets / '+new Set(f.reduce(function(a,x){return a.concat(x[3].split(' '));},[]).filter(Boolean)).size+' subsystems'+(sv?' | '+sv:'')+(q?' | \"'+q+'\"':''));")

# ------------------------------------------------------------ 927 renderTrm rows
ed(926, " let o='';", " let o='';const TPG=pgSlice('trm',f);")
ed(927, ' f.forEach((x,i)=>{o+=',
    ' TPG.view.forEach((x,i)=>{o+=\'<tr><td style="text-align:center">\'+(TPG.p*TPG.size+i+1)+\'</td>\'')

# ------------------------------------------------------------------ 918 h-io
ed(918, " $('h-io').textContent=",
    " pgHint('h-io','io',IPG,IO.length+' instrument tags / '+new Set(f.map(function(x){return x[0];}).filter(Boolean)).size+' subsystems'+(sv?' | '+sv:'')+(q?' | \"'+q+'\"':''));")

# ------------------------------------------------------------ 911 renderIo rows
ed(910, " let o='';", " let o='';const IPG=pgSlice('io',f);")
ed(911, ' f.forEach((x,i)=>{o+=',
    ' IPG.view.forEach((x,i)=>{o+=\'<tr><td style="text-align:center">\'+(IPG.p*IPG.size+i+1)+\'</td>\'')

# ------------------------------------------------------------------ 902 h-fg
ed(902, " $('h-fg').textContent=FG.length",
    " pgHint('h-fg','fg',DPG,FG.length+' devices / '+new Set(FG.filter(x=>x[7]).map(x=>x[7])).size+' subsystems / '+FZ.length+' fire zones'+(sv&&sv!=='__ZONES'?' | '+sv:'')+(q?' | \"'+q+'\"':''));")

# ------------------------------------------------------------ 893 renderFg dev
ed(892, " let o='';", " let o='';const DPG=pgSlice('fg',f);")
ed(893, ' f.forEach((x,i)=>{o+=',
    ' DPG.view.forEach((x,i)=>{o+=\'<tr><td style="text-align:center">\'+(DPG.p*DPG.size+i+1)+\'</td>\'')

# ------------------------------------------------------------------ 885 h-fg z
ed(885, "  $('h-fg').textContent=FZ.length",
    " pgHint('h-fg','fg',ZPG,FZ.length+' fire zones'+(q?' | \"'+q+'\"':''));")

# ------------------------------------------------------------ 883 renderFg z
ed(882, "  let o='';", "  let o='';const ZPG=pgSlice('fg',f);")
ed(883, '  f.forEach((x,i)=>{o+=',
    '  ZPG.view.forEach((x,i)=>{o+=\'<tr><td style="text-align:center">\'+(ZPG.p*ZPG.size+i+1)+\'</td>\'')

# ------------------------------------------------------------------ 874 h-cbl
ed(874, " $('h-cbl').textContent=",
    " pgHint('h-cbl','cbl',PPG,CBL.length+' cables / '+CBL.reduce(function(a,c){return a+c[6];},0)+' cores'+(sf?' | '+sf:'')+(q?' | \"'+q+'\"':''));")

# ------------------------------------------------------------ 863 renderCbl rows
ed(862, " let o='';", " let o='';const PPG=pgSlice('cbl',f);")
ed(862, " let o='';", " let o='';const PPG=pgSlice('cbl',f);")
ed(863, ' f.forEach((x,i)=>{o+=',
    ' PPG.view.forEach((x,i)=>{o+=\'<tr><td style="text-align:center">\'+(PPG.p*PPG.size+i+1)+\'</td>\'')

# ----------------------------------------------------------------- 854 h-layout
ed(854, " $('h-layout').textContent=",
    " pgHint('h-layout','layout',LPG,L.length+' layout sheet(s) from '+src.length+' terminations | '+new Set(f.map(x=>x.sid).filter(Boolean)).size+' subsystem(s) | '+new Set(f.map(x=>x.task).filter(Boolean)).size+' task id(s)'+(rf?' | reference: '+rf:'')+(q?' | \"'+q+'\"':''));")

# ---------------------------------------------------------- 840 renderLayout rows
ed(839, " let out='';", " let out='';const LPG=pgSlice('layout',L);")
ed(839, " let out='';", " let out='';const LPG=pgSlice('layout',L);")
ed(840, ' L.forEach((g,i)=>{out+=',
    ' LPG.view.forEach((g,i)=>{out+=\'<tr><td style="text-align:center">\'+(LPG.p*LPG.size+i+1)+\'</td>\'')

# ------------------------------------------------------------------ 823 h-wire
ed(823, " $('h-wire').textContent=",
    " pgHint('h-wire','wire',PG,d.length+' wiring terminations'+(sid?' | subsystem: '+sid:'')+((d.length)&&!q?' | '+new Set(d.map(r=>r[0])).size+' cables':''));")

# ---------------------------------------------------------- 806-807 renderWiring
ed(807, "  rows+='<tr><td style=\"text-align:center\">'+(i+1)+'</td>';",
    "  rows+='<tr><td style=\"text-align:center\">'+(PG.p*PG.size+i+1)+'</td>';")
ed(806, ' view.forEach((r,i)=>{', ' PG.view.forEach((r,i)=>{')

# ------------------------------------------------------------------- 803 wiring
ed(803, '  const LIMIT=50000', '  const PG=pgSlice(\'wire\',d),view=PG.view;')

# ------------------------------------------------------------------- 793 h-eit9
ed(793, " $('h-eit9').textContent=",
    " pgHint('h-eit9','eit9',EPG,open.length+' open task(s) out of '+I.length+' total E/I/T (CPP AGI)');")

# ------------------------------------------------------------ 782-783 renderEit9
ed(783, "  rows+='<tr><td style=\"text-align:center\">'+(i+1)+'</td>';",
    "  rows+='<tr><td style=\"text-align:center\">'+(EPG.p*EPG.size+i+1)+'</td>';")
ed(782, ' open.forEach((r,i)=>{', ' const EPG=pgSlice(\'eit9\',open);\n EPG.view.forEach((r,i)=>{')

# ------------------------------------------------------------------- 756 h-cab
ed(756, " $('h-cab').textContent=",
    " pgHint('h-cab','cab',PG,d.length+' cables'+(sid?' | subsystem: '+sid:'')+(disc?' | discipline: '+disc:''));")

# ------------------------------------------------------------------- 738 cables
ed(738, ' const LIMIT=2000', ' const PG=pgSlice(\'cab\',d),view=PG.view;')

# --------------------------------------------------------------- 726 showCables
ed(726, " show('cables');renderCables();", " show('cables');")

# ------------------------------------------------------------------- 643 list
ed(643, ' const LIMIT=2000,trunc=d.length>LIMIT;const view=trunc?d.slice(0,LIMIT):d;',
    ' const PG=pgSlice(PKG,d),view=PG.view;')

# ------------------------------------------------------------------- 629 list sig
ed(629, 'function renderList(w){',
    'function renderList(w,rs){\n const PKG=(w===\'pu\')?\'pu\':\'it\';\n if(rs)pgReset(PKG);')

# ------------------------------------------------------------------- 555-556 sub
ed(556, " $('h-si').textContent=",
    " pgHint('h-si','si',pgSlice('si',I),I.length+' itr task(s) - all cells editable & synced to Excel');")
ed(555, " $('h-sp').textContent=",
    " pgHint('h-sp','sp',pgSlice('sp',P),P.length+' punch item(s)');")

# ------------------------------------------------------------------- 484 subRows
ed(484, ' for(const r of src){', " const PG=pgSlice(isP?'sp':'si',src);\n for(const r of PG.view){")

# ------------------------------------------------------------------- 477 goSub
ed(477, " show('sub');renderSub();", " show('sub');")

# ------------------------------------------------------------------- 442/441 dash
ed(442, ' let puT=PUNCH.length,',
    " let puT=KCNT.puT||PUNCH.length,puCl=KCNT.puCl||PUNCH.filter(r=>String(r[5]).trim()==='Closed').length,puOp=puT-puCl,puPct=puT?Math.round(puCl/puT*100):0;")
ed(441, ' let itrT=ITR.length,',
    " let itrT=KCNT.itrT||ITR.length,itrCl=KCNT.itrCl||ITR.filter(r=>doneI(r[5])).length,itrOp=itrT-itrCl,itrPct=itrT?Math.round(itrCl/itrT*100):0;")

# ------------------------------------------------------------------- 368 LIVEVER
ed(368, "const LIVEVER='12';", "const LIVEVER='13';")

# ------------------------------------------------------------ 159 loading bar
ed(159, '<div class="flow">', '<div id="dload"></div>\n<div class="flow">')

# ---------------------------------------------------------------- data lines
PREAMBLE = r"""/* ===== DATA LAYER =====
   The datasets used to be one 6.3 MB JavaScript literal inlined here, so every
   page load downloaded and parsed all of it before anything appeared. They now
   live in d_<name>.js files (see split_data.py): d_core.js loads before this
   script, the rest are pulled in the first time the page that needs them is
   opened. Declared here so no reference can throw before a file arrives. */
var SUBS=[],PUNCH=[],ITR=[],RFC=[],ITRT=[],PUNT=[],MILES=[],CABLES=[],RFCK=[],DPRSUM={},
    FG=[],FZ=[],IO=[],TRM=[],CS=[],CBL=[],WIRE=[],
    NOTES={},PCOL={'PUNCH LIST':{},'ITR LIST':{},'RFC PROGRESS':{}},
    OVR={'RFC PROGRESS':{},'PUNCH LIST':{},'ITR LIST':{}},
    KCNT={itrT:0,itrCl:0,puT:0,puCl:0};
/* mirror of ps5data.unpack_rows - rebuilds the original arrays from a packed
   column store with per-column dictionaries */
var PS5UNPACK=function(p){
 if(!p||!p.R)return p;
 var R=p.R,K=p.k,D=p.D,n=R.length?R[0].length:0,c,i,col,d,ki,row,o=new Array(n);
 for(c=0;c<R.length;c++){col=R[c];ki=K.indexOf(c);if(ki<0)continue;d=D[ki];for(i=0;i<n;i++)col[i]=d[col[i]];}
 for(i=0;i<n;i++){row=new Array(R.length);for(c=0;c<R.length;c++)row[c]=R[c][i];o[i]=row;}
 return o;
};
var DFILES={core:'d_core.js',punch:'d_punch.js',itr:'d_itr.js',cables:'d_cables.js',wire:'d_wire.js',cbl:'d_cbl.js',cs:'d_cs.js',fg:'d_fg.js',io:'d_io.js',trm:'d_trm.js'};
var NEED={dash:['core'],rfc:['core'],blk:['core'],done:['core'],pu:['core','punch'],it:['core','itr'],
 sub:['core','punch','itr','cables'],cables:['core','cables'],eit9:['core','itr'],
 wire:['core','wire'],layout:['core','wire'],cbl:['cbl'],cs:['cs'],fg:['fg'],io:['io'],trm:['trm']};
var DOK={core:1},DPEND={},DHOOK={};
function loadData(keys,done){
 var wait=[];
 keys.forEach(function(k){
  if(DOK[k])return;
  if(DPEND[k]){wait.push(DPEND[k]);return;}
  DPEND[k]=new Promise(function(res){
   var s=document.createElement('script');
   s.async=false;s.src=DFILES[k];
   s.onload=function(){DOK[k]=1;DPEND[k]=null;try{if(DHOOK[k])DHOOK[k]();}catch(e){}res();};
   s.onerror=function(){DPEND[k]=null;res();};
   document.head.appendChild(s);
  });
  wait.push(DPEND[k]);
 });
 if(!wait.length){if(done)done();return}
 Promise.all(wait).then(function(){if(done)done()},function(){if(done)done()});
}
/* ===== PAGINATION =====
   Every table used to be written out in one go - up to 2,000 rows x 10 editable
   cells, or 4,073 wiring rows: 146,795 DOM nodes on the wiring page, which took
   seconds of layout and made the next click unresponsive. Only one page of rows
   is in the DOM at a time now. */
var PSIZE=150,PST={};
function pgSlice(key,rows,size){
 var sz=size||PSIZE,pages=Math.max(1,Math.ceil(rows.length/sz)),p=PST[key]||0;
 if(p>pages-1)p=pages-1;
 if(p<0)p=0;
 PST[key]=p;
 return {view:rows.slice(p*sz,(p+1)*sz),p:p,pages:pages,size:sz,total:rows.length};
}
function pgReset(key){PST[key]=0;}
function pgHint(id,key,pg,text){
 var el=$(id);
 if(!el)return;
 var bar=pg.pages>1?' <span class="pgr"><button type="button" data-pg="'+key+'" data-d="-1"'+(pg.p<=0?' disabled':'')+'>&laquo; PREV</button><b>'+(pg.p+1)+' / '+pg.pages+'</b><button type="button" data-pg="'+key+'" data-d="1"'+(pg.p>=pg.pages-1?' disabled':'')+'>NEXT &raquo;</button></span>':'';
 el.innerHTML=esc(text)+bar;
}
document.addEventListener('click',function(e){
 var b=e.target.closest('[data-pg]');
 if(!b||b.disabled)return;
 var k=b.dataset.pg,d=+b.dataset.d||0;
 PST[k]=Math.max(0,(PST[k]||0)+d);
 renderView(V);
});
function busy(v,on){
 var el=document.getElementById('dload');
 if(!el)return;
 el.style.opacity=on?'1':'0';
}
</script>
<script src="d_core.js"></script>
<script>"""

if os.environ.get('PS5_ALLOW_SPLIT') != '1':
    sys.exit(
        "REFUSING TO RUN.\n"
        "patch_index.py strips the inline datasets out of index.html and moves\n"
        "them into d_*.js. rebuild_data.py and make_platform_excel.py then fail\n"
        "with `cannot find const SUBS=/const NOTES=` and the whole auto-sync\n"
        "chain dies every cycle, so the live platform stops updating.\n"
        "This step is only valid as a one-off manual optimisation, never as part\n"
        "of the sync loop. If you really want it, set PS5_ALLOW_SPLIT=1 and\n"
        "rebuild the parsers for the split layout first.")

ed(338, 'const SUBS=', PREAMBLE)
for ln in (339, 678, 679, 680, 681, 682, 683, 684, 685):
    E.append((ln, None, ['']))  # blank out the rest of the data lines

# --------------------------------------------------------------------- apply
bad = []
for ln, starts, new in E:
    cur = lines[ln - 1]
    if starts and not cur.startswith(starts):
        bad.append((ln, starts, cur[:90]))
if bad:
    print('ANCHOR MISMATCH:')
    for ln, s, c in bad:
        print('  L%-5d want %r\n         got  %r' % (ln, s[:60], c))
    sys.exit(1)

for ln, starts, new in sorted(E, key=lambda t: -t[0]):
    lines[ln - 1:ln] = new

out = '\n'.join(lines)
bak = F + '.before-split'
if not os.path.exists(bak):
    shutil.copy2(F, bak)
    print('backup ->', bak)
io.open(F, 'w', encoding='utf-8', newline='\n').write(out)
print('index.html: %d bytes, %d lines (was 7,639,052 / 1,472)' % (len(out), len(lines)))
print('applied %d edits' % len(E))
