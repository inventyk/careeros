import React,{useEffect,useMemo,useRef,useState} from 'react';
import * as XLSX from 'xlsx';

const STORAGE_KEY='careeros-user-tracker-v1';
async function loadBundledJobs(){
 const response=await fetch('/jobs-data.json.gz.b64',{cache:'no-store'});
 if(!response.ok)throw new Error('Could not load the GitHub job dataset');
 const encoded=(await response.text()).trim();
 const binary=atob(encoded);
 const bytes=Uint8Array.from(binary,ch=>ch.charCodeAt(0));
 const stream=new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
 const workbookBytes=new Uint8Array(await new Response(stream).arrayBuffer());
 const wb=XLSX.read(workbookBytes,{type:'array',cellDates:true});
 return ['Individual Jobs','Applied','not-applied'].filter(n=>wb.Sheets[n]).flatMap(n=>normalizeSheet(wb.Sheets[n],n));
}
const norm = value => String(value ?? '').toLowerCase().trim().replace(/[^a-z0-9]+/g,' ').replace(/\s+/g,' ').trim();
const FIELD_ALIASES = {
  company: ['company'], title: ['job title'], location: ['location'],
  mode: ['work mode type','work mode','work type'], freshness: ['freshness status'],
  salary: ['salary stipend','salary'], skills: ['required skills'],
  url: ['direct apply url'], serial: ['s no','serial no']
};
const headerScore = headers => headers.reduce((score,h) => score + (Object.values(FIELD_ALIASES).some(list=>list.includes(norm(h)))?1:0),0);
function findColumn(headers,field){ return headers.findIndex(h=>FIELD_ALIASES[field].includes(norm(h))); }
function normalizeSheet(sheet, sourceTab){
  const matrix=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false,blankrows:false});
  if(!matrix.length)return [];
  let headerIndex=0,best=-1;
  for(let i=0;i<Math.min(matrix.length,10);i++){const score=headerScore(matrix[i]||[]);if(score>best){best=score;headerIndex=i;}}
  const headers=(matrix[headerIndex]||[]).map(h=>String(h??'').trim());
  const columns=Object.fromEntries(Object.keys(FIELD_ALIASES).map(k=>[k,findColumn(headers,k)]));
  return matrix.slice(headerIndex+1).map((cells,i)=>{
    const get=field=>{const idx=columns[field];return idx>=0?String(cells[idx]??'').trim():'';};
    const company=get('company'),title=get('title'),location=get('location'),mode=get('mode');
    const freshness=get('freshness'),salary=get('salary'),skills=get('skills'),url=get('url');
    if(!company&&!title&&!url)return null;
    // Deliberately whitelist fields: AI Agent Task and all agent-instruction columns are never read or exported.
    const raw={'Company':company,'Job Title':title,'Location':location,'Work Mode / Type':mode,'Salary / Stipend':salary,'Freshness / Status':freshness,'Required Skills':skills,'Direct Apply URL':url};
    const applicationStatus=sourceTab==='Applied'?'Applied':sourceTab==='not-applied'?'Not Applied':'Individual';
    return {id:sourceTab+'-'+(i+1),raw,company:company||'—',title:title||'Untitled role',location:location||'—',mode:mode||'—',status:applicationStatus,freshness:freshness||'—',salary:salary||'Not specified',skills,url,sourceTab};
  }).filter(Boolean);
}
const isApplied = job => job.sourceTab==='Applied';
function App(){
 const [jobs,setJobs]=useState([]),[fileName,setFileName]=useState('GitHub job tracker'),[sheetNames,setSheetNames]=useState([]),[sheet,setSheet]=useState(''),[workbook,setWorkbook]=useState(null),[tab,setTab]=useState('All Jobs'),[query,setQuery]=useState(''),[mode,setMode]=useState('All work modes'),[sort,setSort]=useState('company'),[notice,setNotice]=useState('Upload your Excel or CSV tracker to load your own jobs.');
 const input=useRef(null);
 useEffect(()=>{let active=true;try{const saved=localStorage.getItem(STORAGE_KEY);if(saved){const rows=JSON.parse(saved);if(Array.isArray(rows)&&rows.length){setJobs(rows);setFileName('Saved uploaded tracker');setNotice('Loaded '+rows.length+' saved jobs. This version stays active until you upload another workbook.');return()=>{active=false;};}}}catch(e){localStorage.removeItem(STORAGE_KEY);}loadBundledJobs().then(rows=>{if(active){setJobs(rows);setNotice('Automatically loaded '+rows.length+' jobs from the GitHub tracker.');}}).catch(()=>{if(active)setNotice('Could not load the GitHub tracker. Upload a workbook to set your saved version.');});return()=>{active=false;};},[]);
 const TRACKER_SHEETS=['Individual Jobs','Applied','not-applied'];
 const loadWorkbook=wb=>{const selected=TRACKER_SHEETS.filter(n=>wb.Sheets[n]);const all=selected.flatMap(n=>normalizeSheet(wb.Sheets[n],n));setSheetNames(selected);setJobs(all);setSheet('');setFileName('Saved uploaded tracker');try{localStorage.setItem(STORAGE_KEY,JSON.stringify(all));setNotice('Saved '+all.length+' rows in this browser. This dataset will remain active until the next upload.');}catch(e){setNotice('Loaded '+all.length+' rows, but browser storage is full; this upload may not persist after closing the browser.');}};
 const onFile=async e=>{const f=e.target.files?.[0];if(!f)return;try{const wb=XLSX.read(await f.arrayBuffer(),{type:'array',cellDates:true});setWorkbook(wb);setFileName(f.name);loadWorkbook(wb);setTab('All Jobs');setQuery('');}catch(err){setNotice('Could not read this file. Upload a valid .xlsx, .xls or .csv file.');}e.target.value='';};
 const applied=jobs.filter(j=>j.sourceTab==='Applied').length;
 const individualCount=jobs.filter(j=>j.sourceTab==='Individual Jobs').length;
 const notApplied=jobs.filter(j=>j.sourceTab==='not-applied').length;
 const filtered=useMemo(()=>jobs.filter(j=>{const matchTab=tab==='All Jobs'||(tab==='Applied'&&j.sourceTab==='Applied')||(tab==='Not Applied'&&j.sourceTab==='not-applied')||(tab==='Individual'&&j.sourceTab==='Individual Jobs');const hay=[j.company,j.title,j.location,j.mode,j.status,j.freshness,j.salary,j.skills].join(' ').toLowerCase();return matchTab&&hay.includes(query.toLowerCase())&&(mode==='All work modes'||String(j.mode).toLowerCase().includes(mode.toLowerCase()));}).sort((a,b)=>String(a[sort]||'').localeCompare(String(b[sort]||''))),[jobs,tab,query,mode,sort]);
 const modes=[...new Set(jobs.map(j=>j.mode).filter(x=>x&&x!=='—'))];
 const exportFile=()=>{const ws=XLSX.utils.json_to_sheet(jobs.map(j=>j.raw));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Jobs');XLSX.writeFile(wb,'CareerOS-export.xlsx');};
 return <div className="app-shell"><aside className="sidebar"><div className="brand"><div className="brand-mark">C</div><div><b>CareerOS</b><small>CAREER COMMAND CENTER</small></div></div><div className="side-label">WORKSPACE</div>{['All Jobs','Applied','Not Applied','Individual'].map((t,i)=><button key={t} className={'nav-item '+(tab===t?'active':'')} onClick={()=>setTab(t)}><span>{['▦','✓','◷','◇'][i]}</span>{t}</button>)}<div className="side-bottom"><div className="profile-dot">M</div><div><b>My workspace</b><small>Personal job tracker</small></div></div></aside>
 <main className="main"><header className="topbar"><div><div className="eyebrow">YOUR CAREER, IN CONTROL</div><h1>Job application <span>overview</span></h1></div><div className="top-actions"><button className="ghost-btn" onClick={exportFile} disabled={!jobs.length}>↓ Export</button><button className="primary-btn" onClick={()=>input.current?.click()}>＋ Update Data</button><input ref={input} type="file" accept=".xlsx,.xls,.csv" hidden onChange={onFile}/></div></header>
 <section className="welcome"><div><div className="eyebrow light">CAREEROS DASHBOARD</div><h2>Make your next move<br/><em>count.</em></h2><p>One clear view of every opportunity, application and next step.</p></div><div className="welcome-orb"><div className="orb-ring ring1"/><div className="orb-ring ring2"/><div className="orb-core">✦</div></div><div className="welcome-foot">{fileName?'SOURCE · '+fileName:'READY FOR YOUR TRACKER'} <span>● LIVE IN THIS SESSION</span></div></section>
 <section className="stats"><Stat label="TOTAL OPPORTUNITIES" value={jobs.length} note="Rows loaded from GitHub tracker" icon="▦"/><Stat label="APPLICATIONS SENT" value={applied} note="Applied / submitted status" icon="↗"/><Stat label="YET TO APPLY" value={notApplied} note="Remaining opportunities" icon="◷"/><Stat label="APPLICATION RATE" value={(applied+notApplied)?Math.round(applied/(applied+notApplied)*100)+'%':'—'} note="Applied ÷ (Applied + Not Applied)" icon="◉"/></section>
 <section className="content-card"><div className="section-head"><div><div className="eyebrow">OPPORTUNITY PIPELINE</div><h3>{tab}</h3><p>{jobs.length?filtered.length+' matching opportunities from your tracker':'Your GitHub tracker data will appear here'}</p></div><div className="head-actions"><select value={mode} onChange={e=>setMode(e.target.value)}><option>All work modes</option>{modes.map(m=><option key={m}>{m}</option>)}</select><button className="ghost-btn" onClick={()=>input.current?.click()}>↑ Upload file</button></div></div>
 <div className="tabs">{['All Jobs','Applied','Not Applied','Individual'].map(t=><button key={t} className={tab===t?'selected':''} onClick={()=>setTab(t)}>{t}<span>{t==='All Jobs'?jobs.length:t==='Applied'?applied:t==='Not Applied'?notApplied:individualCount}</span></button>)}</div>
 <div className="table-tools"><div className="search"><span>⌕</span><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search company, role, location, skills..."/></div><label className="sort-label">SORT BY <select value={sort} onChange={e=>setSort(e.target.value)}><option value="company">Company</option><option value="title">Job title</option><option value="location">Location</option><option value="status">Status</option></select></label></div>
 <div className="table-wrap"><table><thead><tr><th>COMPANY / ROLE</th><th>LOCATION</th><th>WORK MODE</th><th>SALARY</th><th>STATUS</th><th>LINK</th></tr></thead><tbody>{filtered.map(j=><tr key={j.id}><td><div className="company-cell"><div className="company-logo">{j.company.slice(0,1).toUpperCase()}</div><div><b>{j.company}</b><small>{j.title}</small></div></div></td><td>{j.location}</td><td><span className="mode-pill">{j.mode}</span></td><td>{j.salary}</td><td><span className={'status '+(isApplied(j)?'done':'pending')}>{j.status}</span></td><td>{j.url?<a className="apply-link" href={j.url} target="_blank" rel="noreferrer">Open ↗</a>:'—'}</td></tr>)}{!filtered.length&&<tr><td colSpan="6" className="empty"><div className="empty-icon">⇧</div><b>{jobs.length?'No matching jobs':'No tracker data available'}</b><p>{jobs.length?'Try changing your search or filters.':'The saved GitHub tracker could not be loaded. You can upload a file as a fallback.'}</p><button className="primary-btn" onClick={()=>input.current?.click()}>Choose spreadsheet</button></td></tr>}</tbody></table></div>
 <div className="table-footer"><span>Showing {filtered.length} of {jobs.length} opportunities</span><span>Default data loads automatically from this repository</span></div></section>
 <footer>CAREEROS <span>BUILT FOR YOUR NEXT OPPORTUNITY</span><span>{notice}</span></footer></main></div>
}
function Stat({label,value,note,icon}){return <div className="stat-card"><div className="stat-top"><span>{label}</span><i>{icon}</i></div><strong>{value}</strong><small>{note}</small></div>}
export default App;
