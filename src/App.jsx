import React,{useMemo,useRef,useState} from 'react';
import * as XLSX from 'xlsx';

const norm = value => String(value ?? '').toLowerCase().trim().replace(/[^a-z0-9]+/g,' ').replace(/\s+/g,' ').trim();
const FIELD_ALIASES = {
  company: ['company','company name','employer','organization','organisation','firm','business name','employer name','hiring company'],
  title: ['job title','title','position','role','job role','designation','opening','profile','job profile','vacancy'],
  location: ['location','job location','city','work location','place','office location','job city','location city'],
  mode: ['work mode','work type','remote onsite','remote or onsite','employment type','job type','workplace type','working mode','work arrangement'],
  status: ['status','application status','applied status','application stage','applied','application','progress'],
  salary: ['salary','ctc','compensation','pay','salary range','package','annual salary','stipend'],
  skills: ['skills','required skills','key skills','technical skills','skill set','technologies'],
  url: ['apply url','apply link','application url','job url','link','source url','source link','job link','application link','url'],
  date: ['date','posted date','job date','application date','last date','deadline']
};
const headerScore = headers => headers.reduce((score,h) => score + (Object.values(FIELD_ALIASES).some(list=>list.some(a=>norm(h)===norm(a)||norm(h).includes(norm(a))))?1:0),0);
function findColumn(headers,field){
  const aliases=FIELD_ALIASES[field].map(norm);
  let index=headers.findIndex(h=>aliases.includes(norm(h)));
  if(index>=0)return index;
  index=headers.findIndex(h=>{const n=norm(h);return aliases.some(a=>a.length>3&&(n.includes(a)||a.includes(n)));});
  return index;
}
function normalizeSheet(sheet){
  const matrix=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false,blankrows:false});
  if(!matrix.length)return [];
  let headerIndex=0,best=0;
  for(let i=0;i<Math.min(matrix.length,15);i++){const score=headerScore(matrix[i]||[]);if(score>best){best=score;headerIndex=i;}}
  const headers=(matrix[headerIndex]||[]).map((h,i)=>String(h??'').trim()||'Column '+(i+1));
  const columns=Object.fromEntries(Object.keys(FIELD_ALIASES).map(k=>[k,findColumn(headers,k)]));
  const rows=matrix.slice(headerIndex+1);
  return rows.map((cells,i)=>{
    const get=field=>{const idx=columns[field];return idx>=0?String(cells[idx]??'').trim():'';};
    const raw=Object.fromEntries(headers.map((h,j)=>[h,cells[j]??'']));
    const company=get('company'),title=get('title'),location=get('location'),mode=get('mode'),status=get('status'),salary=get('salary'),skills=get('skills'),url=get('url'),date=get('date');
    const meaningful=[company,title,location,mode,status,salary,skills,url].some(Boolean);
    if(!meaningful)return null;
    return {id:i+1,raw,company:company||'—',title:title||'Untitled role',location:location||'—',mode:mode||'—',status:status||'Not Applied',salary:salary||'Not specified',skills,url,date};
  }).filter(Boolean);
}
const isApplied = job => {
  const status=norm(job.status);
  if(/^(not applied|not submitted|pending|to apply|yet to apply|saved|shortlisted|interested|new|open)$/.test(status))return false;
  return /(^| )(applied|submitted|interview|selected|offer|rejected|assessment|screening|hired)( |$)/.test(status);
};
function App(){
 const [jobs,setJobs]=useState([]),[fileName,setFileName]=useState(''),[sheetNames,setSheetNames]=useState([]),[sheet,setSheet]=useState(''),[workbook,setWorkbook]=useState(null),[tab,setTab]=useState('All Jobs'),[query,setQuery]=useState(''),[mode,setMode]=useState('All work modes'),[sort,setSort]=useState('company'),[notice,setNotice]=useState('Upload your Excel or CSV tracker to load your own jobs.');
 const input=useRef(null);
 const loadSheet=(wb,name)=>{const parsed=normalizeSheet(wb.Sheets[name]);setJobs(parsed);setSheet(name);setNotice(parsed.length?'Loaded '+parsed.length+' job rows from “'+name+'”.':'No recognizable job rows found. Check that the selected sheet has a header row.');};
 const onFile=async e=>{const f=e.target.files?.[0];if(!f)return;try{const wb=XLSX.read(await f.arrayBuffer(),{type:'array',cellDates:true});setWorkbook(wb);setFileName(f.name);setSheetNames(wb.SheetNames);if(wb.SheetNames.length)loadSheet(wb,wb.SheetNames[0]);setTab('All Jobs');setQuery('');}catch(err){setNotice('Could not read this file. Upload a valid .xlsx, .xls or .csv file.');}e.target.value='';};
 const applied=jobs.filter(isApplied).length;
 const filtered=useMemo(()=>jobs.filter(j=>{const matchTab=tab==='All Jobs'||(tab==='Applied'&&isApplied(j))||(tab==='Not Applied'&&!isApplied(j))||(tab==='Individual'&&/individual|freelance|contract/i.test(String(j.mode)));const hay=[j.company,j.title,j.location,j.mode,j.status,j.salary,j.skills].join(' ').toLowerCase();return matchTab&&hay.includes(query.toLowerCase())&&(mode==='All work modes'||String(j.mode).toLowerCase().includes(mode.toLowerCase()));}).sort((a,b)=>String(a[sort]||'').localeCompare(String(b[sort]||''))),[jobs,tab,query,mode,sort]);
 const modes=[...new Set(jobs.map(j=>j.mode).filter(x=>x&&x!=='—'))];
 const exportFile=()=>{const ws=XLSX.utils.json_to_sheet(jobs.map(j=>j.raw));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Jobs');XLSX.writeFile(wb,'CareerOS-export.xlsx');};
 return <div className="app-shell"><aside className="sidebar"><div className="brand"><div className="brand-mark">C</div><div><b>CareerOS</b><small>CAREER COMMAND CENTER</small></div></div><div className="side-label">WORKSPACE</div>{['All Jobs','Applied','Not Applied','Individual'].map((t,i)=><button key={t} className={'nav-item '+(tab===t?'active':'')} onClick={()=>setTab(t)}><span>{['▦','✓','◷','◇'][i]}</span>{t}</button>)}<div className="side-bottom"><div className="profile-dot">M</div><div><b>My workspace</b><small>Personal job tracker</small></div></div></aside>
 <main className="main"><header className="topbar"><div><div className="eyebrow">YOUR CAREER, IN CONTROL</div><h1>Job application <span>overview</span></h1></div><div className="top-actions"><button className="ghost-btn" onClick={exportFile} disabled={!jobs.length}>↓ Export</button><button className="primary-btn" onClick={()=>input.current?.click()}>＋ Update Data</button><input ref={input} type="file" accept=".xlsx,.xls,.csv" hidden onChange={onFile}/></div></header>
 <section className="welcome"><div><div className="eyebrow light">CAREEROS DASHBOARD</div><h2>Make your next move<br/><em>count.</em></h2><p>One clear view of every opportunity, application and next step.</p></div><div className="welcome-orb"><div className="orb-ring ring1"/><div className="orb-ring ring2"/><div className="orb-core">✦</div></div><div className="welcome-foot">{fileName?'SOURCE · '+fileName:'READY FOR YOUR TRACKER'} <span>● LIVE IN THIS SESSION</span></div></section>
 <section className="stats"><Stat label="TOTAL OPPORTUNITIES" value={jobs.length} note="Rows loaded from your file" icon="▦"/><Stat label="APPLICATIONS SENT" value={applied} note="Applied / submitted status" icon="↗"/><Stat label="YET TO APPLY" value={Math.max(0,jobs.length-applied)} note="Remaining opportunities" icon="◷"/><Stat label="APPLICATION RATE" value={jobs.length?Math.round(applied/jobs.length*100)+'%':'—'} note="Applied ÷ total jobs" icon="◉"/></section>
 <section className="content-card"><div className="section-head"><div><div className="eyebrow">OPPORTUNITY PIPELINE</div><h3>{tab}</h3><p>{jobs.length?filtered.length+' matching opportunities from your tracker':'Your uploaded spreadsheet will appear here'}</p></div><div className="head-actions"><select value={mode} onChange={e=>setMode(e.target.value)}><option>All work modes</option>{modes.map(m=><option key={m}>{m}</option>)}</select><button className="ghost-btn" onClick={()=>input.current?.click()}>↑ Upload file</button></div></div>
 <div className="tabs">{['All Jobs','Applied','Not Applied','Individual'].map(t=><button key={t} className={tab===t?'selected':''} onClick={()=>setTab(t)}>{t}<span>{t==='All Jobs'?jobs.length:t==='Applied'?applied:t==='Not Applied'?jobs.length-applied:jobs.filter(j=>/individual|freelance|contract/i.test(String(j.mode))).length}</span></button>)}</div>
 <div className="table-tools"><div className="search"><span>⌕</span><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search company, role, location, skills..."/></div><label className="sort-label">SORT BY <select value={sort} onChange={e=>setSort(e.target.value)}><option value="company">Company</option><option value="title">Job title</option><option value="location">Location</option><option value="status">Status</option></select></label></div>
 <div className="table-wrap"><table><thead><tr><th>COMPANY / ROLE</th><th>LOCATION</th><th>WORK MODE</th><th>SALARY</th><th>STATUS</th><th>LINK</th></tr></thead><tbody>{filtered.map(j=><tr key={j.id}><td><div className="company-cell"><div className="company-logo">{j.company.slice(0,1).toUpperCase()}</div><div><b>{j.company}</b><small>{j.title}</small></div></div></td><td>{j.location}</td><td><span className="mode-pill">{j.mode}</span></td><td>{j.salary}</td><td><span className={'status '+(isApplied(j)?'done':'pending')}>{j.status}</span></td><td>{j.url?<a className="apply-link" href={j.url} target="_blank" rel="noreferrer">Open ↗</a>:'—'}</td></tr>)}{!filtered.length&&<tr><td colSpan="6" className="empty"><div className="empty-icon">⇧</div><b>{jobs.length?'No matching jobs':'No spreadsheet loaded yet'}</b><p>{jobs.length?'Try changing your search or filters.':'Upload your Excel (.xlsx/.xls) or CSV file to populate this dashboard.'}</p><button className="primary-btn" onClick={()=>input.current?.click()}>Choose spreadsheet</button></td></tr>}</tbody></table></div>
 <div className="table-footer"><span>Showing {filtered.length} of {jobs.length} opportunities</span><span>Data stays in this browser session · No spreadsheet is uploaded to a server</span></div></section>
 <footer>CAREEROS <span>BUILT FOR YOUR NEXT OPPORTUNITY</span><span>{sheetNames.length>1&&workbook?<label className="sheet-picker">SHEET <select value={sheet} onChange={e=>loadSheet(workbook,e.target.value)}>{sheetNames.map(n=><option key={n}>{n}</option>)}</select></label>:notice}</span></footer></main></div>
}
function Stat({label,value,note,icon}){return <div className="stat-card"><div className="stat-top"><span>{label}</span><i>{icon}</i></div><strong>{value}</strong><small>{note}</small></div>}
export default App;
