
// Utility functions
const readFileAsArrayBuffer = (file) => new Promise((res, rej) => {
  const fr = new FileReader();
  fr.onload = e => res(e.target.result);
  fr.onerror = e => rej(e);
  fr.readAsArrayBuffer(file);
});

function normalizeString(s){
  if(s === null || s === undefined) return "";
  return String(s).trim().toUpperCase();
}

function isRegisterLike(str){
  if(!str) return false;
  return /\d/.test(str) && str.trim().length >= 3;
}

function objectToCSV(rows, columns){
  const esc = v => '"' + String(v).replace(/"/g,'""') + '"';
  const header = columns.map(esc).join(",");
  const lines = rows.map(r => columns.map(c => esc(r[c] ?? "")).join(","));
  return [header, ...lines].join("\r\n");
}

function downloadBlob(data, filename, mime){
  const blob = new Blob([data], {type: mime});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

// Template generation
document.getElementById('downloadMasterTpl').onclick = () => {
  const arr = [
    ["RegNo","MA3351","PH3251","CS3391"],
    ["2123001",18,20,19],
    ["2123002",15,18,20]
  ];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(arr);
  XLSX.utils.book_append_sheet(wb, ws, "Master");
  const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
  downloadBlob(wbout, "Master_Template.xlsx", "application/octet-stream");
};

document.getElementById('downloadRovanTpl').onclick = () => {
  const arr = [
    ["RegNo","CS3391","MA3351","PH3251"],
    ["2123001",18,18,20],
    ["2123002",20,15,18]
  ];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(arr);
  XLSX.utils.book_append_sheet(wb, ws, "Rovan");
  const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
  downloadBlob(wbout, "Rovan_Template.xlsx", "application/octet-stream");
};

// Excel parsing
async function parseToTable(file) {
  const ab = await readFileAsArrayBuffer(file);
  const wb = XLSX.read(ab, {type:'array'});
  const name = wb.SheetNames[0];
  const ws = wb.Sheets[name];
  const json = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:""});
  while(json.length && json[json.length-1].every(c => c === "")) json.pop();
  const cols = json[0] ? json[0].length : 0;
  const colNonEmpty = new Array(cols).fill(false);
  json.forEach(row => row.forEach((v,i) => { if(v !== "" && v !== null && v !== undefined) colNonEmpty[i] = true; }));
  const maxCol = colNonEmpty.lastIndexOf(true) + 1 || cols;
  const cleaned = json.map(r => r.slice(0,maxCol));
  return cleaned;
}

function arrayToDataFrame(arr) {
  if(!arr || arr.length===0) return {header:[], rows:[]};
  const header = arr[0].map(c => c===null? "": String(c));
  const rows = arr.slice(1).map(r => {
    const obj = {};
    for(let i=0;i<header.length;i++) obj[header[i]??`COL${i}`] = r[i] ?? "";
    return obj;
  });
  return {header, rows};
}

function detectOrientation(df){
  const firstColVals = df.rows.map(r => normalizeString(r[df.header[0]]));
  const countRegLike = firstColVals.filter(v => isRegisterLike(v)).length;
  if(countRegLike >= Math.max(3, Math.floor(0.3*firstColVals.length))) return "rows";
  const firstRowHeaderLike = df.header.filter(h => isRegisterLike(h)).length;
  if(firstRowHeaderLike >= Math.max(3, Math.floor(0.3*df.header.length))) return "columns";
  return "rows";
}

function dfToMatrix(df, orientation){
  const regs = new Set();
  const courses = new Set();
  const matrix = {};
  if(orientation === "rows"){
    const regCol = df.header[0];
    df.rows.forEach(r => {
      const reg = normalizeString(r[regCol]);
      if(!reg) return;
      regs.add(reg);
      matrix[reg] = matrix[reg] || {};
      for(let i=1;i<df.header.length;i++){
        const course = normalizeString(df.header[i]);
        if(course === "") continue;
        courses.add(course);
        matrix[reg][course] = r[df.header[i]] === null? "" : String(r[df.header[i]]).trim();
      }
    });
  } else {
    const regHeaders = df.header.slice(1).map(h => normalizeString(h));
    df.rows.forEach(r => {
      const course = normalizeString(r[df.header[0]]);
      if(!course) return;
      courses.add(course);
      for(let i=1;i<df.header.length;i++){
        const reg = regHeaders[i-1];
        if(!reg) continue;
        regs.add(reg);
        matrix[reg] = matrix[reg] || {};
        matrix[reg][course] = r[df.header[i]] === null? "" : String(r[df.header[i]]).trim();
      }
    });
  }
  return { regs: Array.from(regs).sort(), courses: Array.from(courses).sort(), matrix };
}

function unifyAndCompare(masterData, rovanData){
  const regs = Array.from(new Set([...masterData.regs, ...rovanData.regs])).sort();
  const courses = Array.from(new Set([...masterData.courses, ...rovanData.courses])).sort();
  const rows = [];
  const summary = {total:0, matches:0, mismatches:0, missing_master:0, missing_rovan:0};
  const perCourse = {};
  courses.forEach(c => perCourse[c] = {total:0, matches:0, mismatches:0, missing_master:0, missing_rovan:0});
  regs.forEach(reg => {
    courses.forEach(course => {
      const m = (masterData.matrix[reg] && masterData.matrix[reg][course]) ? normalizeString(masterData.matrix[reg][course]) : "";
      const r = (rovanData.matrix[reg] && rovanData.matrix[reg][course]) ? normalizeString(rovanData.matrix[reg][course]) : "";
      if(m === "" && r === "") return;
      summary.total++;
      perCourse[course].total++;
      let status = "";
      if(m === "" && r !== ""){
        status = "Missing in Master";
        summary.missing_master++;
        perCourse[course].missing_master++;
      } else if(r === "" && m !== ""){
        status = "Missing in Rovan";
        summary.missing_rovan++;
        perCourse[course].missing_rovan++;
      } else if(m === r){
        status = "Match";
        summary.matches++;
        perCourse[course].matches++;
      } else {
        status = "Mismatch";
        summary.mismatches++;
        perCourse[course].mismatches++;
      }
      rows.push({Register:reg, Course:course, MasterMark: m, RovanMark: r, Status: status});
    });
  });
  summary.matchPercent = summary.total ? Math.round((summary.matches/summary.total)*10000)/100 : 0;
  summary.mismatchPercent = summary.total ? Math.round((summary.mismatches/summary.total)*10000)/100 : 0;
  return {rows, summary, perCourse};
}

// UI + Events
let table = null;
let lastResult = null;

function renderSummary(summary){
  document.getElementById('totalCells').textContent = summary.total || 0;
  document.getElementById('matches').textContent = summary.matches || 0;
  document.getElementById('mismatches').textContent = summary.mismatches || 0;
  document.getElementById('matchPercent').textContent = `${summary.matchPercent || 0}%`;
  document.getElementById('mismatchPercent').textContent = `${summary.mismatchPercent || 0}%`;
  document.getElementById('missingCounts').textContent = `${summary.missing_master || 0} / ${summary.missing_rovan || 0}`;
}

function renderCourseSummary(perCourse){
  const el = document.getElementById('courseSummary');
  el.innerHTML = "";
  const arr = Object.keys(perCourse).map(k => ({course:k, ...perCourse[k]})).sort((a,b)=> b.mismatches - a.mismatches);
  arr.forEach(c => {
    if(c.total === 0) return;
    const pct = Math.round((c.mismatches / c.total)*10000)/100 || 0;
    const div = document.createElement('div');
    div.className = "py-1 mb-1 border-bottom";
    div.innerHTML = `<strong>${c.course}</strong> &nbsp; <span class="small-muted">(${c.total})</span>
      <div class="small-muted">Mismatch: ${c.mismatches} • MissingM:${c.missing_master} • MissingR:${c.missing_rovan} • ${pct}%</div>`;
    el.appendChild(div);
  });
  return arr;
}

let courseChart = null;
function renderChart(perCourse){
  const data = Object.keys(perCourse).map(k => {
    const c = perCourse[k];
    return {course:k, mismatchPercent: c.total? Math.round((c.mismatches/c.total)*10000)/100 : 0, total:c.total};
  }).filter(x=>x.total>0).sort((a,b)=> b.mismatchPercent - a.mismatchPercent).slice(0,12);
  const labels = data.map(d=>d.course);
  const values = data.map(d=>d.mismatchPercent);
  const ctx = document.getElementById('courseChart').getContext('2d');
  if(courseChart) courseChart.destroy();
  courseChart = new Chart(ctx, {
    type:'bar',
    data:{labels, datasets:[{label:'Mismatch %', data:values}]},
    options:{responsive:true, scales:{y:{beginAtZero:true, max:100}}}
  });
}

function populateTable(rows){
  if($.fn.dataTable.isDataTable('#reportTable')) {
    $('#reportTable').DataTable().clear().destroy();
    $('#reportTable tbody').empty();
  }
  const tbody = document.querySelector('#reportTable tbody');
  rows.forEach(r => {
    const tr = document.createElement('tr');
    let statusClass = '';
    if(r.Status === 'Match') statusClass = 'status-Match';
    else if(r.Status === 'Mismatch') statusClass = 'status-Mismatch';
    else if(r.Status === 'Missing in Master') statusClass = 'status-MissingMaster';
    else if(r.Status === 'Missing in Rovan') statusClass = 'status-MissingRovan';
    tr.className = statusClass;
    tr.innerHTML = `<td>${r.Register}</td>
                    <td>${r.Course}</td>
                    <td>${r.MasterMark}</td>
                    <td>${r.RovanMark}</td>
                    <td>${r.Status}</td>`;
    tbody.appendChild(tr);
  });
  table = $('#reportTable').DataTable({
    pageLength: 25,
    lengthMenu: [10,25,50,100],
    columns: [{},{},{},{},{searchable:true}],
    order: [[1,"asc"]]
  });
}

// Events
document.getElementById('compareBtn').onclick = async () => {
  const masterFile = document.getElementById('masterFile').files[0];
  const rovanFile = document.getElementById('rovanFile').files[0];
  if(!masterFile || !rovanFile){
    alert('Please choose both Master and Rovan files.');
    return;
  }
  try {
    const [marr, rarr] = await Promise.all([parseToTable(masterFile), parseToTable(rovanFile)]);
    const mdf = arrayToDataFrame(marr);
    const rdf = arrayToDataFrame(rarr);
    const morient = detectOrientation(mdf);
    const rorient = detectOrientation(rdf);
    const mdata = dfToMatrix(mdf, morient);
    const rdata = dfToMatrix(rdf, rorient);
    const result = unifyAndCompare(mdata, rdata);
    lastResult = {result, mdata, rdata, morient, rorient};
    renderSummary(result.summary);
    renderCourseSummary(result.perCourse);
    renderChart(result.perCourse);
    populateTable(result.rows);
    document.querySelector('#reportTable').scrollIntoView({behavior:'smooth'});
  } catch(err){
    console.error(err);
    alert('Error parsing files. Please ensure they are valid Excel files.');
  }
};

document.getElementById('resetBtn').onclick = () => {
  document.getElementById('masterFile').value = "";
  document.getElementById('rovanFile').value = "";
  if($.fn.dataTable.isDataTable('#reportTable')) $('#reportTable').DataTable().clear().destroy();
  document.querySelector('#reportTable tbody').innerHTML = "";
  document.getElementById('totalCells').textContent = "0";
  document.getElementById('matches').textContent = "0";
  document.getElementById('mismatches').textContent = "0";
  document.getElementById('matchPercent').textContent = "0%";
  document.getElementById('mismatchPercent').textContent = "0%";
  document.getElementById('missingCounts').textContent = "0 / 0";
  if(courseChart) courseChart.destroy();
  lastResult = null;
};

document.getElementById('downloadCSV').onclick = () => {
  if(!lastResult){ alert('No report to download.'); return; }
  const rows = lastResult.result.rows;
  if(rows.length === 0){ alert('No differences found (report empty).'); return; }
  const csv = objectToCSV(rows, ["Register","Course","MasterMark","RovanMark","Status"]);
  downloadBlob(csv, 'compare_report.csv', 'text/csv;charset=utf-8;');
};

document.getElementById('downloadXlsx').onclick = () => {
  if(!lastResult){ alert('No report to download.'); return; }
  const rows = lastResult.result.rows;
  if(rows.length === 0){ alert('No differences found (report empty).'); return; }
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Report');
  const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
  downloadBlob(wbout, 'compare_report.xlsx', 'application/octet-stream');
};

document.getElementById('helpBtn').onclick = (e) => {
  e.preventDefault();
  alert(`How to use:
1. (Optional) Download templates to see expected formats.
2. Upload Master (Dean) and Rovan Excel files. The first column should be RegNo (if not, the app attempts to auto-detect).
3. Click "Compare Now".
4. View summary, chart, and detailed table. Use search to filter.
5. Download CSV/XLSX report.
All processing is local in your browser.`);
};
