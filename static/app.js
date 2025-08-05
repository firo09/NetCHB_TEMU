// app.js

// —— 硬编码 Excel 密码 —— 
const EXCEL_PASSWORD = 'xU$&#3_*VB';

// 全局保存当前文件和默认 MAWB
let currentFile = null;
let currentDefaultMawb = '';

// 防缓存
const ts = Date.now();
const CONFIG_PATH = 'config';

const uploadBtn   = document.getElementById('upload-btn');
const fileInput   = document.getElementById('file-input');
const loadingMsg  = document.getElementById('loading-msg');
const generateBtn = document.getElementById('generate-btn');

let ruleConfig = [], htsData = [], midData = [];

// 日期格式化工具，支持 m/d/yyyy, yyyy-mm-dd, mm/dd/yyyy 等
function formatDateByPattern(date, pattern) {
  const yyyy = date.getFullYear();
  const m = date.getMonth() + 1;
  const d = date.getDate();
  const pad = n => (n < 10 ? '0' + n : n);

  return pattern
    .replace(/yyyy/gi, yyyy)
    .replace(/mm/g, pad(m))
    .replace(/m/g, m)
    .replace(/dd/g, pad(d))
    .replace(/d/g, d);
}

// 只用于mawb sheet提取
function getValueFromMawbSheet(mawbSheetArr, colName) {
  if (!Array.isArray(mawbSheetArr) || mawbSheetArr.length < 2) return '';
  const header = mawbSheetArr[0] || [];
  const row    = mawbSheetArr[1] || [];
  const idx = header.findIndex(h => (h || '').toString().trim().toLowerCase() === colName.trim().toLowerCase());
  if (idx === -1) {
    console.warn('找不到表头对应列', colName, header);
    return '';
  }
  return row[idx] || '';
}

// 工具：字符串转合法 DOM id
function sanitize(label) {
  return label.replace(/[^\w]/g, '_');
}

// 工具：根据 Format 字符串动态生成正则表达式
function buildRegex(fmt) {
  let regexStr = fmt.replace(/([.+?^=!:${}()|\[\]\/\\])/g, '\\$1');
  regexStr = regexStr.replace(/y{4}/g, '\\d{4}');
  regexStr = regexStr.replace(/m{1,2}/gi, '\\d{1,2}');
  regexStr = regexStr.replace(/d{1,2}/gi, '\\d{1,2}');
  return new RegExp('^' + regexStr + '$');
}

// 1. 并行加载三份 JSON 配置
Promise.all([
  fetch(`${CONFIG_PATH}/rule.json?ts=${ts}`).then(r => r.json()),
  fetch(`${CONFIG_PATH}/hts.json?ts=${ts}`).then(r => r.json()),
  fetch(`${CONFIG_PATH}/mid.json?ts=${ts}`).then(r => r.json())
])
.then(([rule, hts, mid]) => {
  ruleConfig = rule;
  htsData    = hts;
  midData    = mid;
  uploadBtn.disabled = false;
  uploadBtn.classList.remove('opacity-50');
  loadingMsg.innerText = '';
})
.catch(e => {
  console.error('Failed to load configs', e);
  loadingMsg.innerText = 'Failed to load configuration';
});

// 2. 绑定“Select File”按钮
uploadBtn.addEventListener('click', () => fileInput.click());

// 3. 处理文件选中（点击或拖拽）
fileInput.addEventListener('change', () => {
  if (!fileInput.files.length) {
    alert('Please select a file');
    return;
  }
  currentFile = fileInput.files[0];
  const base = currentFile.name.replace(/\.(xlsx|xls|csv)$/i, '');
  const m    = base.match(/(\d{11})$/);
  currentDefaultMawb = m ? m[1] : '';
  document.getElementById('upload-section').classList.add('hidden');
  document.getElementById('form-section').classList.remove('hidden');
  renderForm(currentDefaultMawb);
});

// 4. 只绑定一次“Generate & Download”按钮
generateBtn.addEventListener('click', () => {
  if (!currentFile) {
    alert('No file selected');
    return;
  }
  generateAndDownload();
});

// 5. 渲染动态表单：按 Label 去重、保留 default_value、placeholder=Format
function renderForm(defaultMawb) {
  const formEl = document.getElementById('dynamic-form');
  formEl.innerHTML = '';

  const labels = [];
  const primaryRuleFor = {};
  for (const r of ruleConfig) {
    if (r.Source.trim().toLowerCase() === 'user_input') {
      const lab = r.Label.trim();
      if (!labels.includes(lab)) {
        labels.push(lab);
        primaryRuleFor[lab] = r;
      }
    }
  }

  for (const label of labels) {
    const rule = primaryRuleFor[label];
    const id = sanitize(label);

    let defaultVal = '';
    if (rule.default_value?.startsWith('<from_filename:')) {
      defaultVal = defaultMawb;
    } else if (label.toUpperCase() === 'MAWB') {
      defaultVal = defaultMawb;
    } else {
      defaultVal = rule.default_value || '';
    }

    const fmt = (rule.Format || '').trim();
    const placeholder = fmt || '';

    const wrapper = document.createElement('div');

    if (rule.has_dropdown?.trim().toUpperCase() === 'Y') {
      const opts = (rule.dropdown_options || '').split(',').map(o => o.trim()).filter(Boolean);
      wrapper.innerHTML = `
        <label for="${id}" class="font-semibold block mb-1">${label}</label>
        <select id="${id}" class="border rounded px-2 py-1 w-full">
          <option value="">--Select--</option>
          ${opts.map(o=>`<option value="${o}"${o===defaultVal?' selected':''}>${o}</option>`).join('')}
        </select>
      `;
    }
    else {
      wrapper.innerHTML = `
        <label for="${id}" class="font-semibold block mb-1">${label}</label>
        <input type="text"
               id="${id}"
               value="${defaultVal}"
               ${placeholder?`placeholder="${placeholder}"`:''}
               data-format="${fmt}"
               class="border rounded px-2 py-1 w-full placeholder-gray-400" />
      `;
      if (fmt.toLowerCase() === 'yyyy/m/d' && typeof flatpickr === 'function') {
        flatpickr(`#${id}`, { dateFormat: 'Y/m/d', allowInput: true });
      }
    }

    formEl.appendChild(wrapper);
  }
}

// 7. 生成并下载
async function generateAndDownload() {
  const formValues = {};
  document.querySelectorAll('#dynamic-form input, #dynamic-form select')
    .forEach(el => formValues[el.id] = el.value.trim());

  const buf = await currentFile.arrayBuffer();
  let wb;
  try {
    wb = XLSX.read(buf, { type: 'array', password: EXCEL_PASSWORD });
  } catch (err) {
    return alert('Failed to open encrypted file: ' + err.message);
  }

  const sheetData = {};
  let mawbSheetArr = [];
  for (const name of wb.SheetNames) {
    const key = name.trim().toLowerCase();
    const ws  = wb.Sheets[name];
    if (key === 'hawb') {
      const raw    = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
      const header = raw[1] || [];
      const rows   = raw.slice(2).filter(r => r.some(c => c !== ''));
      sheetData['hawb'] = rows.map(rw => {
        const o = {};
        header.forEach((h,i) => o[h] = rw[i] || '');
        return o;
      });
    } else if (key === 'mawb') {
      mawbSheetArr = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
      sheetData['mawb'] = XLSX.utils.sheet_to_json(ws, { defval:'' });
    } else {
      sheetData[key] = XLSX.utils.sheet_to_json(ws, { defval:'' });
    }
  }

  const main = sheetData['hawb'] || [];
  const output = [];
  const prog = document.getElementById('progress');
  const pt   = document.getElementById('progress-text');
  document.getElementById('progress-container').classList.remove('hidden');

  for (let i = 0; i < main.length; i++) {
    const out = {};

    for (const cfg of ruleConfig) {
      const col = cfg.Column;
      const src = cfg.Source.trim().toLowerCase();

      if (src === 'fixed') {
        out[col] = cfg.Value || '';
      }
      else if (src === 'user_upload') {
        const sk = (cfg.Sheet||'').trim().toLowerCase();
        if (sk === 'mawb' && cfg.Reference) {
          out[col] = getValueFromMawbSheet(mawbSheetArr, cfg.Reference);
        } else {
          const arr= sheetData[sk] || [];
          const row= sk === 'mawb' ? (arr[0]||{}) : (arr[i]||{});
          out[col] = row[cfg.Reference] || '';
        }
      }
      else if (src === 'user_input') {
        const label = cfg.Label.trim();
        let v = formValues[sanitize(label)] || '';
        const m = (cfg.Parsing||'').match(/(left|right)\((\d+)\)/i);
        if (m) {
          const n = parseInt(m[2],10);
          v = m[1].toLowerCase() === 'left' ? v.slice(0,n) : v.slice(-n);
        }
        out[col] = v;
      }
      else if (src === 'system') {
        const d = new Date();
        const fmt = (cfg.Format || '').trim();
        if (fmt) {
          out[col] = formatDateByPattern(d, fmt);
        } else {
          out[col] = `${d.getFullYear()}/${d.getMonth()+1}/${d.getDate()}`;
        }
      }
    }

    // HTS 映射
    (() => {
      const raw  = (out.HTS||'').toString();
      const digs = (raw.match(/\d+/g)||[]).join('').slice(0,8);
      const hit  = htsData.find(r=>r.HTS===digs);
      if (hit) ['HTS-1','HTS-2','HTS-3','HTS-4','HTS-5']
        .forEach(c=> out[c] = hit[c] || '');
    })();

    // MID 映射并清空
    (() => {
      const nm  = (out.ManufacturerName||'').trim();
      const hit = midData.find(r => nm.includes(r.ManufacturerName));
      if (hit) {
        out.ManufacturerCode = hit.ManufacturerCode || '';
        ['ManufacturerName','ManufacturerStreetAddress','ManufacturerCity','ManufacturerPostalCode','ManufacturerCountry']
          .forEach(f=> out[f] = '');
      }
    })();

    output.push(out);

    if ((i+1)%20===0 || i===main.length-1) {
      const pct = Math.round(((i+1)/main.length)*100);
      pt.innerText     = `${pct}%`;
      prog.style.width = `${pct}%`;
      await new Promise(r => setTimeout(r,0));
    }
  }

  // ---------- 日期格式自动识别/设置部分 begin ----------
  const header = ruleConfig.map(r=>r.Column);
  const aoa    = [header].concat(output.map(o=> header.map(c=>o[c]||'')));

  // 找所有应设为日期格式的列
  const dateCols = ruleConfig
    .map((r, idx) => ({idx, fmt: (r.Format||'').toLowerCase()}))
    .filter(r => r.fmt.includes('yyyy') && r.fmt.includes('m') && r.fmt.includes('d'))
    .map(r => r.idx);

  const ws2 = XLSX.utils.aoa_to_sheet(aoa);

  // 设置日期列单元格的Excel类型和格式
  for (let r = 1; r < aoa.length; r++) { // r=1是跳过表头
    for (const c of dateCols) {
      const colLetter = XLSX.utils.encode_col(c);
      const cellRef = colLetter + (r+1);
      const val = aoa[r][c];
      if (val) {
        let d = new Date(val);
        // 兼容 m/d/yyyy、yyyy-mm-dd 等格式
        if (isNaN(d.getTime())) {
          const m = String(val).match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
          if (m) d = new Date(m[3], m[1]-1, m[2]);
        }
        if (!isNaN(d.getTime())) {
          ws2[cellRef].t = 'd';
          ws2[cellRef].z = 'm/d/yyyy'; // 如有需要可进一步动态匹配 Format 字段
          ws2[cellRef].v = d;
        }
      }
    }
  }
  // ---------- 日期格式自动识别/设置部分 end ----------

  const wb2    = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws2, 'Sheet1');

  const mawbOrig = formValues[sanitize('MAWB')] || currentDefaultMawb;
  XLSX.writeFile(wb2, `${mawbOrig}_NETChb_TEMU.xlsx`);
}
