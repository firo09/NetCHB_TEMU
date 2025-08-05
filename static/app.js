// app.js

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

// 工具：字符串转合法 DOM id
function sanitize(label) {
  return label.replace(/[^\w]/g, '_');
}

// 1. 并行加载配置
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
  console.log('Configs loaded');
})
.catch(e => {
  console.error('Failed to load configs', e);
  loadingMsg.innerText = 'Failed to load configuration';
});

// 2. 绑定 “Select File” 按钮
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
  console.log('Parsed MAWB=', currentDefaultMawb);

  document.getElementById('upload-section').classList.add('hidden');
  document.getElementById('form-section').classList.remove('hidden');
  renderForm(currentDefaultMawb);
});

// 4. 只绑定一次 “Generate & Download” 按钮
generateBtn.addEventListener('click', () => {
  console.log('Generate clicked, file=', currentFile);
  if (!currentFile) {
    alert('No file selected');
    return;
  }
  generateAndDownload();
});

// 5. 渲染动态表单
function renderForm(defaultMawb) {
  console.log('Rendering form with defaultMawb=', defaultMawb);
  const formEl = document.getElementById('dynamic-form');
  formEl.innerHTML = '';

  // 按 Label 去重、保证顺序
  const seen = new Set();
  const inputs = ruleConfig.filter(r => {
    if (r.Source.trim().toLowerCase() !== 'user_input') return false;
    const lab = r.Label.trim();
    if (seen.has(lab)) return false;
    seen.add(lab);
    return true;
  });

  inputs.forEach(rule => {
    const label = rule.Label.trim() || rule.Column;
    const id    = sanitize(rule.Column);
    let defaultVal = '';

    if (rule.default_value?.startsWith('<from_filename:')) {
      defaultVal = defaultMawb;
    } else if (label.toUpperCase() === 'MAWB') {
      defaultVal = defaultMawb;
    } else {
      defaultVal = rule.default_value || '';
    }

    const wrapper = document.createElement('div');
    if ((rule.Format || '').trim().toLowerCase() === 'yyyy/m/d') {
      wrapper.innerHTML = `
        <label for="${id}" class="font-semibold block mb-1">${label}</label>
        <input type="text"
               id="${id}"
               data-format="yyyy/m/d"
               placeholder="yyyy/m/d"
               value="${defaultVal}"
               class="border rounded px-2 py-1 w-full" />
      `;
      if (typeof flatpickr === 'function') {
        flatpickr(`#${id}`, { dateFormat: 'Y/M/d', allowInput: true });
      }
    }
    else if ((rule.has_dropdown || '').trim().toUpperCase() === 'Y') {
      const opts = (rule.dropdown_options || '')
        .split(',').map(o => o.trim()).filter(Boolean);
      wrapper.innerHTML = `
        <label for="${id}" class="font-semibold block mb-1">${label}</label>
        <select id="${id}" class="border rounded px-2 py-1 w-full">
          <option value="">--Select--</option>
          ${opts.map(o => `<option${o===defaultVal?' selected':''}>${o}</option>`).join('')}
        </select>
      `;
    }
    else {
      wrapper.innerHTML = `
        <label for="${id}" class="font-semibold block mb-1">${label}</label>
        <input type="text"
               id="${id}"
               value="${defaultVal}"
               class="border rounded px-2 py-1 w-full" />
      `;
    }

    formEl.appendChild(wrapper);
  });

  setupValidation();
  generateBtn.disabled = true;
}

// 6. 校验 & 按钮状态控制
function setupValidation() {
  const elems = Array.from(document.querySelectorAll('#dynamic-form input, #dynamic-form select'));
  elems.forEach(el => {
    const fmt = el.dataset.format;
    el.addEventListener('input', () => validateField(el));
    if (fmt) el.addEventListener('blur', () => validateField(el));
    validateField(el);
  });
}

function validateField(el) {
  const fmt = el.dataset.format;
  const v   = el.value.trim();
  let valid = false;

  if (el.tagName === 'SELECT') {
    valid = v !== '';
  } else if (fmt === 'yyyy/m/d') {
    valid = /^[0-9]{4}\/(?:[1-9]|1[0-2])\/(?:[1-9]|[12]\d|3[01])$/.test(v);
  } else {
    valid = v !== '';
  }

  el.classList.toggle('border-red-500', !valid);

  const allValid = Array.from(document.querySelectorAll('#dynamic-form input, #dynamic-form select'))
    .every(e => {
      const f = e.dataset.format;
      const val = e.value.trim();
      if (e.tagName === 'SELECT') return val !== '';
      if (f === 'yyyy/m/d') {
        return /^[0-9]{4}\/(?:[1-9]|1[0-2])\/(?:[1-9]|[12]\d|3[01])$/.test(val);
      }
      return val !== '';
    });

  generateBtn.disabled = !allValid;
}

// 7. 生成并下载
async function generateAndDownload() {
  console.log('generateAndDownload start');
  const defaultMawb = currentDefaultMawb;
  const file        = currentFile;

  // 正确做法：遍历真实 DOM 上的 input/select
  const formValues = {};
  document
    .querySelectorAll('#dynamic-form input, #dynamic-form select')
    .forEach(el => {
      formValues[el.id] = el.value.trim();
    });

  // 读取文件
  const buf = await file.arrayBuffer();
  const wb  = XLSX.read(buf, { type: 'array' });

  // 解析 sheets
  const sheetData = {};
  wb.SheetNames.forEach(name => {
    const key = name.trim().toLowerCase();
    const ws  = wb.Sheets[name];
    if (key === 'hawb') {
      const raw    = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
      const header = raw[1] || [];
      const rows   = raw.slice(2).filter(r => r.some(c => c !== ''));
      sheetData['hawb'] = rows.map(rw => {
        const o = {};
        header.forEach((h,i)=>o[h]=rw[i]||'');
        return o;
      });
    } else if (key === 'mawb') {
      sheetData['mawb'] = XLSX.utils.sheet_to_json(ws, { defval:'' });
    } else {
      sheetData[key] = XLSX.utils.sheet_to_json(ws, { defval:'' });
    }
  });

  // 构建输出
  const main   = sheetData['hawb'] || [];
  const output = [];
  const prog   = document.getElementById('progress');
  const pt     = document.getElementById('progress-text');
  document.getElementById('progress-container').classList.remove('hidden');

  for (let i = 0; i < main.length; i++) {
    const out = {};

    ruleConfig.forEach(cfg => {
      const col = cfg.Column;
      const src = cfg.Source.trim().toLowerCase();

      if (src === 'fixed') {
        out[col] = cfg.Value || '';
      }
      else if (src === 'user_upload') {
        const sk = (cfg.Sheet||'').trim().toLowerCase();
        const arr= sheetData[sk]||[];
        const row= sk==='mawb'? (arr[0]||{}):(arr[i]||{});
        out[col] = row[cfg.Reference]||'';
      }
      else if (src === 'user_input') {
        // 使用 sanitize(Column) 从 formValues 中取值
        out[col] = formValues[sanitize(col)] || '';
      }
      else if (src === 'system') {
        const d = new Date();
        out[col] = `${d.getFullYear()}/${d.getMonth()+1}/${d.getDate()}`;
      }
    });

    // HTS 映射
    (() => {
      const raw  = (out.HTS||'').toString();
      const digs = (raw.match(/\d+/g)||[]).join('').slice(0,8);
      const hit  = htsData.find(r=>r.HTS===digs);
      if (hit) ['HTS-1','HTS-2','HTS-3','HTS-4','HTS-5']
        .forEach(c=> out[c] = hit[c] || '');
    })();

    // MID 映射并清空其它字段
    (() => {
      const nm = out.ManufacturerName || '';
      const hit= midData.find(r=> nm.includes(r.ManufacturerName));
      if (hit) {
        out.ManufacturerCode = hit.ManufacturerCode || '';
        ['ManufacturerName','ManufacturerStreetAddress','ManufacturerCity','ManufacturerPostalCode','ManufacturerCountry']
          .forEach(f=> out[f] = '');
      }
    })();

    output.push(out);

    if ((i+1)%20===0 || i===main.length-1) {
      const pct = Math.round(((i+1)/main.length)*100);
      // 1. 更新进度数字
      pt.innerText = `${pct}%`;
      // 2. 更新进度条宽度
      prog.style.width = `${pct}%`;
      // 3. 让浏览器渲染
      await new Promise(r=>setTimeout(r,0));
    }
  }

  // 导出
  const header = ruleConfig.map(r=>r.Column);
  const aoa    = [header].concat(output.map(o=>header.map(c=>o[c]||'')));
  const ws2    = XLSX.utils.aoa_to_sheet(aoa);
  const wb2    = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws2, 'Sheet1');

  const filename = `${formValues.MAWB || defaultMawb}_NETChb_TEMU.xlsx`;
  XLSX.writeFile(wb2, filename);
}
