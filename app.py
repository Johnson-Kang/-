import os
import re
import tempfile
from flask import Flask, request, send_file, render_template_string

import processor

app = Flask(__name__)

HTML = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>通途订单整理工具</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
  background: #f0f2f5;
  min-height: 100vh;
  display: flex;
  justify-content: center;
  align-items: center;
}
.container {
  background: #fff;
  border-radius: 16px;
  box-shadow: 0 4px 24px rgba(0,0,0,.08);
  padding: 48px;
  width: 520px;
  text-align: center;
}
h1 { font-size: 22px; color: #1a1a1a; margin-bottom: 8px; }
.subtitle { font-size: 14px; color: #888; margin-bottom: 36px; }
.upload-area {
  border: 2px dashed #d0d5dd;
  border-radius: 12px;
  padding: 48px 24px;
  cursor: pointer;
  transition: border-color .2s, background .2s;
  margin-bottom: 24px;
}
.upload-area:hover, .upload-area.dragover {
  border-color: #4f46e5;
  background: #f5f3ff;
}
.upload-icon { font-size: 40px; margin-bottom: 12px; }
.upload-text { font-size: 15px; color: #333; }
.upload-hint { font-size: 13px; color: #999; margin-top: 8px; }
#file-input { display: none; }
#file-name {
  display: inline-block;
  margin-top: 12px;
  padding: 6px 14px;
  background: #eef2ff;
  color: #4f46e5;
  border-radius: 6px;
  font-size: 13px;
}
.btn {
  display: inline-block;
  padding: 12px 36px;
  border: none;
  border-radius: 8px;
  font-size: 15px;
  cursor: pointer;
  transition: opacity .2s;
}
.btn-primary { background: #4f46e5; color: #fff; }
.btn-primary:hover { opacity: .85; }
.btn-primary:disabled { opacity: .45; cursor: not-allowed; }
.btn-download { background: #16a34a; color: #fff; margin-top: 16px; }
.status { margin-top: 20px; font-size: 14px; }
.status.success { color: #16a34a; }
.status.error { color: #dc2626; }
</style>
</head>
<body>
<div class="container">
  <h1>通途订单整理工具</h1>
  <p class="subtitle">导入未整理表格，自动识别工厂并生成整理后的表格</p>

  <div class="upload-area" id="upload-area">
    <div class="upload-icon">📂</div>
    <div class="upload-text">点击上传 或 拖拽 .xlsx 文件到此处</div>
    <div class="upload-hint">仅支持通途 ERP 导出的订单表格</div>
  </div>
  <input type="file" id="file-input" accept=".xlsx">

  <div id="file-name" style="display:none;"></div>

  <button class="btn btn-primary" id="process-btn" disabled>开始整理</button>
  <div class="status" id="status"></div>
  <a class="btn btn-download" id="download-btn" style="display:none;" download>下载已整理表格</a>
</div>

<script>
const uploadArea = document.getElementById('upload-area');
const fileInput = document.getElementById('file-input');
const fileName = document.getElementById('file-name');
const processBtn = document.getElementById('process-btn');
const statusEl = document.getElementById('status');
const downloadBtn = document.getElementById('download-btn');

let selectedFile = null;

uploadArea.addEventListener('click', () => fileInput.click());

uploadArea.addEventListener('dragover', e => { e.preventDefault(); uploadArea.classList.add('dragover'); });
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
uploadArea.addEventListener('drop', e => {
  e.preventDefault();
  uploadArea.classList.remove('dragover');
  const files = e.dataTransfer.files;
  if (files.length) setFile(files[0]);
});

fileInput.addEventListener('change', () => {
  if (fileInput.files.length) setFile(fileInput.files[0]);
});

function setFile(file) {
  if (!file.name.endsWith('.xlsx')) {
    showStatus('请选择 .xlsx 文件', 'error');
    return;
  }
  selectedFile = file;
  fileName.textContent = file.name;
  fileName.style.display = 'inline-block';
  processBtn.disabled = false;
  statusEl.textContent = '';
  downloadBtn.style.display = 'none';
}

processBtn.addEventListener('click', async () => {
  if (!selectedFile) return;
  processBtn.disabled = true;
  showStatus('正在处理中…', '');

  const formData = new FormData();
  formData.append('file', selectedFile);

  try {
    const resp = await fetch('/upload', { method: 'POST', body: formData });
    if (!resp.ok) {
      const err = await resp.json();
      showStatus('处理失败：' + (err.error || '未知错误'), 'error');
      processBtn.disabled = false;
      return;
    }
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    downloadBtn.href = url;
    let name = selectedFile.name.replace(/[(（]?未整理[)）]?/g, '').replace('.xlsx', '（已整理）.xlsx');
    downloadBtn.download = name;
    downloadBtn.style.display = 'inline-block';
    showStatus('整理完成！共处理 ' + resp.headers.get('X-Row-Count') + ' 行数据', 'success');
    processBtn.disabled = false;
  } catch (e) {
    showStatus('网络错误：' + e.message, 'error');
    processBtn.disabled = false;
  }
});

function showStatus(msg, cls) {
  statusEl.textContent = msg;
  statusEl.className = 'status ' + cls;
}
</script>
</body>
</html>'''


@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not f:
        return {'error': '未收到文件'}, 400
    if not f.filename.endswith('.xlsx'):
        return {'error': '仅支持 .xlsx 文件'}, 400

    # 保存上传文件到临时目录
    tmp_in = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    f.save(tmp_in.name)
    tmp_in.close()

    try:
        rows, headers, errors, merge_ranges = processor.process(tmp_in.name)
    except Exception as e:
        os.unlink(tmp_in.name)
        return {'error': f'处理异常：{str(e)}'}, 500

    os.unlink(tmp_in.name)

    # 写入输出到临时文件
    tmp_out = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tmp_out.close()
    processor.write_output(rows, headers, tmp_out.name, merge_ranges)

    out_name = re.sub(r'[（(]?未整理[）)]?', '', f.filename).replace('.xlsx', '（已整理）.xlsx')

    resp = send_file(
        tmp_out.name,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=out_name,
    )
    resp.headers['X-Row-Count'] = str(len(rows))

    # 清理（延迟删除，等响应发送后）
    @resp.call_on_close
    def cleanup():
        try:
            os.unlink(tmp_out.name)
        except OSError:
            pass

    return resp


if __name__ == '__main__':
    import socket
    port = 5099
    for _ in range(10):
        try:
            app.run(host='127.0.0.1', port=port, debug=False)
            break
        except OSError:
            port += 1
    else:
        input(f'无法启动：所有端口均被占用。按回车退出...')
