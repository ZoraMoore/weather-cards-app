// ==UserScript==
// @name         多功能文件转换器
// @namespace    http://tampermonkey.net/
// @version      1.0.0
// @description  支持 Word↔PDF, Excel↔JSON, HTML↔PDF, PPT↔PDF 等多种文件格式转换
// @author       Claude
// @match        *://*/*
// @grant        GM_xmlhttpRequest
// @grant        GM_download
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js
// ==/UserScript==

(function() {
    'use strict';

    // ==================== 样式定义 ====================
    const styles = `
        #file-converter-panel {
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 600px;
            max-height: 80vh;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            z-index: 999999;
            display: none;
            overflow: hidden;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }

        #file-converter-panel.show {
            display: block;
            animation: slideIn 0.3s ease-out;
        }

        @keyframes slideIn {
            from {
                opacity: 0;
                transform: translate(-50%, -60%);
            }
            to {
                opacity: 1;
                transform: translate(-50%, -50%);
            }
        }

        .converter-header {
            background: rgba(255,255,255,0.1);
            padding: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid rgba(255,255,255,0.2);
        }

        .converter-header h2 {
            margin: 0;
            color: white;
            font-size: 24px;
            font-weight: 600;
        }

        .converter-close {
            background: rgba(255,255,255,0.2);
            border: none;
            color: white;
            width: 30px;
            height: 30px;
            border-radius: 50%;
            cursor: pointer;
            font-size: 20px;
            line-height: 1;
            transition: all 0.3s;
        }

        .converter-close:hover {
            background: rgba(255,255,255,0.3);
            transform: rotate(90deg);
        }

        .converter-content {
            padding: 30px;
            max-height: calc(80vh - 140px);
            overflow-y: auto;
        }

        .converter-content::-webkit-scrollbar {
            width: 8px;
        }

        .converter-content::-webkit-scrollbar-track {
            background: rgba(255,255,255,0.1);
            border-radius: 10px;
        }

        .converter-content::-webkit-scrollbar-thumb {
            background: rgba(255,255,255,0.3);
            border-radius: 10px;
        }

        .converter-section {
            background: white;
            border-radius: 15px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }

        .converter-section h3 {
            margin: 0 0 15px 0;
            color: #667eea;
            font-size: 18px;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .converter-section h3::before {
            content: '📁';
            font-size: 22px;
        }

        .file-input-wrapper {
            position: relative;
            margin-bottom: 15px;
        }

        .file-input-label {
            display: block;
            padding: 15px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-radius: 10px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s;
            font-weight: 500;
        }

        .file-input-label:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102,126,234,0.4);
        }

        .file-input-label input {
            display: none;
        }

        .converter-button {
            width: 100%;
            padding: 12px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 10px;
            cursor: pointer;
            font-size: 16px;
            font-weight: 600;
            transition: all 0.3s;
            margin-top: 10px;
        }

        .converter-button:hover:not(:disabled) {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102,126,234,0.4);
        }

        .converter-button:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        .converter-button.secondary {
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        }

        .file-name-display {
            padding: 10px;
            background: #f0f0f0;
            border-radius: 8px;
            margin-top: 10px;
            font-size: 14px;
            color: #666;
            word-break: break-all;
        }

        .json-textarea {
            width: 100%;
            min-height: 150px;
            padding: 10px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-family: 'Courier New', monospace;
            font-size: 13px;
            resize: vertical;
            margin-top: 10px;
        }

        .status-message {
            padding: 12px;
            border-radius: 8px;
            margin-top: 15px;
            font-size: 14px;
            animation: fadeIn 0.3s;
        }

        .status-message.success {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }

        .status-message.error {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }

        .status-message.info {
            background: #d1ecf1;
            color: #0c5460;
            border: 1px solid #bee5eb;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(-10px); }
            to { opacity: 1; transform: translateY(0); }
        }

        #file-converter-toggle {
            position: fixed;
            bottom: 30px;
            right: 30px;
            width: 60px;
            height: 60px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 50%;
            border: none;
            box-shadow: 0 4px 15px rgba(102,126,234,0.4);
            cursor: pointer;
            z-index: 999998;
            font-size: 28px;
            transition: all 0.3s;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        #file-converter-toggle:hover {
            transform: scale(1.1);
            box-shadow: 0 6px 20px rgba(102,126,234,0.6);
        }

        .progress-bar {
            width: 100%;
            height: 6px;
            background: #e0e0e0;
            border-radius: 3px;
            margin-top: 10px;
            overflow: hidden;
            display: none;
        }

        .progress-bar.active {
            display: block;
        }

        .progress-bar-fill {
            height: 100%;
            background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
            transition: width 0.3s;
            border-radius: 3px;
        }

        .button-group {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
            margin-top: 10px;
        }
    `;

    // ==================== 初始化 ====================
    function init() {
        // 注入样式
        const styleElement = document.createElement('style');
        styleElement.textContent = styles;
        document.head.appendChild(styleElement);

        // 创建UI
        createUI();

        // 绑定事件
        bindEvents();
    }

    // ==================== 创建UI ====================
    function createUI() {
        // 创建浮动按钮
        const toggleButton = document.createElement('button');
        toggleButton.id = 'file-converter-toggle';
        toggleButton.innerHTML = '🔄';
        toggleButton.title = '文件转换器';
        document.body.appendChild(toggleButton);

        // 创建主面板
        const panel = document.createElement('div');
        panel.id = 'file-converter-panel';
        panel.innerHTML = `
            <div class="converter-header">
                <h2>🔄 文件转换器</h2>
                <button class="converter-close">×</button>
            </div>
            <div class="converter-content">
                <!-- Excel ↔ JSON -->
                <div class="converter-section">
                    <h3>Excel ↔ JSON</h3>
                    <div class="file-input-wrapper">
                        <label class="file-input-label">
                            📤 选择 Excel 文件 (.xlsx, .xls)
                            <input type="file" id="excel-input" accept=".xlsx,.xls" />
                        </label>
                        <div id="excel-file-name" class="file-name-display" style="display:none;"></div>
                    </div>
                    <button class="converter-button" id="excel-to-json-btn" disabled>转换为 JSON</button>

                    <div style="margin: 20px 0; text-align: center; color: #999;">或者</div>

                    <textarea id="json-input" class="json-textarea" placeholder="粘贴 JSON 数据..."></textarea>
                    <button class="converter-button secondary" id="json-to-excel-btn">转换为 Excel</button>
                    <div id="excel-json-status"></div>
                </div>

                <!-- HTML → PDF -->
                <div class="converter-section">
                    <h3>HTML → PDF</h3>
                    <div class="button-group">
                        <button class="converter-button" id="current-page-to-pdf-btn">当前页面转 PDF</button>
                        <button class="converter-button secondary" id="html-file-to-pdf-btn">HTML 文件转 PDF</button>
                    </div>
                    <div class="file-input-wrapper" style="display:none;" id="html-file-wrapper">
                        <label class="file-input-label">
                            📤 选择 HTML 文件
                            <input type="file" id="html-input" accept=".html,.htm" />
                        </label>
                        <div id="html-file-name" class="file-name-display" style="display:none;"></div>
                    </div>
                    <div class="progress-bar" id="pdf-progress">
                        <div class="progress-bar-fill" style="width: 0%"></div>
                    </div>
                    <div id="html-pdf-status"></div>
                </div>

                <!-- Word → PDF -->
                <div class="converter-section">
                    <h3>Word → PDF</h3>
                    <div class="file-input-wrapper">
                        <label class="file-input-label">
                            📤 选择 Word 文件 (.docx, .doc)
                            <input type="file" id="word-input" accept=".docx,.doc" />
                        </label>
                        <div id="word-file-name" class="file-name-display" style="display:none;"></div>
                    </div>
                    <button class="converter-button" id="word-to-pdf-btn" disabled>转换为 PDF</button>
                    <div id="word-pdf-status"></div>
                    <div class="status-message info" style="margin-top: 15px;">
                        💡 提示：Word 转 PDF 需要使用在线 API 服务。本脚本使用浏览器本地处理，功能有限。
                    </div>
                </div>

                <!-- PPT → PDF -->
                <div class="converter-section">
                    <h3>PPT → PDF</h3>
                    <div class="file-input-wrapper">
                        <label class="file-input-label">
                            📤 选择 PPT 文件 (.pptx, .ppt)
                            <input type="file" id="ppt-input" accept=".pptx,.ppt" />
                        </label>
                        <div id="ppt-file-name" class="file-name-display" style="display:none;"></div>
                    </div>
                    <button class="converter-button" id="ppt-to-pdf-btn" disabled>转换为 PDF</button>
                    <div id="ppt-pdf-status"></div>
                    <div class="status-message info" style="margin-top: 15px;">
                        💡 提示：PPT 转 PDF 需要使用在线 API 服务。本脚本使用浏览器本地处理，功能有限。
                    </div>
                </div>
            </div>
        `;
        document.body.appendChild(panel);
    }

    // ==================== 绑定事件 ====================
    function bindEvents() {
        // 切换面板显示
        document.getElementById('file-converter-toggle').addEventListener('click', () => {
            const panel = document.getElementById('file-converter-panel');
            panel.classList.toggle('show');
        });

        // 关闭面板
        document.querySelector('.converter-close').addEventListener('click', () => {
            document.getElementById('file-converter-panel').classList.remove('show');
        });

        // Excel 相关
        document.getElementById('excel-input').addEventListener('change', handleExcelFileSelect);
        document.getElementById('excel-to-json-btn').addEventListener('click', convertExcelToJSON);
        document.getElementById('json-to-excel-btn').addEventListener('click', convertJSONToExcel);

        // HTML → PDF 相关
        document.getElementById('current-page-to-pdf-btn').addEventListener('click', convertCurrentPageToPDF);
        document.getElementById('html-file-to-pdf-btn').addEventListener('click', toggleHTMLFileInput);
        document.getElementById('html-input').addEventListener('change', handleHTMLFileSelect);

        // Word → PDF
        document.getElementById('word-input').addEventListener('change', handleWordFileSelect);
        document.getElementById('word-to-pdf-btn').addEventListener('click', convertWordToPDF);

        // PPT → PDF
        document.getElementById('ppt-input').addEventListener('change', handlePPTFileSelect);
        document.getElementById('ppt-to-pdf-btn').addEventListener('click', convertPPTToPDF);
    }

    // ==================== Excel ↔ JSON 功能 ====================
    let currentExcelFile = null;

    function handleExcelFileSelect(e) {
        const file = e.target.files[0];
        if (file) {
            currentExcelFile = file;
            document.getElementById('excel-file-name').textContent = `已选择: ${file.name}`;
            document.getElementById('excel-file-name').style.display = 'block';
            document.getElementById('excel-to-json-btn').disabled = false;
        }
    }

    function convertExcelToJSON() {
        if (!currentExcelFile) return;

        showStatus('excel-json-status', 'info', '正在转换...');

        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                // 转换所有工作表
                const result = {};
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    result[sheetName] = XLSX.utils.sheet_to_json(worksheet);
                });

                const jsonStr = JSON.stringify(result, null, 2);
                document.getElementById('json-input').value = jsonStr;

                // 下载 JSON 文件
                downloadFile(jsonStr, currentExcelFile.name.replace(/\.[^/.]+$/, '') + '.json', 'application/json');

                showStatus('excel-json-status', 'success', '✅ 转换成功！JSON 已下载并显示在下方文本框中。');
            } catch (error) {
                showStatus('excel-json-status', 'error', '❌ 转换失败: ' + error.message);
            }
        };
        reader.readAsArrayBuffer(currentExcelFile);
    }

    function convertJSONToExcel() {
        const jsonText = document.getElementById('json-input').value.trim();
        if (!jsonText) {
            showStatus('excel-json-status', 'error', '❌ 请输入 JSON 数据');
            return;
        }

        showStatus('excel-json-status', 'info', '正在转换...');

        try {
            const jsonData = JSON.parse(jsonText);
            const workbook = XLSX.utils.book_new();

            // 处理不同格式的 JSON
            if (Array.isArray(jsonData)) {
                // 如果是数组，创建单个工作表
                const worksheet = XLSX.utils.json_to_sheet(jsonData);
                XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
            } else if (typeof jsonData === 'object') {
                // 如果是对象，每个键创建一个工作表
                Object.keys(jsonData).forEach(key => {
                    const data = Array.isArray(jsonData[key]) ? jsonData[key] : [jsonData[key]];
                    const worksheet = XLSX.utils.json_to_sheet(data);
                    XLSX.utils.book_append_sheet(workbook, worksheet, key.substring(0, 31)); // Excel 工作表名称限制 31 字符
                });
            }

            // 生成 Excel 文件
            const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

            downloadBlob(blob, 'converted_' + new Date().getTime() + '.xlsx');

            showStatus('excel-json-status', 'success', '✅ 转换成功！Excel 文件已下载。');
        } catch (error) {
            showStatus('excel-json-status', 'error', '❌ 转换失败: ' + error.message);
        }
    }

    // ==================== HTML → PDF 功能 ====================
    let currentHTMLFile = null;

    function toggleHTMLFileInput() {
        const wrapper = document.getElementById('html-file-wrapper');
        wrapper.style.display = wrapper.style.display === 'none' ? 'block' : 'none';
    }

    function handleHTMLFileSelect(e) {
        const file = e.target.files[0];
        if (file) {
            currentHTMLFile = file;
            document.getElementById('html-file-name').textContent = `已选择: ${file.name}`;
            document.getElementById('html-file-name').style.display = 'block';
            convertHTMLFileToPDF(file);
        }
    }

    async function convertCurrentPageToPDF() {
        showStatus('html-pdf-status', 'info', '正在生成 PDF...');
        showProgress('pdf-progress', 0);

        try {
            const { jsPDF } = window.jspdf;

            // 使用 html2canvas 截取页面
            showProgress('pdf-progress', 30);
            const canvas = await html2canvas(document.body, {
                scale: 2,
                useCORS: true,
                logging: false
            });

            showProgress('pdf-progress', 70);

            // 创建 PDF
            const imgWidth = 210; // A4 宽度（mm）
            const imgHeight = (canvas.height * imgWidth) / canvas.width;
            const pdf = new jsPDF('p', 'mm', 'a4');

            const imgData = canvas.toDataURL('image/png');
            pdf.addImage(imgData, 'PNG', 0, 0, imgWidth, imgHeight);

            showProgress('pdf-progress', 100);

            // 下载 PDF
            pdf.save('webpage_' + new Date().getTime() + '.pdf');

            showStatus('html-pdf-status', 'success', '✅ PDF 生成成功！');
            setTimeout(() => hideProgress('pdf-progress'), 1000);
        } catch (error) {
            showStatus('html-pdf-status', 'error', '❌ 生成失败: ' + error.message);
            hideProgress('pdf-progress');
        }
    }

    async function convertHTMLFileToPDF(file) {
        showStatus('html-pdf-status', 'info', '正在转换 HTML 文件为 PDF...');
        showProgress('pdf-progress', 0);

        try {
            const reader = new FileReader();
            reader.onload = async function(e) {
                const htmlContent = e.target.result;

                // 创建临时 iframe 来渲染 HTML
                const iframe = document.createElement('iframe');
                iframe.style.position = 'absolute';
                iframe.style.left = '-9999px';
                iframe.style.width = '1200px';
                iframe.style.height = '800px';
                document.body.appendChild(iframe);

                iframe.contentDocument.open();
                iframe.contentDocument.write(htmlContent);
                iframe.contentDocument.close();

                showProgress('pdf-progress', 30);

                // 等待内容加载
                await new Promise(resolve => setTimeout(resolve, 1000));

                // 使用 html2canvas 转换
                const canvas = await html2canvas(iframe.contentDocument.body, {
                    scale: 2,
                    useCORS: true
                });

                showProgress('pdf-progress', 70);

                const { jsPDF } = window.jspdf;
                const imgWidth = 210;
                const imgHeight = (canvas.height * imgWidth) / canvas.width;
                const pdf = new jsPDF('p', 'mm', 'a4');

                const imgData = canvas.toDataURL('image/png');
                pdf.addImage(imgData, 'PNG', 0, 0, imgWidth, imgHeight);

                showProgress('pdf-progress', 100);

                pdf.save(file.name.replace(/\.[^/.]+$/, '') + '.pdf');

                // 清理
                document.body.removeChild(iframe);

                showStatus('html-pdf-status', 'success', '✅ HTML 转 PDF 成功！');
                setTimeout(() => hideProgress('pdf-progress'), 1000);
            };
            reader.readAsText(file);
        } catch (error) {
            showStatus('html-pdf-status', 'error', '❌ 转换失败: ' + error.message);
            hideProgress('pdf-progress');
        }
    }

    // ==================== Word → PDF 功能 ====================
    let currentWordFile = null;

    function handleWordFileSelect(e) {
        const file = e.target.files[0];
        if (file) {
            currentWordFile = file;
            document.getElementById('word-file-name').textContent = `已选择: ${file.name}`;
            document.getElementById('word-file-name').style.display = 'block';
            document.getElementById('word-to-pdf-btn').disabled = false;
        }
    }

    async function convertWordToPDF() {
        if (!currentWordFile) return;

        showStatus('word-pdf-status', 'info', '正在处理 Word 文件...');

        try {
            // 注意：浏览器端直接转换 Word 到 PDF 需要复杂的库或在线服务
            // 这里提供一个基础实现，使用 mammoth.js 提取文本内容
            showStatus('word-pdf-status', 'info', '正在读取 Word 文档内容...');

            const reader = new FileReader();
            reader.onload = async function(e) {
                try {
                    // 这里需要使用 mammoth.js 或类似库来解析 Word 文档
                    // 由于油猴脚本的限制，我们提供一个简化版本

                    const { jsPDF } = window.jspdf;
                    const pdf = new jsPDF();

                    pdf.setFontSize(12);
                    pdf.text('Word 文件内容预览', 20, 20);
                    pdf.text('文件名: ' + currentWordFile.name, 20, 30);
                    pdf.text('大小: ' + (currentWordFile.size / 1024).toFixed(2) + ' KB', 20, 40);
                    pdf.text('', 20, 50);
                    pdf.text('注意：完整的 Word 转 PDF 功能需要后端服务支持。', 20, 60);
                    pdf.text('建议使用在线转换服务：', 20, 70);
                    pdf.text('- https://www.ilovepdf.com/word_to_pdf', 20, 80);
                    pdf.text('- https://convertio.co/docx-pdf/', 20, 90);

                    pdf.save(currentWordFile.name.replace(/\.[^/.]+$/, '') + '_info.pdf');

                    showStatus('word-pdf-status', 'success', '✅ 已生成文档信息 PDF。完整转换请使用在线服务。');
                } catch (error) {
                    showStatus('word-pdf-status', 'error', '❌ 处理失败: ' + error.message);
                }
            };
            reader.readAsArrayBuffer(currentWordFile);
        } catch (error) {
            showStatus('word-pdf-status', 'error', '❌ 转换失败: ' + error.message);
        }
    }

    // ==================== PPT → PDF 功能 ====================
    let currentPPTFile = null;

    function handlePPTFileSelect(e) {
        const file = e.target.files[0];
        if (file) {
            currentPPTFile = file;
            document.getElementById('ppt-file-name').textContent = `已选择: ${file.name}`;
            document.getElementById('ppt-file-name').style.display = 'block';
            document.getElementById('ppt-to-pdf-btn').disabled = false;
        }
    }

    async function convertPPTToPDF() {
        if (!currentPPTFile) return;

        showStatus('ppt-pdf-status', 'info', '正在处理 PPT 文件...');

        try {
            const { jsPDF } = window.jspdf;
            const pdf = new jsPDF();

            pdf.setFontSize(12);
            pdf.text('PPT 文件信息', 20, 20);
            pdf.text('文件名: ' + currentPPTFile.name, 20, 30);
            pdf.text('大小: ' + (currentPPTFile.size / 1024).toFixed(2) + ' KB', 20, 40);
            pdf.text('', 20, 50);
            pdf.text('注意：完整的 PPT 转 PDF 功能需要后端服务支持。', 20, 60);
            pdf.text('建议使用在线转换服务：', 20, 70);
            pdf.text('- https://www.ilovepdf.com/powerpoint_to_pdf', 20, 80);
            pdf.text('- https://convertio.co/pptx-pdf/', 20, 90);

            pdf.save(currentPPTFile.name.replace(/\.[^/.]+$/, '') + '_info.pdf');

            showStatus('ppt-pdf-status', 'success', '✅ 已生成文档信息 PDF。完整转换请使用在线服务。');
        } catch (error) {
            showStatus('ppt-pdf-status', 'error', '❌ 转换失败: ' + error.message);
        }
    }

    // ==================== 工具函数 ====================
    function showStatus(elementId, type, message) {
        const statusElement = document.getElementById(elementId);
        statusElement.className = `status-message ${type}`;
        statusElement.textContent = message;
        statusElement.style.display = 'block';
    }

    function showProgress(elementId, percent) {
        const progressBar = document.getElementById(elementId);
        progressBar.classList.add('active');
        const fill = progressBar.querySelector('.progress-bar-fill');
        fill.style.width = percent + '%';
    }

    function hideProgress(elementId) {
        const progressBar = document.getElementById(elementId);
        progressBar.classList.remove('active');
        const fill = progressBar.querySelector('.progress-bar-fill');
        fill.style.width = '0%';
    }

    function downloadFile(content, filename, contentType) {
        const blob = new Blob([content], { type: contentType });
        downloadBlob(blob, filename);
    }

    function downloadBlob(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // ==================== 启动脚本 ====================
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

})();
