# 🔄 多功能文件转换器 - 油猴脚本

一个功能强大的浏览器端文件转换工具，支持多种文件格式之间的相互转换。

## ✨ 功能特性

### 📊 Excel ↔ JSON
- ✅ **Excel 转 JSON**：支持 `.xlsx` 和 `.xls` 格式
  - 自动转换所有工作表
  - 生成格式化的 JSON 文件
  - 支持多工作表导出

- ✅ **JSON 转 Excel**：
  - 支持数组格式的 JSON
  - 支持对象格式的 JSON（每个键生成一个工作表）
  - 自动生成 `.xlsx` 文件

### 📄 HTML → PDF
- ✅ **当前页面转 PDF**：一键将当前浏览的网页转换为 PDF
- ✅ **HTML 文件转 PDF**：上传本地 HTML 文件并转换为 PDF
- 高质量渲染，保持页面样式

### 📝 Word → PDF
- ⚠️ **基础支持**：生成文档信息 PDF
- 💡 提供在线转换服务推荐
- 适合查看文档基本信息

### 📊 PPT → PDF
- ⚠️ **基础支持**：生成文档信息 PDF
- 💡 提供在线转换服务推荐
- 适合查看文档基本信息

## 🚀 安装步骤

### 1. 安装油猴扩展

首先需要在浏览器中安装油猴（Tampermonkey）扩展：

- **Chrome/Edge**: [Chrome 网上应用店](https://chrome.google.com/webstore/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
- **Firefox**: [Firefox 附加组件](https://addons.mozilla.org/firefox/addon/tampermonkey/)
- **Safari**: [App Store](https://apps.apple.com/app/tampermonkey/id1482490089)
- **Opera**: [Opera 扩展](https://addons.opera.com/extensions/details/tampermonkey-beta/)

### 2. 安装脚本

1. 点击油猴扩展图标
2. 选择 "管理面板"
3. 点击 "+" 号创建新脚本
4. 将 `file-converter.user.js` 的内容复制粘贴进去
5. 按 `Ctrl+S` (或 `Cmd+S`) 保存

或者直接打开 `file-converter.user.js` 文件，油猴会自动识别并提示安装。

## 📖 使用方法

### 启动转换器

安装脚本后，在任何网页右下角会出现一个紫色的 🔄 浮动按钮，点击即可打开文件转换器面板。

### Excel ↔ JSON 转换

#### Excel 转 JSON：
1. 点击 "选择 Excel 文件" 按钮
2. 选择你的 `.xlsx` 或 `.xls` 文件
3. 点击 "转换为 JSON" 按钮
4. JSON 文件会自动下载，同时显示在文本框中

#### JSON 转 Excel：
1. 在文本框中粘贴或输入 JSON 数据
2. 点击 "转换为 Excel" 按钮
3. Excel 文件会自动下载

**支持的 JSON 格式：**

```json
// 数组格式（生成单个工作表）
[
  { "姓名": "张三", "年龄": 25 },
  { "姓名": "李四", "年龄": 30 }
]

// 对象格式（每个键生成一个工作表）
{
  "员工": [
    { "姓名": "张三", "部门": "技术部" }
  ],
  "部门": [
    { "名称": "技术部", "人数": 10 }
  ]
}
```

### HTML → PDF 转换

#### 当前页面转 PDF：
1. 浏览到你想要转换的网页
2. 打开文件转换器
3. 点击 "当前页面转 PDF" 按钮
4. 等待处理完成，PDF 会自动下载

#### HTML 文件转 PDF：
1. 点击 "HTML 文件转 PDF" 按钮
2. 选择你的 `.html` 文件
3. 等待处理完成，PDF 会自动下载

### Word / PPT 转 PDF

1. 选择对应的文件
2. 点击转换按钮
3. 会生成一个包含文件信息的 PDF

**注意**：由于浏览器限制，完整的 Word/PPT 转 PDF 功能需要后端服务支持。脚本会提供推荐的在线转换服务：

- [ILovePDF](https://www.ilovepdf.com/)
- [Convertio](https://convertio.co/)

## 🛠️ 技术栈

- **SheetJS (xlsx)**: Excel 文件处理
- **jsPDF**: PDF 生成
- **html2canvas**: HTML 到图像转换
- **原生 JavaScript**: 核心逻辑

## 📋 依赖库

脚本自动从 CDN 加载以下库：

```javascript
@require https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
@require https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js
@require https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js
```

## ⚙️ 高级配置

### 自定义匹配规则

默认情况下，脚本在所有网页上运行（`@match *://*/*`）。你可以修改脚本头部的 `@match` 规则来限制运行范围：

```javascript
// 只在特定网站运行
// @match https://example.com/*

// 在多个网站运行
// @match https://example.com/*
// @match https://another-site.com/*
```

### 修改界面位置

在脚本中找到以下 CSS 代码来调整浮动按钮位置：

```css
#file-converter-toggle {
    bottom: 30px;  /* 距离底部距离 */
    right: 30px;   /* 距离右侧距离 */
}
```

## 🎨 界面预览

脚本提供了美观的渐变紫色主题界面，包含：

- 🎯 浮动切换按钮
- 📱 响应式面板设计
- ✨ 平滑动画效果
- 📊 实时进度显示
- 💡 状态提示信息

## 🐛 常见问题

### 1. 脚本没有显示浮动按钮？
- 确认油猴扩展已启用
- 检查脚本是否已安装并启用
- 刷新页面

### 2. Excel 转换失败？
- 确认文件格式是 `.xlsx` 或 `.xls`
- 检查文件是否损坏
- 确认文件不是受密码保护的

### 3. PDF 生成质量不好？
- HTML 转 PDF 使用的是截图方式，质量取决于页面渲染
- 复杂页面可能需要更长加载时间
- 建议使用专业 PDF 工具处理复杂文档

### 4. Word/PPT 转换功能有限？
- 浏览器环境限制了完整的文档处理能力
- 建议使用推荐的在线服务进行完整转换
- 或使用桌面软件（如 Microsoft Office、LibreOffice）

## 🔒 隐私说明

- ✅ 所有转换操作都在**浏览器本地**完成
- ✅ **不上传**任何文件到服务器
- ✅ **不收集**任何用户数据
- ✅ 完全**离线**运行（除了加载 CDN 库）

## 📝 版本历史

### v1.0.0 (2024)
- ✨ 初始版本发布
- ✅ 支持 Excel ↔ JSON 转换
- ✅ 支持 HTML → PDF 转换
- ✅ 基础 Word/PPT 信息导出

## 🤝 贡献

欢迎提交问题和改进建议！

## 📄 许可证

MIT License

## 🙏 致谢

感谢以下开源项目：
- [SheetJS](https://github.com/SheetJS/sheetjs)
- [jsPDF](https://github.com/parallax/jsPDF)
- [html2canvas](https://github.com/niklasvh/html2canvas)

---

**享受文件转换的便利！** 🎉
