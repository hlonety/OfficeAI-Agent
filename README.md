# Office AI Agent (Excel & Word)

这是一个基于大模型（DeepSeek, 智谱AI, Qwen）的 Office 智能助手，允许你通过自然语言直接操控 Excel 和 Word。

🌐 **在线演示**: [GitHub Pages](https://hlonety.github.io/OfficeAI-Agent/)

---

## ✨ 核心功能

*   **智能 Excel 操作**：
    *   **公式优先**：自动生成 Excel 公式进行计算，而非死板的数字。
    *   **智能防撞**：自动检测数据碰撞，防止覆盖已有数据。
    *   **图表生成**：一句话生成柱状图、折线图等。
    *   **格式美化**：自动应用金融级配色（蓝色输入值、黑色公式）。
    *   **上下文记忆**：支持多轮对话，AI 记得你之前的指令和修正。
    *   **选区感知**：直接选中单元格或区域，AI 自动读取并理解选中内容。
*   **多模型支持**：
    *   支持 DeepSeek (V3/R1), 智谱 GLM-4, 通义千问。
    *   支持自定义 OpenAI 格式的 API (如本地 NAS 模型)。

---

## 📥 安装指南 (如何分享给朋友)

本插件采用 "Sideloading" (侧加载) 方式安装，无需通过应用商店。

### 方式一：Windows Excel (推荐)
1.  下载本项目中的 [manifest.xml](https://github.com/hlonety/OfficeAI-Agent/raw/main/manifest.xml) 文件。
2.  打开 Excel 桌面版或 [Excel 网页版](https://www.office.com/launch/excel)。
3.  点击菜单栏 **"插入" (Insert)** -> **"获取加载项" (Get Add-ins)**。
4.  点击 **"我的加载项" (My Add-ins)** -> **"上传我的加载项" (Upload My Add-in)**。
5.  选择下载的 `manifest.xml` 文件即可。

### 方式二：Mac Excel
1.  下载 [manifest.xml](https://github.com/hlonety/OfficeAI-Agent/raw/main/manifest.xml)。
2.  打开 Finder，按下 `Cmd + Shift + G`，输入以下路径：
    ```text
    ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
    ```
    *(如果 `wef` 文件夹不存在，请手动创建一个)*
3.  将 XML 文件放入该文件夹。
4.  重启 Excel，在 **"插入"** -> **"我的加载项"** 下拉菜单中即可找到。

---

## ❓ 常见问题

### 1. 更新了代码需要重新安装吗？
*   **不需要**：如果是功能的更新（比如 AI 变聪明了、界面变好看了），你只需要重新打开插件或点击"刷新"即可，它会自动加载 GitHub 上最新的代码。
*   **需要**：只有当 **菜单按钮** 发生变化（比如增加了一个新按钮）或者修改了图标时，才需要重新上传 `manifest.xml`。

### 2. 我需要一直开着开发者模式吗？
不需要。安装一次后，它通常会保留在你的 Excel 中（"我的加载项" -> "开发者加载项"）。你像使用普通插件一样使用即可。

### 3. 数据安全吗？
非常安全。
*   这是一个纯前端插件，**没有后台服务器**存储你的数据。
*   你的 API Key 仅保存在你**本地浏览器的缓存**中，不会上传到任何地方。

---

## 🛠️ 开发者说明

如果你想自己修改代码：
1.  克隆仓库：`git clone https://github.com/hlonety/OfficeAI-Agent.git`
2.  安装依赖：`npm install`
3.  本地启动：`npm run dev` (需要修改 manifest.xml 指向 localhost)

## 📅 更新日志
*   **v2026.01.26**:
    *   [新增] **上下文记忆 (Memory)**: 解决了"失忆"问题，支持连续对话。
    *   [新增] **选中区域感知 (Selection Awareness)**: AI 可直接操作鼠标选中的单元格。
    *   [优化] **ReAct 引擎**: 增强了 DeepSeek 和 GLM 在复杂推理任务中的执行力。
    *   [修复] **日期逻辑**: 修复了 GLM 模型无法正确理解相对日期（如"上周"）的问题。
