/*
 * Office AI Agent - 主逻辑
 */

/* global console, document, Excel, Office, Word, AiConnector, marked */

// 预设服务商配置
const PRESET_PROVIDERS = [
    {
        id: "deepseek",
        name: "DeepSeek",
        icon: "D",
        color: "#0066ff",
        baseUrl: "https://api.deepseek.com/v1",
        models: ["deepseek-chat", "deepseek-reasoner", "deepseek-coder"]
    },
    {
        id: "zhipu",
        name: "智谱 AI",
        icon: "智",
        color: "#1a56db",
        baseUrl: "https://open.bigmodel.cn/api/paas/v4",
        models: ["glm-4-flash", "glm-4-air", "glm-4-airx", "glm-4-long", "glm-4-plus", "glm-4-0520", "glm-4", "glm-4v", "glm-4v-plus"]
    },
    {
        id: "gemini",
        name: "Gemini",
        icon: "✦",
        color: "#4285f4",
        baseUrl: "https://generativelanguage.googleapis.com/v1beta/openai",
        models: ["gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-1.5-pro", "gemini-1.5-flash"]
    },
    {
        id: "qwen",
        name: "通义千问",
        icon: "通",
        color: "#6366f1",
        baseUrl: "https://dashscope.aliyuncs.com/compatible-mode/v1",
        models: ["qwen-turbo", "qwen-plus", "qwen-max", "qwen-long"]
    },
    {
        id: "openai",
        name: "OpenAI",
        icon: "◎",
        color: "#10a37f",
        baseUrl: "https://api.openai.com/v1",
        models: ["gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "o1", "o1-mini", "o3-mini"]
    }
];

// 配置存储
let providerConfigs = {}; // { providerId: { apiKey, models?: [] } }
let customProviders = [];
let activeProvider = null;
let activeModel = null;

// 全局对话历史
let globalChatHistory = [];
const MAX_HISTORY_LENGTH = 20; // 限制上下文长度

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel || info.host === Office.HostType.Word) {
        loadConfig();
        renderProviderList();
        updateModelSelector();
        bindEvents();
    }
});

function bindEvents() {
    document.getElementById("run-btn").onclick = run;
    document.getElementById("back-btn").onclick = hideSettings;
    document.getElementById("add-custom-btn").onclick = showCustomModal;
    document.getElementById("cancel-custom-btn").onclick = hideCustomModal;
    document.getElementById("save-custom-btn").onclick = saveCustomProvider;
    document.getElementById("active-model-select").onchange = onModelChange;

    // 新增：直接点击进入配置
    document.getElementById("settings-direct-btn").onclick = () => {
        showSettings();
    };

    /* 移除旧的下拉菜单逻辑
    document.getElementById("menu-toggle-btn").onclick = ...
    document.addEventListener("click", ...
    document.getElementById("menu-settings-btn").onclick = ...
    document.getElementById("menu-about-btn").onclick = ...
    */

    // 新建聊天按钮
    document.getElementById("new-chat-btn").onclick = () => {
        const chatHistory = document.getElementById("chat-history");
        chatHistory.innerHTML = `
            <div class="message system-message">
                <i class="fa-regular fa-lightbulb"></i><br />
                新对话已开始！有什么可以帮你的？
            </div>
        `;
    };

    document.getElementById("user-input").addEventListener("keydown", (e) => {
        if (e.ctrlKey && e.key === "Enter") run();
    });
}

// ===== 配置存储 =====
function loadConfig() {
    try {
        const saved = localStorage.getItem("office-ai-providers");
        if (saved) {
            const data = JSON.parse(saved);
            providerConfigs = data.configs || {};
            customProviders = data.custom || [];
            activeProvider = data.activeProvider || null;
            activeModel = data.activeModel || null;
        }
    } catch (e) {
        console.warn("Load config failed:", e);
    }
}

function saveConfig() {
    localStorage.setItem("office-ai-providers", JSON.stringify({
        configs: providerConfigs,
        custom: customProviders,
        activeProvider,
        activeModel
    }));
}

// ===== 渲染服务商列表 =====
function renderProviderList() {
    const container = document.getElementById("provider-list");
    container.innerHTML = "";

    PRESET_PROVIDERS.forEach(p => {
        const config = providerConfigs[p.id];
        const isConfigured = config && config.apiKey;
        container.appendChild(createProviderCard(p, isConfigured, false));
    });

    customProviders.forEach(p => {
        const isConfigured = !!p.apiKey;
        container.appendChild(createProviderCard(p, isConfigured, true));
    });
}

function createProviderCard(provider, isConfigured, isCustom) {
    const card = document.createElement("div");
    card.className = "provider-card";
    card.dataset.id = provider.id;

    card.innerHTML = `
        <div class="provider-icon" style="background: ${provider.color || '#666'}">
            ${provider.icon || provider.name.charAt(0)}
        </div>
        <div class="provider-info">
            <div class="provider-name">${provider.name}</div>
            <div class="provider-status ${isConfigured ? 'configured' : ''}">
                ${isConfigured ? '<i class="fa-solid fa-check"></i> 已配置' : '<i class="fa-solid fa-key"></i> 未配置'}
            </div>
            ${(provider.models && provider.models.length > 0) ? `<div style="font-size:10px;color:#999">${provider.enabledModels ? provider.enabledModels.length : provider.models.length} / ${provider.models.length} models</div>` : ''}
        </div>
        <div class="provider-actions">
            <button class="edit-btn" title="配置"><i class="fa-solid fa-pen"></i></button>
            ${isCustom ? '<button class="delete-btn" title="删除"><i class="fa-solid fa-trash"></i></button>' : ''}
        </div>
    `;

    card.querySelector(".edit-btn").onclick = () => toggleKeyInput(card, provider, isCustom);

    if (isCustom) {
        card.querySelector(".delete-btn").onclick = () => deleteCustomProvider(provider.id);
    }

    return card;
}

function toggleKeyInput(card, provider, isCustom) {
    const existing = card.querySelector(".key-input-row");
    if (existing) {
        existing.remove();
        return;
    }

    document.querySelectorAll(".key-input-row").forEach(el => el.remove());

    const config = isCustom ? provider : providerConfigs[provider.id];
    const currentKey = config?.apiKey || "";

    const row = document.createElement("div");
    row.className = "key-input-row";
    row.innerHTML = `
        <input type="password" placeholder="输入 API Key" value="${currentKey}" />
        <button class="ms-Button--primary save-key-btn">保存</button>
        <button class="refresh-models-btn" title="获取模型列表"><i class="fa-solid fa-rotate"></i></button>
    `;

    card.appendChild(row);

    const input = row.querySelector("input");
    input.focus();

    row.querySelector(".save-key-btn").onclick = () => {
        const key = input.value.trim();
        if (!key) return;

        if (isCustom) {
            const idx = customProviders.findIndex(p => p.id === provider.id);
            if (idx >= 0) customProviders[idx].apiKey = key;
        } else {
            if (!providerConfigs[provider.id]) providerConfigs[provider.id] = {};
            providerConfigs[provider.id].apiKey = key;
        }

        saveConfig();
        renderProviderList();
        updateModelSelector();
    };

    // 辅助函数：渲染模型筛选列表
    const renderFilterList = (models, enabledModels) => {
        // 注意：将列表添加到 card 而不是 row，以获得完整宽度
        let list = card.querySelector(".model-filter-list");
        if (!list) {
            list = document.createElement("div");
            list.className = "model-filter-list";
            card.appendChild(list);
        }
        list.innerHTML = "";

        models.forEach(m => {
            const div = document.createElement("div");
            div.className = "model-item";
            const isChecked = !enabledModels || enabledModels.includes(m);
            div.innerHTML = `<input type="checkbox" value="${m}" ${isChecked ? "checked" : ""}><label>${m}</label>`;

            div.querySelector("input").onchange = (e) => {
                const checked = e.target.checked;
                // 更新配置
                let targetConfig = isCustom ? customProviders.find(p => p.id === provider.id) : providerConfigs[provider.id];
                if (!targetConfig) return;

                // 初始化 enabledModels
                if (!targetConfig.enabledModels) targetConfig.enabledModels = [...targetConfig.models];

                if (checked) {
                    if (!targetConfig.enabledModels.includes(m)) targetConfig.enabledModels.push(m);
                } else {
                    targetConfig.enabledModels = targetConfig.enabledModels.filter(x => x !== m);
                }
                saveConfig();
                updateModelSelector();

                // 刷新卡片上的计数显示(简单重新渲染整个列表太重，这里只刷新选择器)
            };

            // 点击label也能触发
            div.querySelector("label").onclick = () => div.querySelector("input").click();

            list.appendChild(div);
        });
    };

    // 如果已有模型，直接显示列表
    const currentConfig = isCustom ? provider : providerConfigs[provider.id];
    if (currentConfig && currentConfig.models && currentConfig.models.length > 0) {
        renderFilterList(currentConfig.models, currentConfig.enabledModels);
    }

    // 获取模型列表
    row.querySelector(".refresh-models-btn").onclick = async () => {
        const key = input.value.trim();
        if (!key) {
            alert("请先输入 API Key");
            return;
        }

        const btn = row.querySelector(".refresh-models-btn");
        btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i>';

        try {
            const models = await AiConnector.fetchModels(key, provider.baseUrl);
            if (models.length > 0) {
                let targetConfig;
                if (isCustom) {
                    const idx = customProviders.findIndex(p => p.id === provider.id);
                    if (idx >= 0) {
                        customProviders[idx].models = models;
                        // 默认全选
                        customProviders[idx].enabledModels = [...models];
                        targetConfig = customProviders[idx];
                    }
                } else {
                    // 更新预设的模型列表
                    const preset = PRESET_PROVIDERS.find(p => p.id === provider.id);
                    if (preset) preset.models = models;
                    if (!providerConfigs[provider.id]) providerConfigs[provider.id] = {};
                    providerConfigs[provider.id].models = models;
                    providerConfigs[provider.id].enabledModels = [...models];
                    targetConfig = providerConfigs[provider.id];
                }
                saveConfig();
                updateModelSelector();

                // 显示筛选列表
                renderFilterList(models, targetConfig.enabledModels);

                // alert(`成功获取 ${models.length} 个模型`); // 不再弹窗打扰，直接显示列表
            } else {
                alert("未能获取模型列表，请手动输入模型名称");
            }
        } catch (e) {
            console.error(e);
            alert("获取失败: " + e.message);
        } finally {
            btn.innerHTML = '<i class="fa-solid fa-rotate"></i>';
        }
    };
}

function deleteCustomProvider(id) {
    // Office 环境下 confirm 有时会被拦截，改为二次确认按钮或者直接删除
    // 为了修复"无反应"的问题，这里先改为直接删除（或者我们可以做一个简单的自定义确认）
    // 但鉴于用户反馈"没有任何反应"，大概率是 confirm 被阻塞。

    // 简单的二次点击逻辑：检查按钮文本
    const btn = document.querySelector(`.provider-card[data-id="${id}"] .delete-btn`);
    if (btn) {
        if (!btn.dataset.confirm) {
            btn.dataset.confirm = "true";
            btn.innerHTML = '<i class="fa-solid fa-check"></i> 确定?';
            btn.style.background = "#d13438";
            btn.style.color = "white";
            setTimeout(() => {
                btn.dataset.confirm = "";
                btn.innerHTML = '<i class="fa-solid fa-trash"></i>';
                btn.style.background = "";
                btn.style.color = "";
            }, 3000);
            return;
        }
    }

    customProviders = customProviders.filter(p => p.id !== id);
    if (activeProvider === id) {
        activeProvider = null;
        activeModel = null;
    }
    saveConfig();
    renderProviderList();
    updateModelSelector();
}

// ===== 自定义服务商弹窗 =====
function showCustomModal() {
    document.getElementById("custom-modal").classList.remove("hidden");
    document.getElementById("custom-name").value = "";
    document.getElementById("custom-url").value = "";
    document.getElementById("custom-key").value = "";
}

function hideCustomModal() {
    document.getElementById("custom-modal").classList.add("hidden");
}

function saveCustomProvider() {
    const name = document.getElementById("custom-name").value.trim();
    const url = document.getElementById("custom-url").value.trim();
    const key = document.getElementById("custom-key").value.trim();

    if (!name || !url) {
        alert("请填写服务名称和地址");
        return;
    }

    const id = "custom_" + Date.now();
    customProviders.push({
        id,
        name,
        baseUrl: url,
        apiKey: key,
        icon: name.charAt(0).toUpperCase(),
        color: "#" + Math.floor(Math.random() * 16777215).toString(16).padStart(6, '0'),
        models: []
    });

    saveConfig();
    hideCustomModal();
    renderProviderList();
    updateModelSelector();
}

// ===== 模型选择器 =====
function updateModelSelector() {
    const select = document.getElementById("active-model-select");
    select.innerHTML = "";

    let hasOptions = false;

    PRESET_PROVIDERS.forEach(p => {
        const config = providerConfigs[p.id];
        if (config && config.apiKey) {
            const group = document.createElement("optgroup");
            group.label = p.name;

            // 优先使用动态获取的模型列表，并应用筛选
            let models = config.models || p.models;
            if (config.enabledModels) {
                models = models.filter(m => config.enabledModels.includes(m));
            }

            if (models.length > 0) {
                models.forEach(m => {
                    const opt = document.createElement("option");
                    opt.value = `${p.id}::${m}`;
                    opt.textContent = m;
                    if (activeProvider === p.id && activeModel === m) {
                        opt.selected = true;
                    }
                    group.appendChild(opt);
                });
                select.appendChild(group);
                hasOptions = true;
            }
        }
    });

    customProviders.forEach(p => {
        if (p.apiKey) {
            const group = document.createElement("optgroup");
            group.label = p.name;

            let models = p.models;
            if (p.enabledModels) {
                models = models.filter(m => p.enabledModels.includes(m));
            }
            // 只有当有模型时才显示组
            if (models && models.length > 0) {
                models.forEach(m => {
                    const opt = document.createElement("option");
                    opt.value = `${p.id}::${m}`;
                    opt.textContent = m;
                    if (activeProvider === p.id && activeModel === m) {
                        opt.selected = true;
                    }
                    group.appendChild(opt);
                });
                select.appendChild(group);
                hasOptions = true;
            }
        }
    });



    if (!hasOptions) {
        const opt = document.createElement("option");
        opt.value = "";
        opt.textContent = "请先配置 API Key";
        select.appendChild(opt);
    }
}

function onModelChange() {
    const val = document.getElementById("active-model-select").value;
    if (!val) return;

    const [providerId, model] = val.split("::");
    activeProvider = providerId;
    activeModel = model;
    saveConfig();
}

// ===== 视图切换 =====
function showSettings() {
    document.getElementById("chat-view").classList.add("hidden");
    document.getElementById("settings-view").classList.remove("hidden");
}

function hideSettings() {
    document.getElementById("settings-view").classList.add("hidden");
    document.getElementById("chat-view").classList.remove("hidden");
}

// ===== 聊天逻辑 (流式输出) =====
async function run() {
    const inputEl = document.getElementById("user-input");
    const message = inputEl.value.trim();

    if (!message) return;

    addUserMessage(message);
    inputEl.value = "";

    // 1. 添加用户消息到历史
    globalChatHistory.push({ role: "user", content: message });
    // 超长截断 (保留 SystemPrompt 实际上由 Connector 注入，这里只需管理 User/Assistant)
    if (globalChatHistory.length > MAX_HISTORY_LENGTH) {
        globalChatHistory = globalChatHistory.slice(globalChatHistory.length - MAX_HISTORY_LENGTH);
    }

    const val = document.getElementById("active-model-select").value;
    if (!val || val === "") {
        addSystemMessage("请先配置 API Key 并选择模型");
        showSettings();
        return;
    }

    const [providerId, model] = val.split("::");

    let provider = PRESET_PROVIDERS.find(p => p.id === providerId);
    let apiKey = providerConfigs[providerId]?.apiKey;
    let baseUrl = provider?.baseUrl;

    if (!provider) {
        provider = customProviders.find(p => p.id === providerId);
        if (provider) {
            apiKey = provider.apiKey;
            baseUrl = provider.baseUrl;
        }
    }

    if (!apiKey) {
        addSystemMessage("该服务商未配置 API Key");
        return;
    }

    // 创建 AI 消息容器
    const aiMsgDiv = document.createElement("div");
    aiMsgDiv.className = "message ai-message";

    // 思考块
    const thinkBlock = document.createElement("details");
    thinkBlock.className = "think-block";
    thinkBlock.style.display = "none";
    thinkBlock.innerHTML = `
        <summary><i class="fa-solid fa-brain"></i> <span class="think-label">思考中...</span></summary>
        <div class="think-content"></div>
    `;

    // 回复内容
    const contentDiv = document.createElement("div");
    contentDiv.className = "ai-content";
    contentDiv.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i>';

    aiMsgDiv.appendChild(thinkBlock);
    aiMsgDiv.appendChild(contentDiv);
    appendMessage(aiMsgDiv);

    let thinkStartTime = null;

    try {
        const connector = new AiConnector(apiKey, model, baseUrl);

        const hostType = Office.context.host === Office.HostType.Excel ? "Excel" : "Word";

        let systemPrompt = `你是一个精通 ${hostType} 的智能助手。请简洁地回答用户问题。\nCurrent Date: ${new Date().toLocaleDateString()}`;

        if (hostType === "Excel") {
            // --- Smart Context V1.0 ---
            let contextData = "";
            try {
                if (window.ExcelAgent) {
                    contextData = await window.ExcelAgent.getContextData();
                    console.log("Context loaded:", contextData);
                }
            } catch (e) {
                console.warn("Context load failed:", e);
            }

            systemPrompt += `
\n**重要 Excel 智能助手指令**
你不仅是一个聊天机器人，你是一个专业的 Excel 自动化助手。

*** 当前工作表上下文 ***
${contextData}
*** 上下文结束 ***

如果用户的请求需要对 Excel 进行操作（如创建表格、格式化、图表或设置数值），你**必须**严格返回以下 JSON 格式。

### 1. 公式优先策略 (关键)
- **永远不要硬编码计算结果**。必须使用 Excel 公式。
- 例如：不要自己计算 "Sum = 100" 然后写入 100，而要输出 "=SUM(A1:A10)"。
- 将假设/输入数据与计算逻辑分离开。

### 2. 样式规则 (行业标准)
- **输入值**：蓝色文本 (RGB 0,0,255)。
- **公式**：黑色文本 (RGB 0,0,0)。
- **外部链接**：红色文本 (RGB 255,0,0)。

### 3. JSON 协议
要执行操作，请输出一个由 \`\`\`json ... \`\`\` 包裹的单一 JSON 代码块。

Schema 示例:
\`\`\`json
{
  "thought": "简要解释你要做什么",
  "actions": [
    { 
      "type": "setCell", 
      "params": { "address": "B2", "value": "=SUM(A1:A10)", "formula": true }
    },
    { 
      "type": "formatRange", 
      "params": { "range": "B2", "style": "calculation" }
    },
    { 
      "type": "createTable", 
      "params": { "range": "A1:C5", "name": "SalesData" } 
    }
  ],
  "message": "最终展示给用户的回复"
}
\`\`\`

支持的动作类型 (Action Types):
**写入操作:**
- setCell (设置单元格), setRange (设置区域)
- createTable (创建表格)
- createChart (创建图表)
- formatRange (格式化, style: "input"|"calculation"|"header"|"external")
- autoFit (自动调整列宽)
- scanForErrors (扫描错误，无参数)
- fixError (修复错误, params: address, value)

**读取操作:**
- readRange (读取区域, params: address)
- getUsedRangeInfo (获取使用范围信息)
- findData (查找数据, params: keyword)

**重要提示**:
1. 如果信息不足（例如找不到列名），请先使用 \`readRange\` 或 \`getUsedRangeInfo\` 查看数据。
2. **严禁**输出多余的解释性文字，只输出符合 Schema 的 JSON。

### 4. ReAct 策略 (关键)
如果你不确定数据在哪里，**请先使用读取操作**，然后再写入。
工作流示例：
1. 用户问："统计保费列的总和"
2. 你不知道哪一列是"保费"，所以先调用 \`findData\` 查找它。
3. 我会执行读取操作并把观察结果告诉你。
4. 然后你再生成最终的写入操作（如使用 SUM 公式的 setCell）。

### 5. 重要：智能处理混合数据
当用户要求求和/计算数据时：
- **不要盲目**对整列应用公式。数据可能包含文本、标题或混合类型。
- **第一步**：调用 \`getUsedRangeInfo\` 获取实际数据范围。
- **第二步**：调用 \`readRange\` 读取样本（如前10行）以了解数据结构。
- **第三步**：根据观察结果：
  - 如果有清晰的数字列，只针对该列。
  - 如果是混合的，考虑使用 \`SUMIF\` 或 \`SUMPRODUCT\` 来筛选。
  - 或者识别包含数字的具体单元格地址。
- **示例**：用户说"求所有数字之和"。你读取样本看到数字在 C, D, H 列。使用 \`=SUM(C:C,D:D,H:H)\` 而不是 \`=SUM(A:Z)\`。

如果用户的请求只是一个问题，正常回答即可，不需要 JSON。
`;
        }

        let finalFullContent = "";

        await connector.chatStream(
            globalChatHistory, // 传入完整历史
            systemPrompt,
            // onChunk - 收到正常内容
            (chunk, fullContent) => {
                finalFullContent = fullContent;

                // 优化体验：从显示内容中移除 JSON 代码块，避免闪烁
                const jsonRegex = new RegExp("```json[\\s\\S]*?(?:```|$)", "g");
                let displayContent = fullContent.replace(jsonRegex, '').trim();

                const rawJsonRegex = new RegExp("\\{[\\s\\S]*?\"actions\"\\s*:\\s*\\[[\\s\\S]*?\\]\\s*\\}", "g");
                displayContent = displayContent.replace(rawJsonRegex, '').trim();

                if (typeof marked !== 'undefined') {
                    contentDiv.innerHTML = marked.parse(displayContent);
                } else {
                    contentDiv.textContent = displayContent;
                }

                // 自动滚动到底部
                scrollToBottom();
            },
            // onThinking - 收到思考过程
            () => {
                thinkStartTime = Date.now();
                thinkBlock.style.display = "block";
                contentDiv.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> 正在生成回复...';
            },
            // onThinkEnd - 思考结束
            (thinkContent, duration) => {
                const label = thinkBlock.querySelector(".think-label");
                label.textContent = `Thought for ${duration}s`;
                thinkBlock.querySelector(".think-content").textContent = thinkContent;
            }
        );

        // 如果没有内容，移除加载图标
        if (contentDiv.innerHTML.includes('fa-spinner') || !contentDiv.textContent.trim()) {
            if (!contentDiv.innerHTML.includes('❌')) {
                if (contentDiv.innerHTML.includes('fa-spinner')) contentDiv.innerHTML = '';
            }
        }

        // 核心逻辑：流结束后，检查是否有可执行的 JSON 指令
        // 核心逻辑：流结束后，检查是否有可执行的 JSON 指令
        // 使用增强的提取函数 (支持 Markdown 和 裸 JSON)
        let jsonStr = extractJsonFromContent(finalFullContent);
        let jsonMatch = jsonStr ? [null, jsonStr] : null;

        if (jsonMatch) {
            console.log("JSON extracted:", jsonStr.substring(0, 50) + "...");
        }

        if (jsonMatch) {
            try {
                const plan = JSON.parse(jsonMatch[1]);
                if (plan.actions && Array.isArray(plan.actions)) {
                    // 显示正在执行状态
                    const statusId = addSystemMessage('<i class="fa-solid fa-gear fa-spin"></i> 正在执行 Excel 操作...');

                    const result = await ExcelAgent.execute(plan);

                    removeMessage(statusId);

                    // 2. 将第一轮的 AI 完整回复 (含 JSON) 加入历史，确保 AI 记得自己干了什么
                    globalChatHistory.push({ role: "assistant", content: finalFullContent });

                    // ========= ReAct Loop: 处理观察结果 =========
                    if (result && result.observations && result.observations.length > 0) {
                        // 如果有观察结果 (来自读取操作), 注入并再次调用 AI
                        const observationText = result.observations.join("\n---\n");

                        // 隐藏第一轮的 JSON 输出 (用户不需要看到)
                        // 替换为简洁的状态信息
                        contentDiv.innerHTML = '<span style="color: var(--text-sub);"><i class="fa-solid fa-database"></i> 数据读取完成，正在分析...</span>';

                        console.log("ReAct Observation:", observationText);

                        // --- 自动化第二轮 (包含原始问题) ---
                        // 关键：强调必须输出 JSON 动作
                        const observationPrompt = `Original user request: "${message}"

I have read the spreadsheet data. Here is what I found:
${observationText}

IMPORTANT: Based on this data, you MUST now generate the JSON actions to fulfill the user's request.
Reflect on the data found (columns, rows, content) and map it to the actions logic (e.g., if finding sums, identify the correct column letter).

OUTPUT RULES:
1.  You must output a VALID JSON block with "actions".
2.  Do NOT just explain what you want to do.
3.  Do NOT ask for confirmation.
4.  Directly output the JSON to execute the change.

Example:
\`\`\`json
{
  "thought": "Found 'Price' in column B. Calculating sum.",
  "actions": [{"type": "setCell", "params": {"address": "B100", "value": "=SUM(B2:B99)", "formula": true}}],
  "message": "Calculated total price."
}
\`\`\`
`;

                        // ReAct 中间状态也需要加入历史吗？
                        // 为了让 AI 知道上下文，构建一个临时的 history 用于第二轮
                        // System (Data) -> AI (Action)
                        // 若直接 append 到 globalChatHistory，可能会显得啰嗦。
                        // 但为了 Robustness，我们把 "观察结果" 伪装成 User 消息告诉 AI
                        const reActHistory = [...globalChatHistory, { role: "user", content: observationPrompt }];

                        // 创建新的 AI 回复容器
                        const aiMsgDiv2 = document.createElement("div");
                        aiMsgDiv2.className = "message ai-message";
                        const contentDiv2 = document.createElement("div");
                        contentDiv2.className = "ai-content";
                        contentDiv2.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i>';
                        aiMsgDiv2.appendChild(contentDiv2);
                        appendMessage(aiMsgDiv2);

                        // 重新调用 AI (第二轮)
                        let finalFullContent2 = "";
                        await connector.chatStream(
                            reActHistory, // 使用包含观察结果的历史
                            systemPrompt,
                            (chunk, fullContent) => {
                                finalFullContent2 = fullContent;

                                // 优化体验：从显示内容中移除 JSON 代码块，避免闪烁
                                const jsonRegex = new RegExp("```json[\\s\\S]*?(?:```|$)", "g");
                                let displayContent = fullContent.replace(jsonRegex, '').trim();

                                const rawJsonRegex = new RegExp("\\{[\\s\\S]*?\"actions\"\\s*:\\s*\\[[\\s\\S]*?\\]\\s*\\}", "g");
                                displayContent = displayContent.replace(rawJsonRegex, '').trim();

                                if (typeof marked !== 'undefined') {
                                    contentDiv2.innerHTML = marked.parse(displayContent);
                                } else {
                                    contentDiv2.textContent = displayContent;
                                }
                                scrollToBottom();
                            },
                            () => { },
                            () => { }
                        );

                        // 2.1 将第二轮的 AI 回复加入历史
                        globalChatHistory.push({ role: "user", content: `System Observation:\n${observationText}` }); // 记录观察
                        globalChatHistory.push({ role: "assistant", content: finalFullContent2 }); // 记录最终行动

                        // 检查第二轮是否有 JSON 要执行
                        // 检查第二轮是否有 JSON 要执行 (增强提取)
                        const jsonStr2 = extractJsonFromContent(finalFullContent2);
                        const jsonMatch2 = jsonStr2 ? [null, jsonStr2] : null;

                        if (jsonMatch2) {
                            const plan2 = JSON.parse(jsonMatch2[1]);
                            if (plan2.actions && Array.isArray(plan2.actions)) {
                                const statusId2 = addSystemMessage('<i class="fa-solid fa-gear fa-spin"></i> 执行最终操作...');
                                const result2 = await ExcelAgent.execute(plan2);
                                removeMessage(statusId2);

                                // 显示写入的单元格位置
                                if (result2.writtenCells && result2.writtenCells.length > 0) {
                                    const cellsStr = result2.writtenCells.join(', ');
                                    addSystemMessage(`<i class="fa-solid fa-check-circle" style="color:green"></i> 已写入: <b>${cellsStr}</b>`);
                                } else {
                                    addSystemMessage('<i class="fa-solid fa-check-circle" style="color:green"></i> 操作执行成功');
                                }
                                if (plan2.message) {
                                    contentDiv2.innerHTML = marked.parse(plan2.message);
                                }
                            }
                        } else {
                            // 第二轮AI没有输出JSON，提示用户
                            addSystemMessage('<i class="fa-solid fa-info-circle" style="color:orange"></i> AI 未生成执行操作，请尝试更明确的指令');
                        }
                    } else {
                        // 没有观察结果，说明是纯写入操作，直接完成
                        // 显示写入的单元格位置
                        if (result.writtenCells && result.writtenCells.length > 0) {
                            const cellsStr = result.writtenCells.join(', ');
                            addSystemMessage(`<i class="fa-solid fa-check-circle" style="color:green"></i> 已写入: <b>${cellsStr}</b>`);
                        } else {
                            addSystemMessage('<i class="fa-solid fa-check-circle" style="color:green"></i> 操作执行成功');
                        }

                        // 2. 如果没有 ReAct，直接记录 AI 回复
                        globalChatHistory.push({ role: "assistant", content: finalFullContent });

                        // 优化体验：如果执行成功，仅显示 message 部分，隐藏 JSON 代码块
                        // (前提是 message 字段存在)
                        if (plan.message) {
                            contentDiv.innerHTML = marked.parse(plan.message);
                        }
                    }
                }
            } catch (e) {
                console.error("JSON Execute Error:", e);
                addSystemMessage(`<i class="fa-solid fa-triangle-exclamation"></i> 执行失败: ${e.message}`);
            }
        }


    } catch (error) {
        contentDiv.innerHTML = `<span style="color: #d13438;">❌ ${error.message}</span>`;
    }
}

// ===== UI 辅助 =====
function addUserMessage(text) {
    const div = document.createElement("div");
    div.className = "message user-message";
    div.textContent = text;
    appendMessage(div);
}

function addSystemMessage(html) {
    const div = document.createElement("div");
    const id = "msg-" + Date.now();
    div.id = id;
    div.className = "message system-message";
    div.innerHTML = html;
    appendMessage(div);
    return id;
}

function removeMessage(id) {
    const el = document.getElementById(id);
    if (el) el.remove();
}

function appendMessage(el) {
    const container = document.getElementById("chat-history");
    container.appendChild(el);
    scrollToBottom();
}

function scrollToBottom() {
    const container = document.getElementById("chat-history");
    container.scrollTop = container.scrollHeight;
}

// ===== 工具函数：提取 JSON =====
function extractJsonFromContent(content) {
    if (!content) return null;

    // 1. 优先尝试 Markdown 代码块 (允许未闭合)
    const mdMatch = content.match(/```json\s*([\s\S]*?)(?:```|$)/i);
    if (mdMatch && mdMatch[1]) {
        // 简单验证有效性
        if (mdMatch[1].includes('{') && mdMatch[1].includes('}')) {
            return mdMatch[1];
        }
    }

    // 2. 尝试裸 JSON (从第一个 { 到最后一个 })
    const start = content.indexOf('{');
    const end = content.lastIndexOf('}');

    if (start >= 0 && end > start) {
        const potentialJson = content.substring(start, end + 1);
        if (potentialJson.includes('"actions"')) {
            return potentialJson;
        }
    }

    return null;
}
