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
    document.getElementById("settings-btn").onclick = showSettings;
    document.getElementById("back-btn").onclick = hideSettings;
    document.getElementById("add-custom-btn").onclick = showCustomModal;
    document.getElementById("cancel-custom-btn").onclick = hideCustomModal;
    document.getElementById("save-custom-btn").onclick = saveCustomProvider;
    document.getElementById("active-model-select").onchange = onModelChange;

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
        let list = row.querySelector(".model-filter-list");
        if (!list) {
            list = document.createElement("div");
            list.className = "model-filter-list";
            row.appendChild(list);
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

        let systemPrompt = `你是一个精通 ${hostType} 的智能助手。请简洁地回答用户问题。`;

        if (hostType === "Excel") {
            systemPrompt += `
\n**IMPORTANT EXCEL AGENT INSTRUCTIONS**
You are not just a chatbot; you are an Excel Automation Agent. 
If the user's request requires performing actions in Excel (like creating tables, formatting, charts, or setting values), you MUST return a strict JSON response following constraints below.

### 1. Formula-First Strategy (CRITICAL)
- **NEVER hardcode calculated values**. Always specificy Excel formulas.
- Example: Instead of calculating "Sum = 100" in your head, output "=SUM(A1:A10)".
- Separate assumptions/inputs from calculations.

### 2. Styling Rules (Anthropic Financial Standard)
- **Inputs**: Blue text (RGB 0,0,255).
- **Formulas**: Black text (RGB 0,0,0).
- **External Links**: Red text (RGB 255,0,0).

### 3. JSON Protocol
To execute actions, output a single JSON block wrapped in \`\`\`json ... \`\`\`.

Schema:
\`\`\`json
{
  "thought": "Brief explanation of what you are doing",
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
  "message": "Final response to show to user"
}
\`\`\`

Supported Action Types:
- setCell, setRange
- createTable
- createChart
- formatRange (style: "input"|"calculation"|"header"|"external")
- autoFit
- scanForErrors (no params) - Check for #DIV/0!, #Ref!, etc.
- fixError (params: address, value)

If the user request is just a question, answer normally without JSON.
`;
        }

        let finalFullContent = "";

        await connector.chatStream(
            message,
            systemPrompt,
            // onChunk - 收到正常内容
            (chunk, fullContent) => {
                finalFullContent = fullContent;

                // 1. 尝试检测并提取 JSON 块
                // 简单的正则匹配：```json\n{...}\n```
                const jsonMatch = fullContent.match(/```json\s*([\s\S]*?)\s*```/);

                if (jsonMatch) {
                    // 如果发现了完整 JSON 块，尝试解析并执行
                    try {
                        const jsonStr = jsonMatch[1];
                        // 简单的防抖：只在 JSON 闭合且之前未执行过时执行? 
                        // 或者：流式过程中很难判断何时是"完整"的，除了等待流结束。
                        // 策略调整：流式过程中只显示 Markdown。流结束后再检查 JSON 执行。
                        // 为了用户体验，我们实时显示 JSON 文本，流结束后如果能解析，则执行并替换显示。
                    } catch (e) { }
                }

                if (typeof marked !== 'undefined') {
                    contentDiv.innerHTML = marked.parse(fullContent);
                } else {
                    contentDiv.textContent = fullContent;
                }
                scrollToBottom();
            },
            // onThinkStart - 开始思考
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
        const jsonMatch = finalFullContent.match(/```json\s*([\s\S]*?)\s*```/);
        if (jsonMatch) {
            try {
                const plan = JSON.parse(jsonMatch[1]);
                if (plan.actions && Array.isArray(plan.actions)) {
                    // 显示正在执行状态
                    const statusId = addSystemMessage('<i class="fa-solid fa-gear fa-spin"></i> 正在执行 Excel 操作...');

                    await ExcelAgent.execute(plan);

                    removeMessage(statusId);
                    addSystemMessage('<i class="fa-solid fa-check-circle" style="color:green"></i> 操作执行成功');

                    // 优化体验：如果执行成功，仅显示 message 部分，隐藏 JSON 代码块
                    // (前提是 message 字段存在)
                    if (plan.message) {
                        contentDiv.innerHTML = marked.parse(plan.message);
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
