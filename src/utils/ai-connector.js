/*
 * AI 连接器 - 纯净稳定版 v2026.01.26
 * 设计理念：KISS (Keep It Simple, Stupid)
 * 1. 移除所有 System Prompt 注入：防止通用模型（智谱/Gemini）因格式强迫而崩溃或死循环。
 * 2. 移除标签解析：不再手动解析 <think>，避免缓冲区截断导致的丢字。
 * 3. 原生支持 R1：仅对 DeepSeek R1 的 reasoning_content 字段做折叠处理。
 */

class AiConnector {
    constructor(apiKey, model, baseUrl) {
        this.apiKey = apiKey;
        this.model = model;
        this.baseUrl = baseUrl.replace(/\/+$/, '');
    }

    // 流式聊天
    async chatStream(userPrompt, systemPrompt, onChunk, onThinkStart, onThinkEnd) {
        if (!this.apiKey) throw new Error("请先配置 API Key");

        const chatUrl = this.baseUrl + "/chat/completions";
        console.log('[AI] Stream request to:', chatUrl);

        const headers = {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${this.apiKey}`
        };

        // 纯净模式：不注入任何指令，原样传递 System Prompt
        // 这样 Gemini 和 智谱 会恢复正常的直接回答（解决"只有思考没回答"的问题）
        const body = {
            model: this.model,
            messages: [
                { role: "system", content: systemPrompt },
                { role: "user", content: userPrompt }
            ],
            stream: true,
            temperature: 0.7
        };

        const response = await fetch(chatUrl, {
            method: "POST",
            headers: headers,
            body: JSON.stringify(body)
        });

        if (!response.ok) {
            const errorText = await response.text();
            let errMsg = `API 错误 (${response.status})`;
            try {
                const errData = JSON.parse(errorText);
                errMsg = errData.error?.message || errData.message || errMsg;
            } catch (e) { }
            throw new Error(errMsg);
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();

        // 核心变量
        let streamBuffer = "";
        let fullContent = ""; // 累积完整正文
        let thinkContent = ""; // 累积思考内容

        // 状态机
        let isThinking = false; // 仅用于 R1 原生思考状态
        let thinkStartTime = null;

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            streamBuffer += decoder.decode(value, { stream: true });
            const lines = streamBuffer.split('\n');
            streamBuffer = lines.pop() || "";

            for (const line of lines) {
                const trimmed = line.trim();
                if (!trimmed || !trimmed.startsWith('data: ')) continue;

                const data = trimmed.slice(6);
                if (data === '[DONE]') continue;

                try {
                    const json = JSON.parse(data);
                    const delta = json.choices?.[0]?.delta;

                    if (delta) {
                        const content = delta.content || "";
                        const reasoning = delta.reasoning_content || "";

                        // 1. 处理 DeepSeek R1 原生思考
                        if (reasoning) {
                            if (!isThinking) {
                                isThinking = true;
                                thinkStartTime = Date.now();
                                if (onThinkStart) onThinkStart();
                            }
                            thinkContent += reasoning;
                            if (onThinkEnd) onThinkEnd(thinkContent, Math.round((Date.now() - thinkStartTime) / 1000));
                        }

                        // 2. 处理正文 (无条件输出，确保不丢字)
                        if (content) {
                            fullContent += content;
                            if (onChunk) onChunk(content, fullContent);
                        }
                    }
                } catch (e) {
                    // console.warn('[AI] Parse error:', e);
                }
            }
        }

        return "";
    }

    static async fetchModels(apiKey, baseUrl) {
        if (!baseUrl || !apiKey) return [];
        const url = baseUrl.replace(/\/+$/, '') + '/models';
        try {
            const response = await fetch(url, {
                method: "GET",
                headers: {
                    "Authorization": `Bearer ${apiKey}`,
                    "Content-Type": "application/json"
                }
            });
            if (!response.ok) return [];
            const data = await response.json();
            if (data.data && Array.isArray(data.data)) {
                return data.data.map(m => m.id);
            }
            return [];
        } catch (error) {
            console.error('[AI] Fetch models error:', error);
            return [];
        }
    }
}

window.AiConnector = AiConnector;
