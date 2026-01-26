/*
 * Excel Agent Executor
 * 负责执行来自 AI 的 JSON 指令
 */

class ExcelAgentExecutor {
    constructor() {
        this.queue = [];
        this.isExecuting = false;
    }

    /**
     * 执行 AI 返回的指令集
     * @param {Object} plan - { thought, actions: [], message }
     */
    async execute(plan) {
        if (!plan || !plan.actions || !Array.isArray(plan.actions)) {
            console.error("Invalid plan:", plan);
            return { observations: [] };
        }

        // 收集读取操作的观察结果
        const observations = [];
        // 收集写入的单元格地址
        const writtenCells = [];

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();

                // --- 2.0 智能防撞逻辑: 基于碰撞检测 (Collision Detection) ---
                // 不再盲目检查 A1，而是检查"我想写的地方"是否真的有人占了

                const usedRange = sheet.getUsedRange(true); // getUsedRangeOrNullObject if possible? No, getUsedRange is standard.
                // Standard getUsedRange returns the range. If empty, it might be A1.

                // 批量加载需要的信息
                usedRange.load("rowIndex, columnIndex, rowCount, columnCount, values, address");
                // 同时加载所有 action涉及的 range 的坐标信息，用于计算碰撞
                const actionRanges = [];

                // 过滤出写入操作 (排除读取操作)
                const writeActions = plan.actions.filter(a => !['readRange', 'findData', 'getUsedRangeInfo'].includes(a.type));

                writeActions.forEach(action => {
                    // 提取 action 中的地址参数
                    const addr = action.params?.range || action.params?.address || action.params?.sourceData;
                    if (addr && typeof addr === 'string') {
                        // 创建 range 对象并加载坐标
                        const r = sheet.getRange(addr);
                        r.load("rowIndex, columnIndex, rowCount, columnCount, address");
                        actionRanges.push({ action, rangeObj: r });
                    }
                });

                await context.sync();

                // 检查是否为空表 (如果 UsedRange 是 A1 且值为空)
                // 注意: getUsedRange 对空表通常返回 A1
                const isSheetEmpty = (usedRange.rowCount === 1 && usedRange.columnCount === 1 && usedRange.values[0][0] === "");

                let rowOffset = 0;
                let collisionDetected = false;

                if (!isSheetEmpty && writeActions.length > 0) {
                    // 2.0.1 修正: 禁用自动碰撞偏移
                    // 用户明确指定了单元格 (如 A153)，我们不应该自作主张把它移到下面去
                    // 只有当用户没有指定具体行时，AI 应该自己决定位置，而不是靠底层硬搬
                    console.log("Smart Collision Detection disabled: Prioritizing user specified coordinates.");

                    /* 
                    // 原有逻辑: 只要通过 Bounding Box 检测到任何重叠，就整体平移
                    for (const item of actionRanges) {
                        // ... (Code removed/disabled)
                    }
                    if (collisionDetected) {
                        rowOffset = usedRange.rowIndex + usedRange.rowCount + 2;
                    }
                    */
                }

                // 执行所有动作
                for (const action of plan.actions) {
                    const result = await this.performAction(context, sheet, action, rowOffset);
                    // 如果是读取操作，收集观察结果
                    if (result && result.observation) {
                        observations.push(result.observation);
                    }
                    // 如果是写入操作，收集写入的地址
                    if (result && result.written) {
                        writtenCells.push(result.written);
                    }
                }

                await context.sync();
            });
        } catch (error) {
            console.error("Execution error:", error);
            throw error;
        }

        // 返回观察结果和写入的单元格
        return { observations, writtenCells };
    }

    async performAction(context, sheet, action, rowOffset = 0) {
        // 动态调整地址
        let params = { ...action.params };
        if (rowOffset > 0) {
            // 遍历 params 中所有可能是地址的字段
            ['address', 'range', 'sourceData'].forEach(key => {
                if (params[key] && typeof params[key] === 'string') {
                    params[key] = this.offsetAddress(params[key], rowOffset);
                }
            });
        }

        console.log(`Executing: ${action.type}`, params);

        // 使用调整后的 params
        switch (action.type) {
            case "setCell":
                // params: { address: "A1", value: "Text" or 123, formula: boolean }
                if (!params.address) throw new Error("Action 'setCell' missing required param: 'address'");
                const range = sheet.getRange(params.address);
                if (params.formula) {
                    range.formulas = [[params.value]];
                } else {
                    range.values = [[params.value]];
                }
                return { written: params.address, type: "setCell" };

            case "setRange":
                // 兼容 range 和 address 参数
                const rangeAddr = params.range || params.address;
                if (!rangeAddr) throw new Error("Action 'setRange' missing required param: 'range' or 'address'");
                const multiRange = sheet.getRange(rangeAddr);

                if (params.values) {
                    // 情况1: 传入了完整的二维数组 values
                    multiRange.values = params.values;
                } else if (params.value !== undefined) {
                    // 情况2: 传入了单个 value (如填充公式或固定值)
                    // 如果是公式
                    if (params.formula) {
                        multiRange.formulas = params.value; // Excel JS API 允许给区域赋单个公式，会自动填充
                    } else {
                        multiRange.values = params.value; // 同上，给区域赋单个值
                    }
                }
                return { written: rangeAddr, type: "setRange" };

            case "createTable":
                // params: { range: "A1:C5", name: "MyTable", header: boolean }
                if (!params.range) throw new Error("Action 'createTable' missing required param: 'range'");
                // 表格名称不能重复，如果偏移了，建议改名？
                // 暂不改名，如果重名 Excel 会自动处理 (e.g. MyTable2) 或者是抛错? 
                // Excel API tables.add 如果名字重复会报错。
                // 我们给名字加个随机后缀防撞? 
                // 或者是捕获错误。
                // 简单起见，如果 offset > 0，给 name 加后缀
                if (rowOffset > 0 && params.name) {
                    params.name = `${params.name}_${Date.now().toString().slice(-4)}`;
                }

                const table = sheet.tables.add(params.range, params.header !== false);
                if (params.name) {
                    try {
                        table.name = params.name;
                    } catch (e) {
                        console.warn("Table name conflict, using default.");
                    }
                }
                break;

            case "formatRange":
                // params: { range: "A1", style: "input" | "calculation" | "header" | custom... }
                if (!params.range) throw new Error("Action 'formatRange' missing required param: 'range'");
                await this.applyStyle(sheet.getRange(params.range), params);
                break;

            case "createChart":
                // params: { type: "ColumnClustered", sourceData: "A1:B10" }
                if (!params.sourceData) throw new Error("Action 'createChart' missing required param: 'sourceData'");
                const dataRange = sheet.getRange(params.sourceData);
                const chart = sheet.charts.add(params.type, dataRange, "Auto");

                // 如果有偏移，尝试把图表也往下挪一点? (Auto 模式通常会放在数据旁边)
                if (rowOffset > 0) {
                    chart.top = rowOffset * 20; // 估算高度
                }
                break;

            case "autoFit":
                // params: { range: "A:C" }
                if (!params.range) throw new Error("Action 'autoFit' missing required param: 'range'");
                sheet.getRange(params.range).getEntireColumn().format.autofitColumns();
                break;

            case "scanForErrors":
                // scanForErrors 扫描的是 UsedRange，不需要偏移参数，它会自动扫描新区域
                await this.scanForErrors(context, sheet);
                break;

            case "fixError":
                const fixRange = sheet.getRange(params.address);
                fixRange.values = [[params.value]];
                fixRange.format.fill.clear();
                break;

            // ========= READ ACTIONS (ReAct Loop) =========
            case "readRange":
                // params: { address: "A1:C10" }
                if (!params.address) throw new Error("Action 'readRange' missing required param: 'address'");
                // 返回读取到的内容，供 AI 进一步分析
                const readTarget = sheet.getRange(params.address);
                readTarget.load("values");
                await context.sync();
                // 将二维数组转换为 CSV 便于 AI 理解
                const csvResult = readTarget.values.map(row => row.join(",")).join("\n");
                // 返回观察结果 (observation)
                return { observation: `Range ${params.address} contents:\n${csvResult}` };

            case "findData":
                // params: { keyword: "保费" }
                // 在 UsedRange 中搜索关键词
                const findUsedRange = sheet.getUsedRange();
                findUsedRange.load("values, address");
                await context.sync();

                const keyword = params.keyword;
                const foundCells = [];
                const vals = findUsedRange.values;
                for (let i = 0; i < vals.length; i++) {
                    for (let j = 0; j < vals[i].length; j++) {
                        if (String(vals[i][j]).includes(keyword)) {
                            foundCells.push(this.getA1Address(i, j));
                        }
                    }
                }
                if (foundCells.length > 0) {
                    return { observation: `Found "${keyword}" in cells: ${foundCells.join(", ")}` };
                } else {
                    return { observation: `"${keyword}" not found in the active sheet.` };
                }

            case "getUsedRangeInfo":
                // 获取有效数据范围摘要 (只读取地址，不读取值，节省资源)
                const usedRangeInfo = sheet.getUsedRange();
                usedRangeInfo.load("address, rowCount, columnCount");
                await context.sync();
                // 只返回地址和大小信息
                const summary = `Used range: ${usedRangeInfo.address}, ${usedRangeInfo.rowCount} rows x ${usedRangeInfo.columnCount} cols.`;
                return { observation: summary };

            default:
                console.warn(`Unknown action type: ${action.type}`);
        }
    }

    // 地址偏移辅助函数: "A1:B2" -> offset 10 -> "A11:B12"
    offsetAddress(address, rowOffset) {
        if (!address) return address;
        if (rowOffset === 0) return address;

        // 替换所有数字： A1 -> A(1+offset)
        return address.replace(/(\d+)/g, (match) => {
            const row = parseInt(match);
            return (row + rowOffset).toString();
        });
    }

    async applyStyle(range, params) {
        // 实现 Anthropic Financial Styling
        if (params.style === 'input') {
            range.format.font.color = "blue"; // User Inputs
        } else if (params.style === 'calculation') {
            range.format.font.color = "black"; // Formulas
        } else if (params.style === 'external') {
            range.format.font.color = "red"; // External Links
        } else if (params.style === 'link') {
            range.format.font.color = "green"; // Internal Links
        } else if (params.style === 'header') {
            range.format.font.bold = true;
            range.format.fill.color = "#E0E0E0";
            range.format.horizontalAlignment = "Center";
        }

        // 自定义样式
        if (params.bold) range.format.font.bold = true;
        if (params.color) range.format.font.color = params.color;
        if (params.fill) range.format.fill.color = params.fill;
    }

    /**
     * 扫描表格中的错误值 (#DIV/0!, #REF!, etc)
     */
    async scanForErrors(context, sheet) {
        const usedRange = sheet.getUsedRange();
        usedRange.load("valueTypes");
        await context.sync();

        const types = usedRange.valueTypes;
        let errorCount = 0;

        for (let i = 0; i < types.length; i++) {
            for (let j = 0; j < types[i].length; j++) {
                if (types[i][j] === Excel.RangeValueType.error) {
                    errorCount++;
                    const cell = usedRange.getCell(i, j);
                    cell.format.fill.color = "#FFCCCC"; // 标红
                }
            }
        }

        if (errorCount > 0) {
            console.warn(`Found ${errorCount} errors.`);
        }
    }

    getA1Address(row, col) {
        let colStr = "";
        let c = col;
        while (c >= 0) {
            colStr = String.fromCharCode(65 + (c % 26)) + colStr;
            c = Math.floor(c / 26) - 1;
        }
        return `${colStr}${row + 1}`;
    }
    /**
     * 获取当前表格的上下文数据 (表头 + 前几行)
     * 用于让 AI "看见" 表格结构
     */
    async getContextData() {
        try {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();

                // 1. 获取用户选中区域 (Selection)
                const selection = context.workbook.getSelectedRange();
                selection.load("address, rowCount, columnCount, values");
                // 2. 获取已用区域 (UsedRange)
                const usedRange = sheet.getUsedRange();
                usedRange.load("rowCount, columnCount, rowIndex, columnIndex");

                await context.sync();

                // 构建上下文消息
                let contextMsg = "";

                // 处理选中区域
                if (selection) {
                    // 读取选中区域的值 (限制大小以免爆 Token)
                    const maxCells = 50;
                    const cellCount = selection.rowCount * selection.columnCount;

                    if (cellCount <= maxCells) {
                        const selectionData = selection.values.map(row => row.join(", ")).join("\n");
                        contextMsg += `==当前选中区域: ${selection.address}==\n${selectionData}\n\n`;
                    } else {
                        contextMsg += `==当前选中区域: ${selection.address}== (区域过大，仅显示地址)\n\n`;
                    }
                }

                // 处理整体数据概览
                if (usedRange.rowCount === 0) {
                    contextMsg += "当前表格为空。";
                    return contextMsg;
                }

                // 限制读取行数 (例如只读前 5 行用于理解结构)
                const previewRowCount = Math.min(usedRange.rowCount, 5);
                // 获取预览区域 (从 UsedRange 的起始位置开始)
                const previewRange = sheet.getRangeByIndexes(
                    usedRange.rowIndex,
                    usedRange.columnIndex,
                    previewRowCount,
                    usedRange.columnCount
                );

                previewRange.load("values");
                await context.sync();

                // 将二维数组转换为简单的 CSV 格式字符串
                const csvData = previewRange.values.map(row => row.join(",")).join("\n");
                contextMsg += `==表格数据概览 (前 ${previewRowCount} 行)==\n${csvData}`;

                return contextMsg;
            });
        } catch (error) {
            console.error("Error getting context:", error);
            return "无法读取表格上下文。";
        }
    }
}


// 导出单例
window.ExcelAgent = new ExcelAgentExecutor();
