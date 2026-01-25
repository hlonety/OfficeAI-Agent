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
            return;
        }

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();

                // --- 2.0 智能防撞逻辑: 基于碰撞检测 (Collision Detection) ---
                // 不再盲目检查 A1，而是检查"我想写的地方"是否真的有人占了

                const usedRange = sheet.getUsedRange(true); // getUsedRangeOrNullObject if possible? No, getUsedRange is standard.
                // Standard getUsedRange returns the range. If empty, it might be A1.

                // 批量加载需要的信息
                usedRange.load("rowIndex, columnIndex, rowCount, columnCount, values");
                // 同时加载所有 action涉及的 range 的坐标信息，用于计算碰撞
                const actionRanges = [];

                plan.actions.forEach(action => {
                    // 提取 action 中的地址参数
                    const addr = action.params?.range || action.params?.address || action.params?.sourceData;
                    if (addr && typeof addr === 'string') {
                        // 创建 range 对象并加载坐标
                        const r = sheet.getRange(addr);
                        r.load("rowIndex, columnIndex, rowCount, columnCount");
                        actionRanges.push({ action, rangeObj: r });
                    }
                });

                await context.sync();

                // 检查是否为空表 (如果 UsedRange 是 A1 且值为空)
                // 注意: getUsedRange 对空表通常返回 A1
                const isSheetEmpty = (usedRange.rowCount === 1 && usedRange.columnCount === 1 && usedRange.values[0][0] === "");

                let rowOffset = 0;
                let collisionDetected = false;

                if (!isSheetEmpty) {
                    // 只有当表不为空时，才需要检测碰撞
                    for (const item of actionRanges) {
                        const target = item.rangeObj;
                        const used = usedRange;

                        // 数学判断矩形相交
                        // 矩形1: [r1, c1, r1+h1, c1+w1]
                        // 矩形2: [r2, c2, r2+h2, c2+w2]
                        // 不相交条件: t_bottom <= u_top OR t_top >= u_bottom OR t_right <= u_left OR t_left >= u_right

                        const t_top = target.rowIndex;
                        const t_bottom = target.rowIndex + target.rowCount;
                        const t_left = target.columnIndex;
                        const t_right = target.columnIndex + target.columnCount;

                        const u_top = used.rowIndex;
                        const u_bottom = used.rowIndex + used.rowCount;
                        const u_left = used.columnIndex;
                        const u_right = used.columnIndex + used.columnCount;

                        const isDisjoint = (t_bottom <= u_top) || (t_top >= u_bottom) || (t_right <= u_left) || (t_left >= u_right);

                        if (!isDisjoint) {
                            console.warn("Collision Detected!", { target: target.address, used: usedRange.address });
                            collisionDetected = true;
                            break; // 只要有一个撞了，就整体迁移，保持相对结构
                        }
                    }

                    if (collisionDetected) {
                        // 偏移量 = UsedRange 底部 + 2行缓冲
                        rowOffset = usedRange.rowIndex + usedRange.rowCount + 2;
                        console.log(`Applying Smart Offset: ${rowOffset} rows.`);
                    }
                }

                // 执行动作
                for (const action of plan.actions) {
                    await this.performAction(context, sheet, action, rowOffset);
                }

                await context.sync();
            });
        } catch (error) {
            console.error("Execution error:", error);
            throw error;
        }
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
                const range = sheet.getRange(params.address);
                if (params.formula) {
                    range.formulas = [[params.value]];
                } else {
                    range.values = [[params.value]];
                }
                break;

            case "setRange":
                // 兼容 range 和 address 参数
                const rangeAddr = params.range || params.address;
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
                break;

            case "createTable":
                // params: { range: "A1:C5", name: "MyTable", header: boolean }
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
                await this.applyStyle(sheet.getRange(params.range), params);
                break;

            case "createChart":
                // params: { type: "ColumnClustered", sourceData: "A1:B10" }
                // chart 放置位置也需要偏移吗? sheet.charts.add 默认放在可见区域。
                // 我们可以设置 chart.top / left，但这比较复杂。
                // 暂时只确保 sourceData 是对的。
                const dataRange = sheet.getRange(params.sourceData);
                const chart = sheet.charts.add(params.type, dataRange, "Auto");

                // 如果有偏移，尝试把图表也往下挪一点? (Auto 模式通常会放在数据旁边)
                if (rowOffset > 0) {
                    chart.top = rowOffset * 20; // 估算高度
                }
                break;

            case "autoFit":
                // params: { range: "A:C" }
                // 整列 autoFit 不需要偏移 range (A:C 还是 A:C)
                // 但如果是具体区域 A1:C5，就要偏移。
                // 正则判断：如果包含数字，则偏移。如果不含数字(整列)，不偏移。
                if (/\d/.test(action.params.range)) { // check original
                    sheet.getRange(params.range).getEntireColumn().format.autofitColumns();
                } else {
                    sheet.getRange(action.params.range).getEntireColumn().format.autofitColumns();
                }
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
}


// 导出单例
window.ExcelAgent = new ExcelAgentExecutor();
