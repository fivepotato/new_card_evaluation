"use strict";
const fs = require("fs");
const xlsx = require("xlsx");

//特定评分项目
const special_ratings = new Map([
    [101023001, [{ type: "后排输出", rate: 7 }]],
    [401033001, [{ type: "后排输出", rate: 4 }]],
    [302023001, [{ type: "后排输出", rate: 7 }]],

    [101053001, [{ type: "后排耐久", rate: 6 }]],

    [300083001, [{ type: "特技后排", rate: 6 }]],
    [201053001, [{ type: "特技后排", rate: 6 }]],
    [202083001, [{ type: "SP后排", rate: 8.5 }]],
    [202073002, [{ type: "SP后排", rate: 8.5 }]],
]);

let filename = null;
{
    const a = fs.readdirSync("./");
    for (const fn of a.values()) {
        if (fn.split(".").pop() === "xlsm") {
            filename = fn;
            break;
        }
    }
    if (!filename) throw "excel not found in current directory";
}
const workbook = xlsx.readFile(filename);

const card_attrs_1 = (() => {
    const sheet = workbook.Sheets["卡片强度"];
    const structure = {};
    let classify_line2 = null;
    for (let col = 2; col <= 300; col++) {
        const line2 = sheet[`${table_column(col)}2`];
        const line3 = sheet[`${table_column(col)}3`];
        if (!line3)
            continue;
        if (line2)
            structure[line2.v] = {};
        classify_line2 = line2 || classify_line2;
        structure[classify_line2.v][line3.v] = table_column(col);
    }
    const card_lines = [];
    for (let line = 4; line < 10000; line++) {
        const card_master_id = sheet[`${structure["卡片基本属性"]["数据库ID"]}${line}`];
        if (card_master_id) {
            card_lines.push(line);
        }
    }

    return card_lines.map((line) => ({
        no: sheet[structure["卡片基本属性"]["ID"] + line].v,
        id: sheet[structure["卡片基本属性"]["数据库ID"] + line].v,
        rarity: sheet[structure["卡片基本属性"]["稀有度"] + line].v,
        event_no: sheet[structure["卡片基本属性"]["配信活动"] + line].v,
        前排输出: sheet[structure["前排强度"]["输出"] + line].v,
        前排真输出: sheet[structure["前排强度"]["真输出"] + line].v,
        后排输出: sheet[structure["后排强度"]["输出"] + line].v,
        好友支援: sheet[structure["好友支援"]["输出"] + line].v,
        前排耐久: sheet[structure["前排强度"]["耐久/键"] + line].v,
        后排耐久: sheet[structure["后排强度"]["耐久/键"] + line].v,
        后排输出同属性: (() => {
            const backline = [];
            const v0 = sheet[structure["槽强度"]["输出"] + line].v;
            const v1 = sheet[structure["被动个性1强度"]["后排输出"] + line] && sheet[structure["被动个性1强度"]["后排输出"] + line].v || 0;
            const v2 = sheet[structure["被动个性2强度"]["后排输出"] + line] && sheet[structure["被动个性2强度"]["后排输出"] + line].v || 0;
            const v3 = sheet[structure["主动个性1强度"]["后排输出"] + line] && sheet[structure["主动个性1强度"]["后排输出"] + line].v || 0;
            const v4 = sheet[structure["主动个性2强度"]["后排输出"] + line] && sheet[structure["主动个性2强度"]["后排输出"] + line].v || 0;
            backline.push(v0);
            if (v1 && sheet[structure["被动个性1强度"]["范围"] + line].v === "同属性")
                backline.push(v1 * 3);
            else
                backline.push(v1);
            if (v2 && sheet[structure["被动个性2强度"]["范围"] + line].v === "同属性")
                backline.push(v2 * 3);
            else
                backline.push(v2);
            if (v3 && sheet[structure["主动个性1强度"]["对象"] + line] && sheet[structure["主动个性1强度"]["对象"] + line].v === "同属性")
                backline.push(v3 * 3);
            else
                backline.push(v3);
            if (v4 && sheet[structure["主动个性2强度"]["对象"] + line] && sheet[structure["主动个性2强度"]["对象"] + line].v === "同属性")
                backline.push(v4 * 3);
            else
                backline.push(v4);
            return backline.reduce((prev, curr) => prev + curr, 0);
        })(),
        特殊评分: (() => {
            const array = [];
            if (sheet[`${structure["卡片基本属性"]["type"]}${line}`].v === "sk") {
                const skill_1 = sheet[`${structure["技能1强度"]["类型"]}${line}`].v, skill_2 = sheet[`${structure["技能2强度"]["类型"]}${line}`].v;
                if (skill_1 === "sp获得" || skill_2 === "sp获得")
                    array.push({ type: "sk充电", rate: 4 });
                else if (skill_1 === "回血" || skill_2 === "回血")
                    array.push({ type: "sk奶", rate: 4 });
            }
            const id = sheet[structure["卡片基本属性"]["数据库ID"] + line].v;
            const sp = special_ratings.get(id);
            if (sp) sp.forEach((s) => array.push(s));
            return array;
        })(),

    })).filter(({ rarity }) => rarity === "UR");
});

const card_attrs_2 = (() => {
    const sheet = workbook.Sheets["奶盾强度计算"];
    const structure = {};
    let classify_line2 = null;
    for (let col = 2; col <= 300; col++) {
        const line2 = sheet[`${table_column(col)}2`];
        const line3 = sheet[`${table_column(col)}3`];
        if (!line3) continue;
        if (line2) structure[line2.v] = {};
        classify_line2 = line2 || classify_line2;
        structure[classify_line2.v][line3.v] = table_column(col);
    }
    const card_attrs_21 = [];
    for (let line = 4; line < 10000; line++) {
        const card_master_id = sheet[`${structure["站撸奶强度"]["ID"]}${line}`];
        if (card_master_id) {
            card_attrs_21.push({
                id: sheet[`${structure["站撸奶强度"]["ID"]}${line}`].v,
                单前排平均: sheet[`${structure["站撸奶强度"]["平均"]}${line}`].v,
            });
        }
    }
    const card_attrs_22 = [];
    for (let line = 4; line < 10000; line++) {
        const card_master_id = sheet[`${structure["切队奶强度"]["ID"]}${line}`];
        if (card_master_id) {
            card_attrs_22.push({
                id: sheet[`${structure["切队奶强度"]["ID"]}${line}`].v,
                双前排平均: sheet[`${structure["切队奶强度"]["平均"]}${line}`].v,
            });
        }
    }
    return { card_attrs_21, card_attrs_22 };
});

module.exports = {
    card_attrs_1, card_attrs_2,
}

function table_column(number) {
    if (number < 1)
        throw `number ${number} < 1`;
    if (number > 702)
        throw `number ${number} > 702`;
    let code = [];
    code[1] = (number - 1) % 26 + 65;
    number = Math.floor((number - 1) / 26);
    if (number)
        code[0] = number + 64;
    return code.map(c => String.fromCharCode(c)).join("");
}