"use strict";
const puppeteer = require("puppeteer");
const excel_api = require("./excel_api");
const fs = require("fs");
const CHROME_PATH = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe";

async function get_data() {
    const card_attrs = excel_api.card_attrs_1()
    const { card_attrs_21, card_attrs_22 } = excel_api.card_attrs_2();

    console.log("data got from excel");
    card_attrs.sort((a, b) => a.no - b.no);
    const card_attrs_new = (() => {
        const event_no_new = card_attrs[card_attrs.length - 1].event_no;
        return card_attrs.filter(({ id, event_no }) => event_no - event_no_new >= 0/*id === 300033002*/);
    })().map(({ no, id, 前排输出, 前排真输出, 后排输出, 好友支援, 前排耐久, 后排耐久, 后排输出同属性, 好友支援同属性, 特殊评分 }) => {
        const 前排输出排名 = card_attrs.sort((a, b) => b.前排输出 - a.前排输出).map(({ id }) => id).indexOf(id) + 1;
        const 前排真输出排名 = card_attrs.sort((a, b) => b.前排真输出 - a.前排真输出).map(({ id }) => id).indexOf(id) + 1;
        const 后排输出排名 = card_attrs.sort((a, b) => b.后排输出 - a.后排输出).map(({ id }) => id).indexOf(id) + 1;
        const 好友支援排名 = card_attrs.sort((a, b) => b.好友支援 - a.好友支援).map(({ id }) => id).indexOf(id) + 1;
        const 前排耐久排名 = card_attrs.sort((a, b) => b.前排耐久 - a.前排耐久).map(({ id }) => id).indexOf(id) + 1;
        const 后排耐久排名 = card_attrs.sort((a, b) => b.后排耐久 - a.后排耐久).map(({ id }) => id).indexOf(id) + 1;
        const 后排输出同属性排名 = card_attrs.sort((a, b) => b.后排输出同属性 - a.后排输出同属性).map(({ id }) => id).indexOf(id) + 1;
        const 好友支援同属性排名 = card_attrs.sort((a, b) => b.好友支援同属性 - a.好友支援同属性).map(({ id }) => id).indexOf(id) + 1;

        if (card_attrs_21.map(({ id }) => id).indexOf(id) !== -1) {
            const 单前排奶盾能力 = card_attrs_21[card_attrs_21.map(({ id }) => id).indexOf(id)].单前排平均;
            const 单前排奶盾能力排名 = card_attrs_21.sort((a, b) => b.单前排平均 - a.单前排平均).map(({ id }) => id).indexOf(id) + 1;
            const 双前排奶盾能力 = card_attrs_22[card_attrs_22.map(({ id }) => id).indexOf(id)].双前排平均;
            const 双前排奶盾能力排名 = card_attrs_22.sort((a, b) => b.双前排平均 - a.双前排平均).map(({ id }) => id).indexOf(id) + 1;

            const 奶盾中前排输出排名 = card_attrs
                .sort((a, b) => b.前排输出 - a.前排输出)
                .filter(({ id }) => card_attrs_21.map(({ id }) => id).indexOf(id) !== -1)
                .map(({ id }) => id)
                .indexOf(id) + 1;
            const 奶盾中后排输出排名 = card_attrs
                .sort((a, b) => b.后排输出 - a.后排输出)
                .filter(({ id }) => card_attrs_21.map(({ id }) => id).indexOf(id) !== -1)
                .map(({ id }) => id)
                .indexOf(id) + 1;
            return {
                no, id, 前排输出, 前排输出排名, 前排真输出, 前排真输出排名, 后排输出, 后排输出排名, 好友支援, 好友支援排名, 前排耐久, 前排耐久排名, 后排耐久, 后排耐久排名,
                后排输出同属性, 后排输出同属性排名, 好友支援同属性, 好友支援同属性排名,
                单前排奶盾能力, 单前排奶盾能力排名, 双前排奶盾能力, 双前排奶盾能力排名,
                奶盾中前排输出排名, 奶盾中后排输出排名,
                特殊评分,
            };
        }

        return {
            no, id, 前排输出, 前排输出排名, 前排真输出, 前排真输出排名, 后排输出, 后排输出排名, 好友支援, 好友支援排名, 前排耐久, 前排耐久排名, 后排耐久, 后排耐久排名,
            后排输出同属性, 后排输出同属性排名, 好友支援同属性, 好友支援同属性排名,
            特殊评分,
        }

    });
    return (card_attrs_new);
}

class rating_constants {
    static absv = 86013.54832;
    static absvr = 20000;
    static rating_borders = {
        前排输出: [
            { border: 1.00, rating: 9.5, description: "SS+" },
            { border: 0.95, rating: 9, description: "SS-" },
            { border: 0.90, rating: 9, description: "S+" },
            { border: 0.85, rating: 8.5, description: "S-" },
            { border: 0.80, rating: 8, description: "A+" },
            { border: 0.75, rating: 7.5, description: "A-" },
            { border: 0.70, rating: 7, description: "B+" },
            { border: 0.65, rating: 6.5, description: "B-" },
            { border: 0.60, rating: 6, description: "C+" },
            { border: 0.55, rating: 5, description: "C-" },
            { border: 0.50, rating: 4, description: "D+" },
            { border: 0.45, rating: 3, description: "D-" },
            { border: 0.40, rating: 2, description: "E+" },
            { border: 0.35, rating: 1, description: "E-" },
            { border: 0.00, rating: 0, description: "F" },
        ],
        前排真输出: [
            { border: 1.24, rating: 9.5, description: "SS+" },
            { border: 1.215, rating: 9, description: "SS-" },
            { border: 1.19, rating: 9, description: "S+" },
            { border: 1.17, rating: 8.5, description: "S-" },
            { border: 1.15, rating: 8, description: "A+" },
            { border: 1.13, rating: 7.5, description: "A-" },
            { border: 1.11, rating: 7, description: "B+" },
        ],
        前排耐久: [
            { border: 1150, description: "SS" },
            { border: 1025, description: "S" },
            { border: 900, description: "A" },
            { border: 775, description: "B" },
            { border: 650, description: "C" },
            { border: 525, description: "D" },
            { border: 400, description: "E" },
        ],
        单前排奶盾能力: [
            { border: 100000, rating: 8, description: "A+" },
            { border: 95000, rating: 7.5, description: "A-" },
            { border: 90000, rating: 7, description: "B+" },
            { border: 85000, rating: 6.5, description: "B-" },
            { border: 80000, rating: 6, description: "C+" },
            { border: 75000, rating: 5, description: "C-" },
            { border: 70000, rating: 4, description: "D+" },
            { border: 65000, rating: 3, description: "D-" },
            { border: 60000, rating: 2, description: "E+" },
            { border: 0, rating: 1, description: "E-" },
        ],
        双前排奶盾能力: [
            { border: 124000, rating: 9, description: "S+" },
            { border: 121000, rating: 8.5, description: "S-" },
            { border: 118000, rating: 8, description: "A+" },
            { border: 116500, rating: 7.5, description: "A-" },
            { border: 115000, rating: 7, description: "B+" },
            { border: 113500, rating: 6.5, description: "B-" },
            { border: 112000, rating: 6, description: "C+" },
            { border: 109000, rating: 5, description: "C-" },
            { border: 106000, rating: 4, description: "D+" },
            { border: 103000, rating: 3, description: "D-" },
            { border: 100000, rating: 2, description: "E+" },
            { border: 0, rating: 1, description: "E-" },
        ],
        后排输出: [
            { border: 0.18, rating: 9.5, description: "SS" },
            { border: 0.15, rating: 9, description: "S+" },
            { border: 0.14, rating: 8.5, description: "S-" },
            { border: 0.13, rating: 8, description: "A+" },
            { border: 0.12, rating: 7.5, description: "A-" },
            { border: 0.11, rating: 7, description: "B" },
            { border: 0.10, rating: 6, description: "C" },
            { border: 0.089, rating: 4, description: "D" },
        ],
        后排输出同属性: [
            { border: 0.18, rating: 7.5, description: "SS" },
            { border: 0.15, rating: 7, description: "S+" },
            { border: 0.14, rating: 6.5, description: "S-" },
            { border: 0.13, rating: 6, description: "A+" },
        ],
        后排耐久: [
            { border: 130, rating: 7.5, description: "S" },
            { border: 100, rating: 6, description: "A+" },
            { border: 70, rating: 5, description: "A-" },
        ],
        好友支援: [
            { border: 0.20, description: "S" },
            { border: 0.18, description: "A+" },
            { border: 0.16, description: "A-" },
        ],
        好友支援同属性: [
            { border: 0.20, description: "S" },
            { border: 0.18, description: "A+" },
            { border: 0.16, description: "A-" },
        ],
    }
    static rating_descriptions = new Map([
        [10, "神"],
        [9.5, "准神"],
        [9, "超强"],
        [8.5, "强力"],
        [8, "较强"],
        [7.5, "可用"],
        [7, "可用"],
        [6.5, "较弱"],
        [6, "较弱"],
        [5, "弱"],
        [4, "弱"],
        [3, "弱"],
        [2, "弱"],
        [1, "弱"],
        [0, "弱"],
    ])
    static rating(a) {
        const {
            no, id, 前排输出, 前排输出排名, 前排真输出, 前排真输出排名, 后排输出, 后排输出排名, 好友支援, 好友支援排名, 前排耐久, 前排耐久排名, 后排耐久, 后排耐久排名,
            后排输出同属性, 后排输出同属性排名, 好友支援同属性, 好友支援同属性排名,
            单前排奶盾能力, 单前排奶盾能力排名, 双前排奶盾能力, 双前排奶盾能力排名,
            奶盾中前排输出排名, 奶盾中后排输出排名,
            特殊评分,
        } = a;
        const ratings = [];
        const results = {
            no, id,
            评分: null,
            分项能力: null,
        }
        for (const { border, rating, description } of rating_constants.rating_borders["前排输出"]) {
            if (前排输出 / rating_constants.absv >= border) {
                ratings.push({ type: "前排输出", rate: rating });
                break;
            }
        }
        for (const { border, rating, description } of rating_constants.rating_borders["前排真输出"]) {
            if (前排真输出 / rating_constants.absvr >= border) {
                ratings.push({ type: "前排真输出", rate: rating });
                break;
            }
        }
        {
            const local_ratings = [];
            for (const sp of 特殊评分) {
                if (sp.type === "后排输出") local_ratings.push(sp);
            }
            for (const { border, rating, description } of rating_constants.rating_borders["后排输出"]) {
                if (后排输出 / rating_constants.absv >= border) {
                    local_ratings.push({ type: "后排输出", rate: rating });
                    break;
                }
            }
            for (const { border, rating, description } of rating_constants.rating_borders["后排输出同属性"]) {
                if (后排输出同属性 / rating_constants.absv >= border) {
                    local_ratings.push({ type: "后排输出同属性", rate: rating });
                    break;
                }
            }
            if (local_ratings[0]) ratings.push(local_ratings.sort((a, b) => b.rate - a.rate)[0]);

        }
        if (单前排奶盾能力) {
            const local_ratings = [];
            for (const { border, rating, description } of rating_constants.rating_borders["单前排奶盾能力"]) {
                if (单前排奶盾能力 >= border) {
                    local_ratings.push({ type: "单前排奶盾能力", rate: rating });
                    break;
                }
            }
            for (const { border, rating, description } of rating_constants.rating_borders["双前排奶盾能力"]) {
                if (双前排奶盾能力 >= border) {
                    local_ratings.push({ type: "双前排奶盾能力", rate: rating });
                    break;
                }
            }
            if (local_ratings[0]) ratings.push(local_ratings.sort((a, b) => b.rate - a.rate)[0]);
        }
        {
            const local_ratings = [];
            for (const sp of 特殊评分) {
                if (sp.type === "后排耐久") local_ratings.push(sp);
            }
            for (const { border, rating, description } of rating_constants.rating_borders["后排耐久"]) {
                if (a["后排耐久"] >= border) {
                    local_ratings.push({ type: "后排耐久", rate: rating });
                    break;
                }
            }
            if (local_ratings[0]) ratings.push(local_ratings.sort((a, b) => b.rate - a.rate)[0]);
        }
        特殊评分
            .filter(({ type }) => ["前排输出", "前排真输出", "前排耐久", "单前排奶盾能力", "双前排奶盾能力", "后排输出", "后排输出同属性", "后排耐久", "好友支援", "好友支援同属性"].indexOf(type) === -1)
            .forEach((sp) => ratings.push(sp));
        //总评分和参与合成的最低评分
        //评分合成只能从高到低，不能从低到高
        const [lower_rating, final_rating] = ratings.sort((a, b) => b.rate - a.rate).reduce((prev, curr) => {
            if (prev[1] < 0) return [curr.rate, curr.rate];
            if (prev[1] - curr.rate <= 1) {
                prev[1] += prev[1] >= 6 ? 0.5 : 1;
                prev[0] = curr.rate;
                return prev;
            }
            return prev;
        }, [-114514, -1919810]);
        results.评分 = `<p><span style="font-weight:bold;${final_rating >= 9 ? "color:#ee230d;" : ""}">${final_rating.toFixed(1)}</span>，${rating_constants.rating_descriptions.get(final_rating)}</p>`;

        results.分项能力 = ["前排输出", "前排真输出", "前排耐久", "单前排奶盾能力", "双前排奶盾能力", "后排输出", "后排输出同属性", "后排耐久", "好友支援", "好友支援同属性", "特殊评分"].map((type) => {
            if (type === "后排输出同属性" && 后排输出 === 后排输出同属性) return null;
            if (type === "好友支援同属性" && 好友支援 === 好友支援同属性) return null;
            if (["前排输出", "后排输出", "后排输出同属性", "好友支援", "好友支援同属性"].indexOf(type) !== -1) {
                for (const { border, rating, description } of rating_constants.rating_borders[type]) {
                    if (a[type] >= border * rating_constants.absv) {
                        return `${type}　<span style="font-weight:bold;${rating >= 9 ? "color:#ee230d;" : ""}">${description || "特殊"}</span>`
                            + (rating ? `(<span${rating >= lower_rating ? ' style="color:#ff654e;"' : ''}>${rating.toFixed(1)}分</span>)` : "")
                            + `，<span style="font-weight:bold;">${parseInt(a[type] / rating_constants.absv * 1000).toString().replace(/([0-9]*)([0-9])/, "$1.$2")}</span>(${parseInt(a[type])}, ${a[type + "排名"]}位${a[`奶盾中${type}排名`] ? `, 奶盾中${a[`奶盾中${type}排名`]}位` : ""})`;
                    }
                }
            }
            else if (["前排耐久", "单前排奶盾能力", "双前排奶盾能力", "后排耐久"].indexOf(type) !== -1) {
                for (const { border, rating, description } of rating_constants.rating_borders[type]) {
                    if (a[type] >= border) {
                        return `${type}　<span style="font-weight:bold;${rating >= 9 ? "color:#ee230d;" : ""}">${description || "特殊"}</span>`
                            + (rating ? `(<span${rating >= lower_rating ? ' style="color:#ff654e;"' : ''}>${rating.toFixed(1)}分</span>)` : "")
                            + `，<span style="font-weight:bold;">${a[type] > 1000 ? parseInt(a[type]) : parseInt(a[type] * 10).toString().replace(/([0-9]*)([0-9])/, "$1.$2")}</span>(${a[type + "排名"]}位)`;
                    }
                }
            }
            else if (type === "前排真输出") {
                for (const { border, rating, description } of rating_constants.rating_borders[type]) {
                    if (a[type] >= border * rating_constants.absvr) {
                        return `${type}　<span style="font-weight:bold;${rating >= 9 ? "color:#ee230d;" : ""}">${description}</span>(<span${rating >= lower_rating ? ' style="color:#ff654e;"' : ''}>${rating.toFixed(1)}分</span>)`
                            + `，<span style="font-weight:bold;">${parseInt(a[type] / rating_constants.absvr * 1000).toString().replace(/([0-9]*)([0-9])/, "$1.$2")}</span>(${parseInt(a[type])}, ${a[type + "排名"]}位)`;
                    }
                }
            }
            else if (type === "特殊评分") {
                return 特殊评分
                    .filter(({ type }) => ["前排输出", "前排真输出", "前排耐久", "单前排奶盾能力", "双前排奶盾能力", "后排输出", "后排输出同属性", "后排耐久", "好友支援"].indexOf(type) === -1)
                    .map(({ type, rate }) => `${type}　<span style="font-weight:bold;${rate >= 9 ? "color:#ee230d;" : ""}">特殊</span>(<span${rate >= lower_rating ? ' style="color:#ff654e;"' : ''}>${rate.toFixed(1)}分</span>)`)
                    .join("<br>");
            }
            return null;
        }).filter((v) => v).join("<br>");

        if (process.argv.indexOf("en") !== -1) {
            results.评分 = results.评分
                .replace(/([0-9](.[0-9])?)分/g, "$1")
                .replace(/，神/g, "，Legendary")
                .replace(/，准神/g, "，Godlike")
                .replace(/，超强/g, "，Powerful +")
                .replace(/，强力/g, "，Powerful")
                .replace(/，较强/g, "，Powerful -")
                .replace(/，可用/g, "，Available")
                .replace(/，较弱/g, "，Soso")
                .replace(/，弱/g, "，Weak");

            results.分项能力 = results.分项能力
                .replace(/前排输出/g, "Frontline Voltage")
                .replace(/前排真输出/g, "60000-Cap Voltage")
                .replace(/前排耐久/g, "Frontline Recovery/Shield")
                .replace(/单前排奶盾能力/g, "Healer General (Single Strategy)")
                .replace(/双前排奶盾能力/g, "Healer General (Double Strategy)")
                .replace(/后排输出同属性/g, "Backline Voltage (Same Attribute)")
                .replace(/后排输出/g, "Backline Voltage")
                .replace(/后排耐久/g, "Backline Recovery/Shield")
                .replace(/好友支援同属性/g, "Friend Support")
                .replace(/好友支援/g, "Friend Support")
                .replace(/sk充电/g, "Skill Typed SPG")
                .replace(/sk奶/g, "Skill Typed Healer")
                .replace(/特技后排/g, "Backline Active Skill")
                .replace(/SP后排/g, "Backline SP Gauge")
                .replace(/([0-9](.[0-9])?)分/g, "$1")
                .replace(/特殊/g, "Special")
                .replace(/(?<=([0456789])|(1[123]))位/g, "th")
                .replace(/(?<=1)位/g, "st")
                .replace(/(?<=2)位/g, "nd")
                .replace(/(?<=3)位/g, "rd")
        }

        return results;
    }
}

const memid_to_fullname = { 1: "高坂穗乃果", 2: "绚濑绘里", 3: "南小鸟", 4: "园田海未", 5: "星空凛", 6: "西木野真姬", 7: "东条希", 8: "小泉花阳", 9: "矢泽妮可", 101: "高海千歌", 102: "樱内梨子", 103: "松浦果南", 104: "黑泽黛雅", 105: "渡边曜", 106: "津岛善子", 107: "国木田花丸", 108: "小原鞠莉", 109: "黑泽露比", 201: "上原步梦", 202: "中须霞", 203: "樱坂雫", 204: "朝香果林", 205: "宫下爱", 206: "近江彼方", 207: "优木雪菜", 208: "艾玛·维尔德", 209: "天王寺璃奈", 210: "三船栞子", 211: "米娅·泰勒", 212: "钟岚珠" };
get_data().then((card_attrs_new) => {
    //readFile
    const comments = (() => {
        const { map: comments, context } = fs.readFileSync("./评价.txt").toString().split(/[\r\n]{1,2}/).reduce(({ map, context }, text) => {
            const id = parseInt(text);
            if (isNaN(id)) {
                context.push(text);
                return { map, context };
            } else {
                let s = context.pop();
                while (s === "" && context.length) {
                    s = context.pop();
                }
                context.push(s);
                map.set(context[0], context.slice(1));
                return { map, context: [id] };
            }
        }, { map: new Map(), context: [] });
        let s = context.pop();
        while (s === "" && context.length) {
            s = context.pop();
        }
        context.push(s);
        comments.set(context[0], context.slice(1));
        return comments;
    })();

    console.log(comments);

    const texts = card_attrs_new.map(rating_constants.rating).map(({ id, 评分, 分项能力 }) => {
        const text = `<div style='font-size:20px;padding-left:24px;'>
        <div style="display:grid;grid-template:auto / 120px 1fr;">
        <div style='color:#02a2ff;font-weight:bold;'><p>评分: </p></div>
        <div style="">${评分}</div>
        </div>
        
        <div style="display:grid;grid-template:auto / 120px 1fr;">
        <div style='color:#02a2ff;font-weight:bold;'><p>分项能力: </p></div>
        <div style=""><p>${分项能力}</p></div>
        </div>
        
        <div style="display:grid;grid-template:auto / 80px 1fr;line-height:160%;">
        <div style='color:#02a2ff;font-weight:bold;'><p>评价: </p></div>
        <div style="white-space:pre-wrap;">${comments.get(id) ? comments.get(id).map((s) => s.length ? `<p>　　${s}</p>` : "<br>").join("") : ("寄").repeat(384)}</div>
        </div>

        <p style="text-align:right;opacity:30%;color:#ef5fa8;padding-right:12px;font-size:70%;">
            数据支持: SIFAS综合数据.xlsx (by 潜水/CoffeePot)
            <br>评价提供者: ${comments.get(0).map((s) => s.replace(/(\s|\n)/, "")).filter((s) => s.length).join("、")}
            <br>${Intl.DateTimeFormat({}, { year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date())}
        </p>
    </div>
    `;
        return { id, text };
    })
    return texts;
}).then(async (texts) => {
    console.log("start browser");

    const browser = await puppeteer.launch({ /*executablePath: CHROME_PATH,*/ headless: false });
    for (const { id, text } of texts.values()) {
        const page = await browser.newPage();
        await page.setViewport({ width: 960, height: 1280, deviceScaleFactor: 2, });

        //阻止所有图片加载
        await page.setRequestInterception(true);
        page.on('request', (req) => {
            if (req.resourceType() === "image") req.abort();
            else req.continue();
        })
        if (process.argv.indexOf("purge") !== -1) {
            await page.goto(`https://wiki.loveliv.es/${memid_to_fullname[parseInt(id.toString().slice(2, 5))]}?action=purge`, { timeout: 114514, waitUntil: "domcontentloaded" });
            await Promise.all([
                page.click("button"),
                page.waitForNavigation({ waitUntil: "domcontentloaded" }),
            ])
        } else {
            await page.goto(`https://wiki.loveliv.es/${memid_to_fullname[parseInt(id.toString().slice(2, 5))]}`, { timeout: 114514, waitUntil: "domcontentloaded" });
        }

        //恢复所有图片加载
        page.removeAllListeners("request");
        await page.setRequestInterception(false);

        const url_list = await page.evaluate((id, tx) => {
            const t = document.querySelector(`#card_long_id_${id.toString()} + h4 + table`);
            const url_list = [];
            document.querySelectorAll("body > *").forEach((a) => a.remove());
            document.querySelector("body").attributeStyleMap.set("overflow-x", "hidden");
            document.querySelector("body").attributeStyleMap.set("overflow-y", "hidden");
            document.querySelector("body").appendChild(t);
            //reload
            document.querySelectorAll("img").forEach((n) => {
                const src = n.getAttribute("src");
                n.setAttribute("src", src);
                url_list.push(src);
            });
            //蓝色标题02a2ff 红色评分ee230d
            document.querySelector("body").innerHTML += tx
                + '<div style="position:absolute;top:20%;left:-2%;font-size:128px;transform:rotate(37deg);font-weight:bold;color:blue;opacity:2.5%;white-space:pre;z-index:0;overflow:hidden;">LoveLive! AS Wiki</div>'
                + '<div style="position:absolute;top:70%;left:-2%;font-size:128px;transform:rotate(37deg);font-weight:bold;color:blue;opacity:2.5%;white-space:pre;z-index:0;overflow:hidden;">LoveLive! AS Wiki</div>';
            return url_list;
        }, id, text);

        await Promise.all(url_list.map((url) => page.waitForResponse(async (res) => { return res.url().match(url) && await res.buffer(); }, { timeout: 114514 })));

        await page.screenshot({ path: `./${id}.png`, fullPage: true });
        console.log(id, "OK");

        await page.close();
    }
    await browser.close();
});
