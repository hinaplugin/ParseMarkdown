// data/parse.js
import fs from "fs";
import path from "path";
import mammoth from "mammoth";
import TurndownService from "turndown";
import XLSX from "xlsx";
import PPTX2Json from "pptx2json";

const inDir = path.resolve("file");
const outDir = path.resolve("out");

if (!fs.existsSync(outDir)) {
    fs.mkdirSync(inDir, { recursive: true });
    fs.mkdirSync(outDir, { recursive: true });
}

const turndown = new TurndownService({
    headingStyle: "atx",
    bulletListMarker: "-",
    codeBlockStyle: "fenced"
});

async function convertDocx(inPath, outPath) {
    const result = await mammoth.convertToHtml({ path: inPath });
    const md = turndown.turndown(result.value);
    fs.writeFileSync(outPath, md, "utf-8");
}

function convertXlsx(inPath, outPath) {
    const wb = XLSX.readFile(inPath);
    let md = "";

    for (const sheetName of wb.SheetNames) {
        md += `# ${sheetName}\n\n`;
        const sheet = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (rows.length > 0) {
            const header = rows[0];
            md += `| ${header.join(" | ")} |\n`;
            md += `| ${header.map(() => "---").join(" | ")} |\n`;

            for (let i = 1; i < rows.length; i++) {
                md += `| ${rows[i].join(" | ")} |\n`;
            }
            md += "\n";
        }
    }

    fs.writeFileSync(outPath, md, "utf-8");
}

async function convertPptx(inPath, outPath) {
    const parser = new PPTX2Json();
    const json = await parser.toJson(inPath);
    let md = "";

    json.slides.forEach((slide, i) => {
        md += `# Slide ${i + 1}\n\n`;
        slide.texts.forEach(t => {
            md += `${t.text}\n\n`;
        });
    });

    fs.writeFileSync(outPath, md, "utf-8");
}

async function convertAll() {
    const files = fs.readdirSync(inDir)
        .filter(name => /\.(docs|docx|xlsx|pptx)$/i.test(name));

    if (files.length === 0) {
        console.log("変換対象のファイルがありません");
        return;
    }

    for (const name of files) {
        const inPath = path.join(inDir, name);
        const outName = name.replace(/\.(docs|docx|xlsx|pptx)$/i, ".md");
        const outPath = path.join(outDir, outName);

        try {
            if (/\.(docs|docx)$/i.test(name)) {
                await convertDocx(inPath, outPath);
            } else if (/\.xlsx$/i.test(name)) {
                convertXlsx(inPath, outPath);
            } else if (/\.pptx$/i.test(name)) {
                await convertPptx(inPath, outPath);
            }

            console.log(`変換完了: ${name} -> ${outName}`);
        } catch (e) {
            console.error(`失敗: ${name}`, e);
        }
    }
}

convertAll();
