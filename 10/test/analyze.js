const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

function decodeXml(str) {
    if (!str) return '';
    return str.replace(/&amp;/g, '&')
              .replace(/&lt;/g, '<')
              .replace(/&gt;/g, '>')
              .replace(/&quot;/g, '"')
              .replace(/&apos;/g, "'");
}

function colLetterToNum(colLetter) {
    let num = 0;
    for (let i = 0; i < colLetter.length; i++) {
        num = num * 26 + (colLetter.charCodeAt(i) - 64);
    }
    return num - 1;
}

function parseXlsx(filePath) {
    let sharedStrings = [];
    try {
        const sharedStringsXml = execSync(`unzip -p "${filePath}" xl/sharedStrings.xml`, { stdio: ['pipe', 'pipe', 'ignore'] }).toString();
        const siRegex = /<si>([\s\S]*?)<\/si>/g;
        let siMatch;
        while ((siMatch = siRegex.exec(sharedStringsXml)) !== null) {
            const siContent = siMatch[1];
            const tRegex = /<t[^>]*>([\s\S]*?)<\/t>/g;
            let tMatch;
            let val = "";
            while ((tMatch = tRegex.exec(siContent)) !== null) {
                val += tMatch[1];
            }
            sharedStrings.push(decodeXml(val));
        }
    } catch (e) {}

    try {
        const sheetXml = execSync(`unzip -p "${filePath}" xl/worksheets/sheet1.xml`, { stdio: ['pipe', 'pipe', 'ignore'] }).toString();
        const rowRegex = /<row [^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
        let rowMatch;
        const rows = [];
        
        while ((rowMatch = rowRegex.exec(sheetXml)) !== null) {
            const rowNum = parseInt(rowMatch[1], 10) - 1;
            const rowContent = rowMatch[2];
            const cellRegex = /<c [^>]*r="([A-Z]+)(\d+)"([^>]*)>([\s\S]*?)<\/c>/g;
            let cellMatch;
            const rowData = [];
            
            while ((cellMatch = cellRegex.exec(rowContent)) !== null) {
                const colLetter = cellMatch[1];
                const colIdx = colLetterToNum(colLetter);
                const attrs = cellMatch[3];
                const cellValXml = cellMatch[4];
                const vMatch = /<v>([\s\S]*?)<\/v>/.exec(cellValXml);
                let rawVal = vMatch ? vMatch[1] : "";
                const isString = attrs.includes('t="s"');
                let finalVal = rawVal;
                if (isString && rawVal !== "") {
                    const idx = parseInt(rawVal, 10);
                    finalVal = sharedStrings[idx] || "";
                } else if (rawVal !== "") {
                    const numVal = Number(rawVal);
                    if (!isNaN(numVal)) finalVal = numVal;
                }
                rowData[colIdx] = finalVal;
            }
            rows[rowNum] = rowData;
        }
        
        const cleanedRows = [];
        for (let i = 0; i < rows.length; i++) {
            cleanedRows.push(rows[i] || []);
        }
        return cleanedRows;
    } catch (e) {
        console.error(e);
        return [];
    }
}

function normalizeRegion(name) {
    if (!name) return "";
    name = name.trim();
    if (name === "전라북도" || name === "전북특별자치도") return "전북특별자치도";
    if (name === "제주도" || name === "제주특별자치도") return "제주특별자치도";
    return name;
}

const dataDir = "/Users/bagchaeeun/Desktop/test/data";
const satisfactionFile = path.join(dataDir, "삶의_만족도_시도__20260606195059.xlsx");
const suicideFile = path.join(dataDir, "인구십만명당_자살률_시도_시_군_구__20260606194913.xlsx");

const satRows = parseXlsx(satisfactionFile);
const suiRows = parseXlsx(suicideFile);

const row1 = satRows[0];
const row2 = satRows[1];
const colYearMapping = [];
const colTypeMapping = [];

for (let colIdx = 3; colIdx < row1.length; colIdx++) {
    const yearStr = row1[colIdx];
    const typeStr = row2[colIdx];
    colYearMapping[colIdx] = parseInt(yearStr, 10);
    colTypeMapping[colIdx] = typeStr;
}

const satData = [];
let currentSatRegion = "";
let currentTrait1 = "";

for (let r = 2; r < satRows.length; r++) {
    const row = satRows[r];
    if (!row || row.length === 0) continue;
    
    if (row[0] && row[0].trim()) {
        currentSatRegion = normalizeRegion(row[0]);
    }
    
    if (row[1] && row[1].trim()) {
        currentTrait1 = row[1].trim();
    }
    
    const trait1 = currentTrait1;
    const trait2 = row[2]; // '계' or '남자' or '여자'
    
    let gender = "전체";
    if (trait1 === "성별") {
        gender = trait2; // '남자' or '여자'
    } else if (trait1 === "전체" && trait2 === "계") {
        gender = "전체";
    } else {
        continue;
    }
    
    const yearValues = {};
    for (let colIdx = 3; colIdx < row.length; colIdx++) {
        const year = colYearMapping[colIdx];
        const type = colTypeMapping[colIdx];
        const val = row[colIdx];
        
        if (!year || !type) continue;
        if (!yearValues[year]) yearValues[year] = {};
        
        yearValues[year][type] = (typeof val === 'number') ? val : parseFloat(val) || 0;
    }
    
    for (const year in yearValues) {
        const vals = yearValues[year];
        const verySatisfied = vals["매우 만족"] || 0;
        const slightlySatisfied = vals["약간 만족"] || 0;
        const neutral = vals["보통"] || 0;
        const slightlyDissatisfied = vals["약간 불만족"] || 0;
        const veryDissatisfied = vals["매우 불만족"] || 0;
        
        const sumRatios = verySatisfied + slightlySatisfied + neutral + slightlyDissatisfied + veryDissatisfied;
        
        let score = null;
        if (sumRatios > 0) {
            // weighted average score (1~5 scale)
            score = (verySatisfied * 5 + slightlySatisfied * 4 + neutral * 3 + slightlyDissatisfied * 2 + veryDissatisfied * 1) / 100;
        }
        
        satData.push({
            region: currentSatRegion,
            gender: gender,
            year: parseInt(year, 10),
            score: score,
            positive: verySatisfied + slightlySatisfied,
            negative: slightlyDissatisfied + veryDissatisfied,
            raw: vals
        });
    }
}

const suiRow1 = suiRows[0];
const suiRow2 = suiRows[1];
const suiColYearMapping = [];
const suiColGenderMapping = [];

for (let colIdx = 1; colIdx < suiRow1.length; colIdx++) {
    suiColYearMapping[colIdx] = parseInt(suiRow1[colIdx], 10);
    suiColGenderMapping[colIdx] = suiRow2[colIdx];
}

const suiData = [];

for (let r = 2; r < suiRows.length; r++) {
    const row = suiRows[r];
    if (!row || row.length === 0) continue;
    
    const rawRegion = row[0];
    if (!rawRegion) continue;
    const region = normalizeRegion(rawRegion);
    
    for (let colIdx = 1; colIdx < row.length; colIdx++) {
        const year = suiColYearMapping[colIdx];
        const genderRaw = suiColGenderMapping[colIdx];
        const rateVal = row[colIdx];
        
        if (!year || !genderRaw) continue;
        
        let gender = "전체";
        if (genderRaw === "남자") gender = "남자";
        if (genderRaw === "여자") gender = "여자";
        
        suiData.push({
            region: region,
            gender: gender,
            year: year,
            rate: (typeof rateVal === 'number') ? rateVal : parseFloat(rateVal) || 0
        });
    }
}

const joinedMap = new Map();
satData.forEach(sat => {
    const key = `${sat.region}_${sat.gender}_${sat.year}`;
    joinedMap.set(key, { sat });
});

suiData.forEach(sui => {
    const key = `${sui.region}_${sui.gender}_${sui.year}`;
    const existing = joinedMap.get(key);
    if (existing) {
        existing.sui = sui;
    }
});

const joinedData = [];
for (const [key, value] of joinedMap.entries()) {
    if (value.sat && value.sui) {
        joinedData.push({
            region: value.sat.region,
            gender: value.sat.gender,
            year: value.sat.year,
            satScore: value.sat.score,
            positive: value.sat.positive,
            negative: value.sat.negative,
            suicideRate: value.sui.rate
        });
    }
}

console.log(`Total joined data points: ${joinedData.length}`);

function getPearsonCorrelation(x, y) {
    const n = x.length;
    if (n === 0) return 0;
    let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0, sumY2 = 0;
    for (let i = 0; i < n; i++) {
        sumX += x[i];
        sumY += y[i];
        sumXY += x[i] * y[i];
        sumX2 += x[i] * x[i];
        sumY2 += y[i] * y[i];
    }
    const num = n * sumXY - sumX * sumY;
    const den = Math.sqrt((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY));
    if (den === 0) return 0;
    return num / den;
}

const regionData = joinedData.filter(d => d.region !== "전국");

function analyzeSubset(subset, label) {
    const xScore = subset.map(d => d.satScore);
    const xPos = subset.map(d => d.positive);
    const xNeg = subset.map(d => d.negative);
    const yRate = subset.map(d => d.suicideRate);
    
    const corrScore = getPearsonCorrelation(xScore, yRate);
    const corrPos = getPearsonCorrelation(xPos, yRate);
    const corrNeg = getPearsonCorrelation(xNeg, yRate);
    
    return {
        label,
        count: subset.length,
        corrScore: corrScore.toFixed(4),
        corrPos: corrPos.toFixed(4),
        corrNeg: corrNeg.toFixed(4)
    };
}

console.log("\n=== 1. Correlation Analysis (Region-level pooled, 2020-2024) ===");
const overallAll = analyzeSubset(regionData.filter(d => d.gender === "전체"), "Overall (Total)");
const overallMale = analyzeSubset(regionData.filter(d => d.gender === "남자"), "Overall (Male)");
const overallFemale = analyzeSubset(regionData.filter(d => d.gender === "여자"), "Overall (Female)");

console.table([overallAll, overallMale, overallFemale]);

console.log("\n=== 2. Correlation Analysis by Year (Total Gender) ===");
const yearlyCorrs = [];
for (let y = 2020; y <= 2024; y++) {
    const yData = regionData.filter(d => d.year === y && d.gender === "전체");
    yearlyCorrs.push(analyzeSubset(yData, `Year ${y}`));
}
console.table(yearlyCorrs);

console.log("\n=== 3. Cross-sectional Correlation (5-Year Average per Region) ===");
const uniqueRegions = Array.from(new Set(regionData.map(d => d.region)));
const genderOptions = ["전체", "남자", "여자"];
const crossSectionalResults = [];

genderOptions.forEach(g => {
    const avgList = [];
    uniqueRegions.forEach(reg => {
        const rData = regionData.filter(d => d.region === reg && d.gender === g);
        if (rData.length === 0) return;
        
        const avgSatScore = rData.reduce((sum, d) => sum + d.satScore, 0) / rData.length;
        const avgPos = rData.reduce((sum, d) => sum + d.positive, 0) / rData.length;
        const avgNeg = rData.reduce((sum, d) => sum + d.negative, 0) / rData.length;
        const avgSuiRate = rData.reduce((sum, d) => sum + d.suicideRate, 0) / rData.length;
        
        avgList.push({
            region: reg,
            satScore: avgSatScore,
            positive: avgPos,
            negative: avgNeg,
            suicideRate: avgSuiRate
        });
    });
    
    const xScore = avgList.map(d => d.satScore);
    const xPos = avgList.map(d => d.positive);
    const xNeg = avgList.map(d => d.negative);
    const yRate = avgList.map(d => d.suicideRate);
    
    const corrScore = getPearsonCorrelation(xScore, yRate);
    const corrPos = getPearsonCorrelation(xPos, yRate);
    const corrNeg = getPearsonCorrelation(xNeg, yRate);
    
    crossSectionalResults.push({
        Gender: g,
        corrScore: corrScore.toFixed(4),
        corrPos: corrPos.toFixed(4),
        corrNeg: corrNeg.toFixed(4)
    });
});
console.table(crossSectionalResults);

// Region-wise average lists for each gender to print in report
const genderAverages = {};
genderOptions.forEach(g => {
    const avgList = [];
    uniqueRegions.forEach(reg => {
        const rData = regionData.filter(d => d.region === reg && d.gender === g);
        if (rData.length === 0) return;
        const avgSatScore = rData.reduce((sum, d) => sum + d.satScore, 0) / rData.length;
        const avgPos = rData.reduce((sum, d) => sum + d.positive, 0) / rData.length;
        const avgNeg = rData.reduce((sum, d) => sum + d.negative, 0) / rData.length;
        const avgSuiRate = rData.reduce((sum, d) => sum + d.suicideRate, 0) / rData.length;
        
        avgList.push({
            region: reg,
            satScore: avgSatScore,
            positive: avgPos,
            negative: avgNeg,
            suicideRate: avgSuiRate
        });
    });
    genderAverages[g] = avgList;
});

// Let's write out some summarized results to a JSON file so we can build the report easily
const reportStats = {
    overall: { overallAll, overallMale, overallFemale },
    yearly: yearlyCorrs,
    crossSectional: crossSectionalResults,
    genderAverages: genderAverages,
    nationalTrends: joinedData.filter(d => d.region === "전국").sort((a,b) => a.year - b.year)
};

fs.writeFileSync(
    "/Users/bagchaeeun/.gemini/antigravity/brain/665c397e-0ba7-4a4c-904e-89e5f8c4ecad/scratch/report_stats.json",
    JSON.stringify(reportStats, null, 2)
);
console.log("Stats exported to report_stats.json");
