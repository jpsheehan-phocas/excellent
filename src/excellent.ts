export type Workbook = {
    sheets: Spreadsheet[];
    author: string;
}

export enum AlignHorizontal {
    Default,
    Left,
    Center,
    Right
}

export enum CellType {
    String,
    Number
}

export type BaseCell = {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    span?: number;
    alignHorizontal?: AlignHorizontal,
}

export type Cell =
    (BaseCell & { type?: undefined, data: string }) |
    (BaseCell & { type?: CellType.String, data: string }) |
    (BaseCell & { type?: CellType.Number, data: number })

export type Row = {
    cells: Cell[];
}

export type Spreadsheet = {
    title: string;
    rows: Row[];
}

declare class JSZip {
    file(filename: string, content: string): void;
    generateAsync(options: { type: string }): Promise<Uint8Array>;
}

function generateFiles(workbook: Workbook): { [key: string]: string } {
    const contentTypesForSheets = workbook.sheets
        .map((_, i) =>
            `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />`)
        .join("\n");
    const titlesOfParts = workbook.sheets
        .map((sheet) =>
            `<vt:lpstr>${sheet.title}</vt:lpstr>`)
        .join("\n");
    const sheets = workbook.sheets.map((sheet, i) => `<sheet name="${sheet.title}" sheetId="${i + 1}" r:id="rId${i + 3}" />`);
    const workbookRelationships = workbook.sheets
        .map((_, i) => `<Relationship Id="rId${3 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml" />`)
        .join("\n");
    const strings: string[] = [];

    const styleMapping = getStyleMapping();

    workbook.sheets
        .flatMap(sheet => sheet.rows)
        .flatMap(row => row.cells)
        .filter(isString)
        .map(cell => cell.data.toString())
        .forEach(content => {
            if (!strings.includes(content)) {
                strings.push(content);
            }
        });

    const stringData = strings.map(str => `<si><t>${str}</t></si>`).join("\n");

    const now = new Date().toISOString();

    const fonts = [];
    for (let bold of [false, true]) {
        for (let italic of [false, true]) {
            for (let underline of [false, true]) {
                fonts.push(`<font>
                    ${bold ? "<b />" : ""} ${italic ? "<i />" : ""} ${underline ? "<u />" : ""}
                    <sz val="11" />
                    <color theme="1" />
                    <name val="Aptos Narrow" />
                    <family val="2" />
                    <scheme val="minor" />
                </font>`)
            }
        }
    }

    const cellXfs = [...styleMapping.keys()].map((key) => {
        const [fontId, alignHorizontal] = key.split("-").map(x => parseInt(x));
        const applyAlignment = alignHorizontal !== AlignHorizontal.Default ? "1" : "0";
        let alignment = "";

        if (applyAlignment === "1") {
            let alignmentText = "";
            switch (alignHorizontal) {
                case AlignHorizontal.Center:
                    alignmentText = "center";
                    break;
                case AlignHorizontal.Left:
                    alignmentText = "left";
                    break;
                case AlignHorizontal.Right:
                    alignmentText = "right";
                    break;
            }
            alignment = `<alignment horizontal="${alignmentText}" />`;
        }

        return `<xf numFmtId="0" fontId="${fontId}" fillId="0" borderId="0" xfId="0" applyFont="${fontId === 0 ? '0' : '1'}" applyAlignment="${applyAlignment}">
            ${alignment}
        </xf>`;
    });

    const files: { [key: string]: string } = {
        "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
    <Default Extension="xml" ContentType="application/xml" />
    <Override PartName="/xl/workbook.xml"
        ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
    ${contentTypesForSheets}
    <Override PartName="/xl/styles.xml"
        ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
    <Override PartName="/xl/sharedStrings.xml"
        ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />
    <Override PartName="/docProps/core.xml"
        ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
    <Override PartName="/docProps/app.xml"
        ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
</Types>`,
        "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId3"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
                Target="docProps/app.xml" />
            <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                Target="docProps/core.xml" />
            <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="xl/workbook.xml" />
        </Relationships>`,
        "docProps/app.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
            <Application>Microsoft Excel</Application>
            <DocSecurity>0</DocSecurity>
            <ScaleCrop>false</ScaleCrop>
            <HeadingPairs>
                <vt:vector size="2" baseType="variant">
                    <vt:variant>
                        <vt:lpstr>Worksheets</vt:lpstr>
                    </vt:variant>
                    <vt:variant>
                        <vt:i4>1</vt:i4>
                    </vt:variant>
                </vt:vector>
            </HeadingPairs>
            <TitlesOfParts>
                <vt:vector size="${workbook.sheets.length}" baseType="lpstr">
                    ${titlesOfParts}
                </vt:vector>
            </TitlesOfParts>
            <Company></Company>
            <LinksUpToDate>false</LinksUpToDate>
            <SharedDoc>false</SharedDoc>
            <HyperlinksChanged>false</HyperlinksChanged>
            <AppVersion>16.0300</AppVersion>
        </Properties>`,
        "docProps/core.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <cp:coreProperties
            xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
            xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
            xmlns:dcmitype="http://purl.org/dc/dcmitype/"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <dc:creator>${workbook.author}</dc:creator>
            <cp:lastModifiedBy>${workbook.author}</cp:lastModifiedBy>
            <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
            <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
        </cp:coreProperties>`,
        "xl/sharedStrings.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">
            ${stringData}
        </sst>`,
        "xl/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="x14ac x16r2 xr"
            xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
            xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"
            xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
            <fonts count="${fonts.length}" x14ac:knownFonts="1">
                ${fonts.join("\n")}
            </fonts>
            <fills count="2">
                <fill>
                    <patternFill patternType="none" />
                </fill>
                <fill>
                    <patternFill patternType="gray125" />
                </fill>
            </fills>
            <borders count="1">
                <border>
                    <left />
                    <right />
                    <top />
                    <bottom />
                    <diagonal />
                </border>
            </borders>
            <cellStyleXfs count="1">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
            </cellStyleXfs>
            <cellXfs count="${cellXfs.length}">
                ${cellXfs.join("\n")}
            </cellXfs>
            <cellStyles count="1">
                <cellStyle name="Normal" xfId="0" builtinId="0" />
            </cellStyles>
            <dxfs count="0" />
            <tableStyles count="0" defaultTableStyle="TableStyleMedium2"
                defaultPivotStyle="PivotStyleLight16" />
            <extLst>
                <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"
                    xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
                    <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1" />
                </ext>
                <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}"
                    xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
                    <x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1" />
                </ext>
            </extLst>
        </styleSheet>`,
        "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="x15 xr xr6 xr10 xr2"
            xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
            xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
            xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"
            xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"
            xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">
            <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27425" />
            <workbookPr defaultThemeVersion="202300" />
            <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
                <mc:Choice Requires="x15">
                    <x15ac:absPath url="C:\\Users\\jesse.sheehan\\Documents\\"
                        xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac" />
                </mc:Choice>
            </mc:AlternateContent>
            <xr:revisionPtr revIDLastSave="0" documentId="8_{1757FEB2-880B-432D-971E-DFCBE4FADE3B}"
                xr6:coauthVersionLast="47" xr6:coauthVersionMax="47"
                xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}" />
            <bookViews>
                <workbookView xWindow="-120" yWindow="-120" windowWidth="29040" windowHeight="15720"
                    xr2:uid="{1AA42C6E-2B1E-4349-B114-01DED749F222}" />
            </bookViews>
            <sheets>
                ${sheets}
            </sheets>
            <calcPr calcId="191029" />
            <extLst>
                <ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}"
                    xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
                    <x15:workbookPr chartTrackingRefBase="1" />
                </ext>
                <ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}"
                    xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures">
                    <xcalcf:calcFeatures>
                        <xcalcf:feature name="microsoft.com:RD" />
                        <xcalcf:feature name="microsoft.com:Single" />
                        <xcalcf:feature name="microsoft.com:FV" />
                        <xcalcf:feature name="microsoft.com:CNMTM" />
                        <xcalcf:feature name="microsoft.com:LET_WF" />
                        <xcalcf:feature name="microsoft.com:LAMBDA_WF" />
                        <xcalcf:feature name="microsoft.com:ARRAYTEXT_WF" />
                    </xcalcf:calcFeatures>
                </ext>
            </extLst>
        </workbook>`,
        "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                Target="styles.xml" />
            <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
                Target="sharedStrings.xml" />
            ${workbookRelationships}
        </Relationships>`,
    };

    workbook.sheets.forEach((sheet, i) => {
        const filename = `xl/worksheets/sheet${i + 1}.xml`;
        const mergedCells: [string, string][] = [];
        const rowData = sheet.rows
            .map((row, rowIdx) => {
                let columnIdx = 0;
                const cellData = row.cells
                    .flatMap((cell) => {
                        const stringIdx = isString(cell) ? strings.indexOf(cell.data.toString()) : cell.data;
                        const cellRef = getRef(rowIdx, columnIdx);
                        const cellStyle = getStyleId(styleMapping, cell);

                        const cellSpan = getCellSpan(cell);
                        const cellType = isString(cell) ? `t="s"` : "";
                        const cells = [`<c r="${cellRef}" s="${cellStyle}" ${cellType}><v>${stringIdx}</v></c>`];
                        columnIdx++;

                        for (let i = 1; i < cellSpan; i++) {
                            const cellRef = getRef(rowIdx, columnIdx);
                            cells.push(`<c r="${cellRef}" s="${cellStyle}" ${cellType} />`);
                            columnIdx++;
                        }

                        if (cellSpan > 1) {
                            const startRef = cellRef;
                            const endRef = getRef(rowIdx, columnIdx - 1);
                            mergedCells.push([startRef, endRef]);
                        }

                        return cells;
                    })
                    .join("\n");

                return `<row r="${rowIdx + 1}" spans="1:${getRowLength(row)}" x14ac:dyDescent="0.25">${cellData}</row>`
            })
            .join("\n");

        const topLeftRef = getRef(0, 0);
        const bottomRightRef = getRef(sheet.rows.length - 1, sheet.rows[sheet.rows.length - 1].cells.length - 1);

        const mergedCellData = mergedCells
            .map(([fromRef, toRef]) => `<mergeCell ref="${fromRef}:${toRef}" />`)
            .join("\n");

        const uid = randomUUID();
        files[filename] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="x14ac xr xr2 xr3"
            xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
            xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
            xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
            xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
            xr:uid="{${uid}}">
            <dimension ref="${topLeftRef}:${bottomRightRef}" />
            <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0">
                    <selection activeCell="${topLeftRef}" sqref="${topLeftRef}" />
                </sheetView>
            </sheetViews>
            <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25" />
            <sheetData>${rowData}</sheetData>
            <mergeCells count="${mergedCells.length}">${mergedCellData}</mergeCells>
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />
        </worksheet>`;
    });

    return files;
}

function isString(cell: Cell): boolean {
    return typeof(cell.type) === 'undefined'|| cell.type === CellType.String;
}

function getRef(rowIdx: number, columnIdx: number): string {
    if (columnIdx > 26) {
        throw "Need to make this function more general for double-lettered column names";
    }

    const letter = String.fromCharCode('A'.charCodeAt(0) + columnIdx);
    const number = rowIdx + 1;
    return `${letter}${number}`
}

function getStyleMapping(): Map<string, string> {
    const styles = new Map<string, string>();
    let index = 0;

    for (let bold of [0, 1]) {
        for (let italic of [0, 1]) {
            for (let underline of [0, 1]) {
                const fontId = bold << 2 | italic << 1 | underline;
                for (let alignHorizontal of [AlignHorizontal.Default, AlignHorizontal.Left, AlignHorizontal.Center, AlignHorizontal.Right]) {
                    styles.set(`${fontId}-${alignHorizontal}`, index.toString());
                    index++;
                }
            }
        }
    }

    return styles;
}

function getStyleId(mapping: Map<string, string>, cell: Cell): string {
    const fontId = (cell.bold ? 1 : 0) << 2 | (cell.italic ? 1 : 0) << 1 | (cell.underline ? 1 : 0);
    const alignHorizontal = cell.alignHorizontal ?? AlignHorizontal.Default;
    const index = mapping.get(`${fontId}-${alignHorizontal}`);
    if (typeof index === 'undefined') throw "Could not find style";
    return index;
}

function getCellSpan(cell: Cell): number {
    if (typeof (cell.span) === 'undefined') {
        return 1;
    }
    return cell.span;
}

function getRowLength(row: Row): number {
    return row.cells.map(getCellSpan).reduce((x, y) => x + y, 0);
}

function randomUUID() {
    if (Object.keys(crypto).includes("randomUUID")) {
        return crypto.randomUUID().toUpperCase();
    }

    return [8, 4, 4, 4, 12].map(getRandomHexChars).join("-");

    function getRandomHexChars(n: number): string {
        const hexChars = "0123456789ABCDEF";
        let s = "";

        while (s.length < n) {
            s += hexChars[Math.floor(Math.random() * hexChars.length)];
        }

        return s;
    }
}

async function createXlsx(spreadsheet: Workbook): Promise<Blob> {
    const files = generateFiles(spreadsheet);

    const zip = new JSZip();

    for (const filename in files) {
        const content = files[filename];
        zip.file(filename, content);
    }

    const data = await zip.generateAsync({ type: "uint8array" });

    return new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
}

function downloadBlob(filename: string, blob: Blob) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.download = filename;
    a.href = url;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
}

export async function downloadXlsx(filename: string, spreadsheet: Workbook) {
    const blob = await createXlsx(spreadsheet);
    downloadBlob(filename, blob);
}
