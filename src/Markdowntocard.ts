/**
 * markdownToCard.ts
 * Converts Boomi agent markdown into Adaptive Card body elements.
 */

export function markdownToCardBody(markdown: string): any[] {
    const lines = markdown.replace(/\r\n/g, "\n").split("\n");
    return parseLines(lines, 0).elements;
}

// ── Embedded ```chart spec → native Adaptive Card chart element ───────────────
// The agent's ```chart fences carry Chart.js-shaped JSON (data.labels /
// data.datasets[].data|label / options.title.text), not a pre-built card, so
// this converts it the same way the old standalone chart-response path did.

const toNum = (v: unknown): number => {
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
};

function buildChartElementFromSpec(spec: any): any {
    const rawType = String(spec?.type ?? "bar").toLowerCase();
    const labels: string[] = Array.isArray(spec?.data?.labels) ? spec.data.labels.map(String) : [];
    const datasets: Array<{ label?: string; data?: unknown[] }> =
        Array.isArray(spec?.data?.datasets) && spec.data.datasets.length ? spec.data.datasets : [{ data: [] }];
    const isMultiSeries = datasets.length > 1;
    const stacked = Boolean(
        spec?.options?.scales?.xAxes?.[0]?.stacked || spec?.options?.scales?.yAxes?.[0]?.stacked
    );
    const chartTitle: string | undefined = spec?.options?.title?.text;

    const type =
        rawType.includes("pie") ? "pie" :
        rawType.includes("doughnut") || rawType.includes("donut") ? "donut" :
        rawType.includes("horizontal") ? "horizontalbar" :
        rawType.includes("line") ? "line" :
        "bar";

    const summaryRows = labels.map(
        (label, i) => `${label}: ${datasets.map((ds) => String(ds.data?.[i] ?? "-")).join(" | ")}`
    );
    const fallback = {
        type: "TextBlock",
        text: summaryRows.length ? summaryRows.join("\n") : "No data to display.",
        wrap: true,
        size: "Small",
    };

    let chartElement: any;

    if (type === "pie" || type === "donut") {
        const series = datasets[0];
        chartElement = {
            type: type === "pie" ? "Chart.Pie" : "Chart.Donut",
            colorSet: "categorical",
            data: labels.map((label, i) => ({ legend: label, value: toNum(series.data?.[i]) })),
            fallback,
        };
    } else if (type === "line" || (type === "bar" && isMultiSeries)) {
        chartElement = {
            type: type === "line" ? "Chart.Line" : "Chart.VerticalBar.Grouped",
            yAxisTitle: datasets[0].label ?? "Value",
            colorSet: "categorical",
            ...(type !== "line" ? { showBarValues: true, stacked } : {}),
            data: datasets.map((ds) => ({
                legend: ds.label ?? "Series",
                values: labels.map((label, i) => ({ x: label, y: toNum(ds.data?.[i]) })),
            })),
            fallback,
        };
    } else if (type === "horizontalbar") {
        const series = datasets[0];
        chartElement = {
            type: "Chart.HorizontalBar",
            yAxisTitle: series.label ?? "Value",
            colorSet: "categorical",
            data: labels.map((label, i) => ({ x: label, y: toNum(series.data?.[i]) })),
            fallback,
        };
    } else {
        const series = datasets[0];
        chartElement = {
            type: "Chart.VerticalBar",
            yAxisTitle: series.label ?? "Value",
            colorSet: "categorical",
            showBarValues: true,
            data: labels.map((label, i) => ({ x: label, y: toNum(series.data?.[i]) })),
            fallback,
        };
    }

    if (!chartTitle) return chartElement;

    return {
        type: "Container",
        spacing: "Small",
        items: [
            { type: "TextBlock", text: chartTitle, weight: "Bolder", size: "Small", wrap: true, spacing: "None" },
            chartElement,
        ],
    };
}

// The agent doesn't always emit valid JSON in a ```chart fence — it sometimes emits a
// YAML-ish shorthand instead (bare `key: value` lines, `- ` prefixed dataset entries,
// inline `[...]` arrays). This is not general YAML — just enough of that shape to cover
// what's actually been observed — parsed without indentation being significant, since
// the agent's own indentation here isn't reliable.
function parseValueLoosely(raw: string): any {
    const v = raw.trim();
    if (v.startsWith("[")) {
        try { return JSON.parse(v); } catch { /* fall through */ }
    }
    if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) {
        return v.slice(1, -1);
    }
    if (/^-?\d+(\.\d+)?$/.test(v)) return Number(v);
    return v;
}

function parseChartSpecLoosely(specLines: string[]): any {
    const spec: any = { type: undefined, options: { title: {} }, data: { labels: [], datasets: [] } };
    let currentDataset: Record<string, any> | null = null;

    for (const raw of specLines) {
        const line = raw.trim();
        if (!line) continue;

        // "- key: value" (or bare "-") starts a new dataset entry.
        const datasetStart = line.match(/^-+\s*(.*)$/);
        if (datasetStart) {
            if (currentDataset) spec.data.datasets.push(currentDataset);
            currentDataset = {};
            const rest = datasetStart[1];
            const kv = rest.match(/^(\w+):\s*(.*)$/);
            if (kv) currentDataset[kv[1].toLowerCase()] = parseValueLoosely(kv[2]);
            continue;
        }

        const m = line.match(/^(\w+):\s*(.*)$/);
        if (!m) continue;
        const key = m[1].toLowerCase();
        const valueRaw = m[2];

        if (key === "type") { spec.type = parseValueLoosely(valueRaw); continue; }
        if (key === "title") { spec.options.title.text = parseValueLoosely(valueRaw); continue; }
        if (key === "data" && !valueRaw) continue;   // "data:" section header
        if (key === "labels") { spec.data.labels = parseValueLoosely(valueRaw); continue; }
        if (key === "datasets" && !valueRaw) continue; // "datasets:" section header

        // Anything else (label / data / backgroundColor) belongs to the dataset in progress.
        if (currentDataset) currentDataset[key] = parseValueLoosely(valueRaw);
    }
    if (currentDataset) spec.data.datasets.push(currentDataset);

    return spec;
}

const KNOWN_FENCE_LANGUAGE_NAMES = new Set([
    "text", "plaintext", "plain", "json", "sql", "python", "javascript", "js",
    "typescript", "ts", "bash", "shell", "sh", "yaml", "yml", "xml", "html", "css",
]);

function monospaceBlock(text: string): any {
    return { type: "TextBlock", text, wrap: true, fontType: "Monospace", size: "Small" };
}

// The agent sometimes draws a bar chart out of block characters instead of using a
// ```chart fence, e.g.:
//   POSSIBLE_DUPLICATE        ████████████████████████████ 28
//   Data Validation Error     ████████ 8
// Detects that shape and converts it into a real Chart.HorizontalBar. Axis/ruler
// lines (no bar characters) don't match and are simply dropped — the native chart
// draws its own axis. Returns null (caller falls back to a monospace block) unless
// at least two bars are found, to avoid misfiring on ordinary code/log output.
function tryParseAsciiBarChart(blockLines: string[]): any | null {
    const barLineRegex = /^(.{1,60}?)\s{2,}[█▓▒░#*=]+\s+(-?\d+(?:\.\d+)?)\s*$/;
    const bars: { label: string; value: number }[] = [];
    for (const raw of blockLines) {
        const m = raw.match(barLineRegex);
        if (m) bars.push({ label: m[1].trim(), value: Number(m[2]) });
    }
    if (bars.length < 2) return null;

    return {
        type: "Chart.HorizontalBar",
        colorSet: "categorical",
        data: bars.map((b) => ({ x: b.label, y: b.value })),
        fallback: {
            type: "TextBlock",
            text: bars.map((b) => `${b.label}: ${b.value}`).join("\n"),
            wrap: true,
            size: "Small",
        },
    };
}

const STAT_LABEL_REGEX = /^[A-Za-z][A-Za-z0-9 /&()-]{2,50}$/;
const STAT_VALUE_REGEX = /^-?\d[\d,]*(\.\d+)?%?$/;

function buildStatRow(stats: { label: string; value: string }[]): any {
    return {
        type: "ColumnSet",
        spacing: "Medium",
        columns: stats.map((s) => ({
            type: "Column",
            width: "stretch",
            items: [
                { type: "TextBlock", text: s.value, weight: "Bolder", size: "ExtraLarge", horizontalAlignment: "Center", wrap: true, spacing: "None" },
                { type: "TextBlock", text: s.label, size: "Small", isSubtle: true, horizontalAlignment: "Center", wrap: true, spacing: "None" },
            ],
        })),
    };
}

// ── Core line parser ──────────────────────────────────────────────────────────

function parseLines(
    lines: string[],
    startIdx: number
): { elements: any[]; nextIdx: number } {
    const elements: any[] = [];
    let i = startIdx;

    // Matches "- **Key:** value" or "**Key:** value" or "Key: value"
    // but NOT "1. **Text** - description" (numbered list items)
    const kvRegex = /^[-*•\s]*\*{0,2}([^:*\n]{2,40})\*{0,2}:\s+([^-].+)$/;

    while (i < lines.length) {
        const line = lines[i];
        const trimmed = line.trim();

        if (!trimmed) { i++; continue; }

        // ── Fenced code block ─────────────────────────────────────────────
        // ```chart fences carry a chart spec (JSON or YAML-ish shorthand) → native
        // Adaptive Card chart element. Any other fence — including unlabeled ones
        // containing an ASCII bar chart (block characters + trailing numbers, a
        // pattern this agent also emits) — is fully consumed here too, so its
        // ``` markers never leak into the chat as literal text; ASCII bar charts
        // get converted to a real Chart.HorizontalBar, everything else renders
        // as a monospace block that at least preserves alignment.
        const fenceMatch = trimmed.match(/^```(\w*)\s*$/);
        if (fenceMatch) {
            const lang = fenceMatch[1].toLowerCase();
            i++;
            const blockLines: string[] = [];
            while (i < lines.length && !/^```\s*$/.test(lines[i].trim())) {
                blockLines.push(lines[i]);
                i++;
            }
            i++; // skip closing ```

            if (lang === "chart") {
                try {
                    const spec = JSON.parse(blockLines.join("\n"));
                    elements.push(buildChartElementFromSpec(spec));
                } catch {
                    try {
                        // Not valid JSON — the agent sometimes emits a YAML-ish shorthand instead.
                        const spec = parseChartSpecLoosely(blockLines);
                        elements.push(buildChartElementFromSpec(spec));
                    } catch {
                        // Still unparseable — show it as plain text rather than dropping it silently.
                        elements.push(monospaceBlock(blockLines.join("\n")));
                    }
                }
                continue;
            }

            const asciiChart = tryParseAsciiBarChart(blockLines);
            if (asciiChart) {
                elements.push(asciiChart);
                continue;
            }

            const content = blockLines.join("\n").trim();
            // A fenced block whose entire content is just a bare language name (e.g. a
            // stray ```\nText\n``` with nothing else) is a mis-split language tag, not
            // real content — drop it rather than showing a meaningless standalone line.
            const isBareLanguageTag = KNOWN_FENCE_LANGUAGE_NAMES.has(content.toLowerCase());
            if (content && !isBareLanguageTag) elements.push(monospaceBlock(content));
            continue;
        }

        // ── H1 ──────────────────────────────────────────────────────────
        if (/^# (?!#)/.test(line)) {
            elements.push({
                type: "TextBlock",
                text: line.replace(/^# /, "").replace(/\*\*/g, "").trim(),
                weight: "Bolder",
                size: "ExtraLarge",
                wrap: true,
                spacing: "None",
            });
            i++; continue;
        }

        // ── H2 → section Container ───────────────────────────────────────
        if (/^## (?!#)/.test(line)) {
            const title = line.replace(/^##\s*/, "").replace(/\*\*/g, "").trim();
            i++;

            const sectionLines: string[] = [];
            while (i < lines.length && !/^#{1,2} /.test(lines[i])) {
                sectionLines.push(lines[i]);
                i++;
            }

            const inner = parseLines(sectionLines, 0).elements;
            elements.push({
                type: "Container",
                style: "emphasis",
                spacing: "Medium",
                items: [
                    {
                        type: "TextBlock",
                        text: title,
                        weight: "Bolder",
                        size: "Medium",
                        color: "Accent",
                        wrap: true,
                        spacing: "None",
                    },
                    ...inner,
                ],
            });
            continue;
        }

        // ── H3 ──────────────────────────────────────────────────────────
        if (/^### /.test(line)) {
            const title = line.replace(/^###\s*/, "").replace(/\*\*/g, "").trim();
            const isInsights = /insight|summary|note|highlight/i.test(title);

            if (isInsights) {
                i++;
                const insightItems: string[] = [];
                while (i < lines.length && /^[-*•\s]*[-*•]\s/.test(lines[i]) && lines[i].trim()) {
                    insightItems.push("• " + lines[i].replace(/^[\s]*[-*•]\s+/, "").trim());
                    i++;
                }
                elements.push(buildInsightsContainer(title, insightItems));
                continue;
            }

            elements.push({
                type: "TextBlock",
                text: title,
                weight: "Bolder",
                size: "Small",
                wrap: true,
                spacing: "Medium",
                isSubtle: true,
            });
            i++; continue;
        }

        // ── Horizontal rule ──────────────────────────────────────────────
        if (/^---+$/.test(trimmed)) {
            elements.push({ type: "TextBlock", text: " ", separator: true, spacing: "Small" });
            i++; continue;
        }

        // ── Numeric stat pairs ("Label" line immediately followed by a bare number) ──
        // e.g. "Active Records" / "21" / "Resolved Records" / "79" — rendered as
        // side-by-side big-number stat tiles instead of a run-on paragraph. Scoped
        // tightly (label has no punctuation, value line is nothing but a number) to
        // avoid misfiring on ordinary two-line prose.
        if (
            STAT_LABEL_REGEX.test(trimmed) &&
            i + 1 < lines.length &&
            STAT_VALUE_REGEX.test(lines[i + 1].trim())
        ) {
            const stats: { label: string; value: string }[] = [];
            while (
                i + 1 < lines.length &&
                STAT_LABEL_REGEX.test(lines[i].trim()) &&
                STAT_VALUE_REGEX.test(lines[i + 1].trim())
            ) {
                stats.push({ label: lines[i].trim(), value: lines[i + 1].trim() });
                i += 2;
            }
            elements.push(buildStatRow(stats));
            continue;
        }

        // ── Markdown table ───────────────────────────────────────────────
        if (line.startsWith("|")) {
            const tableLines: string[] = [];
            while (i < lines.length && lines[i].startsWith("|")) {
                if (!/^\|[-: |]+\|$/.test(lines[i].trim())) tableLines.push(lines[i]);
                i++;
            }
            if (tableLines.length > 0) elements.push(buildTable(tableLines));
            continue;
        }

        // ── Numbered list (with optional nested sub-bullets) ─────────────
        // Handles both "1. Text" (text on the same line) and a bare "1." with its
        // text on the following non-blank line(s) — this agent emits both shapes.
        if (/^\d+\.\s/.test(trimmed) || /^\d+\.\s*$/.test(trimmed)) {
            const items: { num: number; text: string; subitems: string[] }[] = [];
            let counter = 1;

            while (i < lines.length) {
                const cur = lines[i].trim();

                // Skip blank lines between list items without breaking the list
                if (!cur) { i++; continue; }

                // Stop if we hit a heading, table, or non-list content
                if (/^#+\s/.test(lines[i]) || lines[i].startsWith("|")) break;

                // Indented sub-bullet — belongs to previous item, handled below
                if (/^[\s]{2,}[-*•]\s/.test(lines[i])) {
                    if (items.length > 0) {
                        items[items.length - 1].subitems.push(
                            lines[i].replace(/^[\s]*[-*•]\s+/, "").trim()
                        );
                    }
                    i++; continue;
                }

                // Numbered item, text on the same line
                const inlineMatch = cur.match(/^\d+\.\s+(.+)$/);
                if (inlineMatch) {
                    items.push({ num: counter++, text: inlineMatch[1], subitems: [] });
                    i++; continue;
                }

                // Bare "N." — its text is on the following non-blank line(s)
                if (/^\d+\.\s*$/.test(cur)) {
                    i++;
                    while (i < lines.length && !lines[i].trim()) i++; // skip blank lines after "N."
                    const textLines: string[] = [];
                    while (
                        i < lines.length &&
                        lines[i].trim() &&
                        !/^\d+\.\s*$/.test(lines[i].trim()) &&
                        !/^\d+\.\s+/.test(lines[i].trim()) &&
                        !/^#+\s/.test(lines[i]) &&
                        !lines[i].startsWith("|") &&
                        !/^```/.test(lines[i].trim()) &&
                        !(STAT_LABEL_REGEX.test(lines[i].trim()) && i + 1 < lines.length && STAT_VALUE_REGEX.test(lines[i + 1].trim()))
                    ) {
                        textLines.push(lines[i].trim());
                        i++;
                    }
                    items.push({ num: counter++, text: textLines.join(" "), subitems: [] });
                    continue;
                }

                // Anything else ends the list
                break;
            }

            for (const item of items) {
                elements.push({
                    type: "ColumnSet",
                    spacing: "Small",
                    columns: [
                        {
                            type: "Column",
                            width: "24px",
                            items: [{
                                type: "TextBlock",
                                text: `${item.num}.`,
                                weight: "Bolder",
                                size: "Small",
                                spacing: "None",
                                wrap: false,
                            }],
                        },
                        {
                            type: "Column",
                            width: "stretch",
                            items: [
                                {
                                    type: "TextBlock",
                                    text: item.text,
                                    wrap: true,
                                    size: "Small",
                                    spacing: "None",
                                },
                                ...(item.subitems.length > 0
                                    ? [{
                                        type: "TextBlock",
                                        text: item.subitems.map((s) => `• ${s}`).join("\n"),
                                        wrap: true,
                                        size: "Small",
                                        isSubtle: true,
                                        spacing: "Small",
                                    }]
                                    : []),
                            ],
                        },
                    ],
                });
            }
            continue;
        }

        // ── Key-value bullets → ColumnSet rows ───────────────────────────
        if (kvRegex.test(trimmed)) {
            const facts: { title: string; value: string }[] = [];
            while (i < lines.length && kvRegex.test(lines[i].trim())) {
                const m = lines[i].trim().match(kvRegex);
                if (m) facts.push({ title: m[1].trim(), value: m[2].trim() });
                i++;
            }
            if (facts.length > 0) elements.push(buildFactRows(facts));
            continue;
        }

        // ── Plain bullet list ────────────────────────────────────────────
        // Strip only the leading bullet marker — don't double-prefix
        if (/^[-*•]\s/.test(trimmed) || /^[\s]{0,3}[-*•]\s/.test(line)) {
            const items: string[] = [];
            while (
                i < lines.length &&
                lines[i].trim() &&
                (/^[-*•]\s/.test(lines[i].trim()) || /^[\s]{0,3}[-*•]\s/.test(lines[i])) &&
                !kvRegex.test(lines[i].trim())
            ) {
                // Remove leading spaces + any single bullet character + space
                const clean = lines[i].replace(/^[\s]*[-*•]\s+/, "").trim();
                items.push(`• ${clean}`);
                i++;
            }
            if (items.length > 0) {
                elements.push({
                    type: "TextBlock",
                    text: items.join("\n"),
                    wrap: true,
                    spacing: "Small",
                });
            }
            continue;
        }

        // ── Regular paragraph ─────────────────────────────────────────────
        // Stops before anything a more specific branch above would otherwise handle
        // on its next pass — a fence start, a numbered item (inline or bare "N."), or
        // a stat-pair label — so those don't get swallowed as plain text mid-paragraph.
        const para: string[] = [];
        while (
            i < lines.length &&
            lines[i].trim() &&
            !/^#+\s/.test(lines[i]) &&
            !lines[i].startsWith("|") &&
            !/^[-*•]\s/.test(lines[i].trim()) &&
            !/^[\s]{0,3}[-*•]\s/.test(lines[i]) &&
            !/^\d+\.\s*$/.test(lines[i].trim()) &&
            !/^\d+\.\s/.test(lines[i].trim()) &&
            !/^```/.test(lines[i].trim()) &&
            !kvRegex.test(lines[i].trim()) &&
            !(STAT_LABEL_REGEX.test(lines[i].trim()) && i + 1 < lines.length && STAT_VALUE_REGEX.test(lines[i + 1].trim()))
        ) {
            para.push(lines[i]);
            i++;
        }
        if (para.length > 0) {
            elements.push({
                type: "TextBlock",
                text: para.join("\n"),
                wrap: true,
                spacing: "Small",
            });
        }
    }

    return { elements, nextIdx: i };
}

// ── Key-value ColumnSet rows ──────────────────────────────────────────────────

function buildFactRows(facts: { title: string; value: string }[]): any {
    return {
        type: "Container",
        spacing: "Small",
        items: facts.map((f) => ({
            type: "ColumnSet",
            spacing: "Small",
            columns: [
                {
                    type: "Column",
                    width: "160px",
                    items: [{
                        type: "TextBlock",
                        text: f.title,
                        weight: "Bolder",
                        size: "Small",
                        wrap: true,
                        spacing: "None",
                    }],
                },
                {
                    type: "Column",
                    width: "stretch",
                    items: [{
                        type: "TextBlock",
                        text: f.value,
                        size: "Small",
                        wrap: true,
                        spacing: "None",
                    }],
                },
            ],
        })),
    };
}

// ── Key Insights container ────────────────────────────────────────────────────

function buildInsightsContainer(title: string, items: string[]): any {
    return {
        type: "Container",
        style: "good",
        spacing: "Medium",
        items: [
            {
                type: "TextBlock",
                text: `💡 ${title}`,
                weight: "Bolder",
                size: "Small",
                wrap: true,
                spacing: "None",
            },
            ...(items.length > 0
                ? [{
                    type: "TextBlock",
                    text: items.join("\n"),
                    wrap: true,
                    spacing: "Small",
                    size: "Small",
                }]
                : []),
        ],
    };
}

// ── Table builder ─────────────────────────────────────────────────────────────

function parseTableRow(row: string): string[] {
    return row
        .split("|")
        .filter((_, idx, arr) => idx > 0 && idx < arr.length - 1)
        .map((c) => c.trim());
}

function buildTable(tableLines: string[]): any {
    if (tableLines.length === 0) return { type: "TextBlock", text: "" };

    const headers = parseTableRow(tableLines[0]);
    const dataRows = tableLines.slice(1).map(parseTableRow);
    const colCount = Math.max(headers.length, ...dataRows.map((r) => r.length));

    const headerRow = {
        type: "TableRow",
        style: "accent",
        cells: headers.map((h) => ({
            type: "TableCell",
            items: [{
                type: "TextBlock",
                text: h || " ",
                weight: "Bolder",
                wrap: true,
                size: "Small",
            }],
        })),
    };

    const bodyRows = dataRows.map((row) => ({
        type: "TableRow",
        cells: Array.from({ length: colCount }, (_, c) => ({
            type: "TableCell",
            items: [{
                type: "TextBlock",
                text: String(row[c] ?? "").replace(/^-$/, "\u200B-"),
                wrap: true,
                size: "Small",
            }],
        })),
    }));

    return {
        type: "Table",
        gridStyle: "accent",
        firstRowAsHeaders: true,
        showGridLines: true,
        spacing: "Small",
        columns: Array.from({ length: colCount }, () => ({ width: 1 })),
        rows: [headerRow, ...bodyRows],
    };
}