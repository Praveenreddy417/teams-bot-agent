// src/cardBuilder.ts

export type AdaptiveCardJson = {
    $schema?: string;
    type: "AdaptiveCard";
    version: "1.5";
    body: any[];
    actions?: any[];
    [key: string]: any;
};

export const buildChatTextCard = (answer: string): AdaptiveCardJson => ({
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    msteams: {
        width: "Full"
    },
    body: [
        {
            type: "TextBlock",
            text: answer,
            wrap: true,
        },
    ],
});

export const buildChatTableCard = (
    title: string,
    columns: string[],
    rows: string[][]
): AdaptiveCardJson => {

    const headerRow = {
        type: "TableRow",
        style: "accent",
        cells: columns.map((col) => ({
            type: "TableCell",
            items: [
                {
                    type: "TextBlock",
                    text: col,
                    weight: "Bolder",
                    wrap: true,
                    size: "Small"
                }
            ],
        })),
    };

    const dataRows = rows.map((row) => ({
        type: "TableRow",
        cells: row.map((cell) => ({
            type: "TableCell",
            items: [
                {
                    type: "TextBlock",
                    // Prefix lone "-" with a zero-width space so Teams does
                    // not interpret it as a markdown bullet point
                    text: String(cell ?? "").replace(/^-$/, "​-"),
                    wrap: true,
                    size: "Small",
                }
            ],
        })),
    }));

    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",

        msteams: {
            width: "Full"
        },

        body: [
            {
                type: "TextBlock",
                text: `📊 ${title}`,
                weight: "Bolder",
                size: "Large",
                wrap: true,
            },
            {
                type: "TextBlock",
                text: `${rows.length} record(s)`,
                isSubtle: true,
                spacing: "Small",
                wrap: true,
            },
            {
                type: "Table",
                firstRowAsHeaders: true,
                showGridLines: true,
                gridStyle: "accent",

                columns: columns.map(() => ({
                    width: 1
                })),

                rows: [headerRow, ...dataRows],
            }
        ],
    };
};
/**
 * Renders chat "chart" responses using Teams' native Adaptive Card chart elements.
 * The backend (esi_chat_bot) already builds the complete, ready-to-render AdaptiveCard
 * for the chart — including the correct native element type (Chart.Pie /
 * Chart.Donut / Chart.VerticalBar / Chart.HorizontalBar / Chart.VerticalBar.Grouped /
 * Chart.Line / Chart.HorizontalBar.Stacked / Chart.Gauge) and its correctly-shaped
 * data — so this does not reconstruct anything from labels/datasets. It only
 * forwards the card, merging in the msteams full-width hint.
 *
 * `note` is set only when the user asked for a chart type Teams doesn't support
 * (e.g. scatter, bubble, radar, funnel, heatmap...) and the backend substituted
 * the closest supported type instead — it's rendered as a small italic disclaimer
 * above the chart so the user knows why what they see doesn't match what they asked for.
 */
export const buildChatChartCard = (
    _chartType: string | undefined,
    _title: string | undefined,
    card: AdaptiveCardJson,
    note?: string
): AdaptiveCardJson => {
    const body = note
        ? [
              {
                  type: "TextBlock",
                  text: `ℹ️ ${note}`,
                  isSubtle: true,
                  wrap: true,
                  size: "Small",
              },
              ...card.body,
          ]
        : card.body;

    return {
        ...card,
        body,
        msteams: {
            ...(card.msteams ?? {}),
            width: "Full",
        },
    };
};