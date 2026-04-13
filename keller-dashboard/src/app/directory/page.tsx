import { readFileSync } from "fs";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";

export const dynamic = "force-dynamic";

const QR_PATH = "/Users/dancumberland/Documents/Work/Client_Projects/Keller_Associates/Discovery/Quick_Reference.md";

interface TableData {
  headers: string[];
  rows: string[][];
}

function parseMarkdownSections(md: string): { title: string; tables: TableData[] }[] {
  const lines = md.split("\n");
  const sections: { title: string; tables: TableData[] }[] = [];
  let currentSection: { title: string; tables: TableData[] } | null = null;
  let currentTable: TableData | null = null;
  let headerSeparatorSeen = false;

  for (const line of lines) {
    // H2 heading
    if (line.startsWith("## ")) {
      if (currentTable && currentSection) currentSection.tables.push(currentTable);
      if (currentSection) sections.push(currentSection);
      currentSection = { title: line.replace("## ", "").trim(), tables: [] };
      currentTable = null;
      headerSeparatorSeen = false;
      continue;
    }

    // Table row
    if (line.trim().startsWith("|") && currentSection) {
      const cells = line
        .split("|")
        .slice(1, -1)
        .map((c) => c.trim());

      // Separator row (|---|---|)
      if (cells.every((c) => /^[-:]+$/.test(c))) {
        headerSeparatorSeen = true;
        continue;
      }

      if (!currentTable) {
        currentTable = { headers: cells, rows: [] };
      } else if (headerSeparatorSeen) {
        currentTable.rows.push(cells);
      }
      continue;
    }

    // Non-table line after a table — close the table
    if (currentTable && !line.trim().startsWith("|") && line.trim() !== "") {
      if (currentSection) currentSection.tables.push(currentTable);
      currentTable = null;
      headerSeparatorSeen = false;
    }
  }

  // Flush remaining
  if (currentTable && currentSection) currentSection.tables.push(currentTable);
  if (currentSection) sections.push(currentSection);

  return sections;
}

function cleanBold(text: string): string {
  return text.replace(/\*\*/g, "");
}

function isBold(text: string): boolean {
  return text.startsWith("**") && text.endsWith("**");
}

export default function DirectoryPage() {
  const md = readFileSync(QR_PATH, "utf-8");
  const sections = parseMarkdownSections(md);

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <h1 className="text-2xl font-bold">Directory</h1>
        <Badge variant="outline" className="text-xs text-muted-foreground">
          Quick_Reference.md
        </Badge>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {sections.map((section, si) => (
          <Card key={si} className="bg-card border-border">
            <CardHeader className="pb-3">
              <CardTitle className="text-base">{section.title}</CardTitle>
            </CardHeader>
            <CardContent>
              {section.tables.map((table, ti) => (
                <div key={ti} className={`overflow-x-auto ${ti > 0 ? "mt-4" : ""}`}>
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="border-b border-border text-left">
                        {table.headers.map((h, hi) => (
                          <th
                            key={hi}
                            className="pb-2 pr-4 font-medium text-muted-foreground text-xs"
                          >
                            {cleanBold(h)}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {table.rows.map((row, ri) => (
                        <tr
                          key={ri}
                          className="border-b border-border/50 last:border-0"
                        >
                          {row.map((cell, ci) => (
                            <td
                              key={ci}
                              className={`py-1.5 pr-4 ${
                                ci === 0
                                  ? isBold(cell)
                                    ? "font-semibold"
                                    : "font-medium"
                                  : "text-muted-foreground"
                              }`}
                            >
                              {cleanBold(cell)}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ))}
            </CardContent>
          </Card>
        ))}
      </div>
    </div>
  );
}
