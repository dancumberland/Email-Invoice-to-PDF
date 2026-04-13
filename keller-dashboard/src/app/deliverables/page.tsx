import { getDashboardData } from "@/lib/data";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { statusColor, statusLabel, DELIVERABLE_STATUS, RESEARCH_PRIORITY } from "@/lib/constants";
import { Check, X, ArrowRight } from "lucide-react";

export const dynamic = "force-dynamic";

export default function DeliverablesPage() {
  const data = getDashboardData();
  const { deliverables, research } = data;

  return (
    <div className="space-y-6">
      <h1 className="text-2xl font-bold">Deliverables</h1>

      {/* Deliverable Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {deliverables.map((del) => {
          const pct = del.inputs_total > 0 ? Math.round((del.inputs_available / del.inputs_total) * 100) : 0;
          return (
            <Card key={del.id} className="bg-card border-border">
              <CardHeader className="pb-3">
                <div className="flex items-start justify-between">
                  <CardTitle className="text-base">{del.title}</CardTitle>
                  <Badge className={`text-[10px] text-white ${statusColor(del.status, DELIVERABLE_STATUS)}`}>
                    {statusLabel(del.status, DELIVERABLE_STATUS)}
                  </Badge>
                </div>
                <p className="text-xs text-muted-foreground">{del.sow_spec}</p>
              </CardHeader>
              <CardContent>
                {/* Input readiness bar */}
                <div className="mb-3">
                  <div className="flex items-center justify-between text-xs mb-1">
                    <span className="text-muted-foreground">Inputs ready</span>
                    <span className="font-mono">{del.inputs_available}/{del.inputs_total} ({pct}%)</span>
                  </div>
                  <div className="w-full bg-muted rounded-full h-1.5">
                    <div
                      className={`h-1.5 rounded-full ${pct === 100 ? "bg-green-500" : pct > 50 ? "bg-blue-500" : "bg-amber-500"}`}
                      style={{ width: `${pct}%` }}
                    />
                  </div>
                </div>

                {/* Input checklist */}
                <div className="space-y-1">
                  {del.inputs_needed.map((input, i) => (
                    <div key={i} className="flex items-center gap-2 text-xs">
                      {input.available ? (
                        <Check className="h-3 w-3 text-green-500 flex-shrink-0" />
                      ) : (
                        <X className="h-3 w-3 text-red-400 flex-shrink-0" />
                      )}
                      <span className={input.available ? "text-muted-foreground" : "text-foreground"}>
                        {input.label}
                      </span>
                      <Badge variant="outline" className="text-[9px] ml-auto">{input.type}</Badge>
                    </div>
                  ))}
                </div>

                {/* Dependencies */}
                {del.depends_on.length > 0 && (
                  <div className="mt-3 pt-3 border-t border-border">
                    <p className="text-[10px] text-muted-foreground flex items-center gap-1">
                      <ArrowRight className="h-3 w-3" />
                      Depends on: {del.depends_on.map((id) => {
                        const dep = deliverables.find((d) => d.id === id);
                        return dep?.title || `#${id}`;
                      }).join(", ")}
                    </p>
                  </div>
                )}

                {del.notes && (
                  <p className="text-xs text-muted-foreground mt-2 italic">{del.notes}</p>
                )}
              </CardContent>
            </Card>
          );
        })}
      </div>

      {/* Research Feed */}
      <Card className="bg-card border-border">
        <CardHeader>
          <CardTitle className="text-lg">Research Portfolio</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-border text-left">
                  <th className="pb-2 font-medium text-muted-foreground">Topic</th>
                  <th className="pb-2 font-medium text-muted-foreground">Priority</th>
                  <th className="pb-2 font-medium text-muted-foreground">Status</th>
                  <th className="pb-2 font-medium text-muted-foreground">Feeds</th>
                </tr>
              </thead>
              <tbody>
                {research.map((item, i) => (
                  <tr key={i} className="border-b border-border/50 last:border-0">
                    <td className="py-2 font-medium">{item.topic}</td>
                    <td className="py-2">
                      <span className={`font-mono text-xs font-bold ${
                        RESEARCH_PRIORITY.find((p) => p.value === item.priority)?.color || ""
                      }`}>
                        {item.priority}
                      </span>
                    </td>
                    <td className="py-2">
                      <Badge variant="outline" className={`text-[10px] ${
                        item.status === "complete" ? "border-green-500 text-green-500" :
                        item.status === "in_progress" ? "border-blue-500 text-blue-500" :
                        "border-gray-500 text-gray-500"
                      }`}>
                        {item.status}
                      </Badge>
                    </td>
                    <td className="py-2 text-xs text-muted-foreground">
                      {item.feeds_deliverable.map((id) => {
                        const del = deliverables.find((d) => d.id === id);
                        return del?.title.split(" ").slice(0, 2).join(" ") || `#${id}`;
                      }).join(", ")}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
