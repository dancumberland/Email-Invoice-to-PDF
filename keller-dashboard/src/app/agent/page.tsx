import { getDashboardData } from "@/lib/data";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Mic, MousePointer, Phone, CheckCircle, AlertTriangle, Clock } from "lucide-react";

export const dynamic = "force-dynamic";

export default function AgentPage() {
  const data = getDashboardData();
  const { agent, clicks } = data;

  const uniqueClickers = clicks?.unique_clickers ?? 0;
  const totalConversations = agent?.total_conversations ?? 0;
  const goodConversations = agent?.good_conversations ?? 0;
  const droppedConversations = totalConversations - goodConversations;
  const employees = clicks?.employees_total ?? 226;

  const clickedNotInterviewed = clicks?.clickers.filter((c) => !c.interviewed) ?? [];
  const clickedWithIssue = clicks?.clickers.filter((c) => c.issue) ?? [];

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <h1 className="text-2xl font-bold">AI Interview Agent</h1>
        {agent?.phone_active && (
          <div className="flex items-center gap-2 text-sm text-muted-foreground">
            <Phone className="h-4 w-4 text-green-500" />
            <span>{agent.phone_number}</span>
            <Badge variant="outline" className="text-[10px] border-green-500 text-green-500">Live</Badge>
          </div>
        )}
      </div>

      {/* Funnel */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <FunnelCard
          label="Employees"
          value={employees}
          sublabel="Total invited"
          color="text-muted-foreground"
        />
        <FunnelCard
          label="Clicked Link"
          value={uniqueClickers}
          sublabel={`${Math.round((uniqueClickers / employees) * 100)}% of employees`}
          color="text-blue-400"
        />
        <FunnelCard
          label="Started"
          value={totalConversations}
          sublabel={`${Math.round((totalConversations / employees) * 100)}% of employees`}
          color="text-amber-500"
        />
        <FunnelCard
          label="Completed"
          value={goodConversations}
          sublabel="3+ min conversations"
          color="text-green-500"
        />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Completed Interviews */}
        <Card className="bg-card border-border">
          <CardHeader>
            <CardTitle className="text-lg flex items-center gap-2">
              <CheckCircle className="h-4 w-4 text-green-500" />
              Completed ({goodConversations})
            </CardTitle>
          </CardHeader>
          <CardContent>
            {agent?.completed_people && agent.completed_people.length > 0 ? (
              <div className="space-y-2">
                {agent.completed_people.map((p, i) => (
                  <div key={i} className="flex items-center justify-between bg-muted rounded-lg px-3 py-2">
                    <div>
                      <p className="text-sm font-medium">{p.first_name}</p>
                      <p className="text-xs text-muted-foreground">{p.role}</p>
                    </div>
                    <Badge variant="outline" className="text-[10px] border-green-500 text-green-500">{p.office}</Badge>
                  </div>
                ))}
              </div>
            ) : (
              <p className="text-sm text-muted-foreground">No completions yet</p>
            )}
          </CardContent>
        </Card>

        {/* Issues */}
        <Card className="bg-card border-border">
          <CardHeader>
            <CardTitle className="text-lg flex items-center gap-2">
              <AlertTriangle className="h-4 w-4 text-amber-500" />
              Issues &amp; Follow-ups ({clickedWithIssue.length})
            </CardTitle>
          </CardHeader>
          <CardContent>
            {clickedWithIssue.length > 0 ? (
              <div className="space-y-2">
                {clickedWithIssue.map((c, i) => (
                  <div key={i} className="bg-muted rounded-lg px-3 py-2">
                    <div className="flex items-center justify-between">
                      <p className="text-sm font-medium">{c.first_name}</p>
                      <Badge variant="outline" className="text-[10px]">{c.office}</Badge>
                    </div>
                    <p className="text-xs text-muted-foreground mt-0.5">{c.role}</p>
                    <p className="text-xs text-amber-500 mt-1">{c.issue}</p>
                  </div>
                ))}
              </div>
            ) : (
              <p className="text-sm text-green-500 flex items-center gap-1">
                <CheckCircle className="h-4 w-4" /> No open issues
              </p>
            )}
          </CardContent>
        </Card>
      </div>

      {/* Dropped / Short Conversations */}
      {droppedConversations > 0 && (
        <Card className="bg-card border-border">
          <CardHeader>
            <CardTitle className="text-lg flex items-center gap-2">
              <Clock className="h-4 w-4 text-muted-foreground" />
              Short / Dropped Conversations ({droppedConversations})
            </CardTitle>
          </CardHeader>
          <CardContent>
            {agent?.person_attempts && Object.keys(agent.person_attempts).length > 0 ? (
              <div className="space-y-2">
                {Object.entries(agent.person_attempts)
                  .filter(([key]) => !key.startsWith("Unknown"))
                  .sort((a, b) => b[1] - a[1])
                  .map(([key, count], i) => {
                    const parts = key.split("|");
                    return (
                      <div key={i} className="flex items-center justify-between bg-muted rounded-lg px-3 py-2">
                        <div>
                          <p className="text-sm font-medium">{parts[0]}</p>
                          <p className="text-xs text-muted-foreground">{parts[1] || ""}</p>
                        </div>
                        <div className="flex items-center gap-2">
                          <Badge variant="outline" className="text-[10px]">{parts[2] || ""}</Badge>
                          <span className="text-xs text-muted-foreground">{count}x attempt{count > 1 ? "s" : ""}</span>
                        </div>
                      </div>
                    );
                  })}
              </div>
            ) : (
              <p className="text-sm text-muted-foreground">No attempt data available</p>
            )}
          </CardContent>
        </Card>
      )}

      {/* Clicked but Not Interviewed */}
      <Card className="bg-card border-border">
        <CardHeader>
          <CardTitle className="text-lg flex items-center gap-2">
            <MousePointer className="h-4 w-4 text-blue-400" />
            Clicked but Not Yet Interviewed ({clickedNotInterviewed.length})
          </CardTitle>
        </CardHeader>
        <CardContent>
          {clickedNotInterviewed.length > 0 ? (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-border text-left">
                    <th className="pb-2 font-medium text-muted-foreground">Name</th>
                    <th className="pb-2 font-medium text-muted-foreground">Role</th>
                    <th className="pb-2 font-medium text-muted-foreground">Office</th>
                    <th className="pb-2 font-medium text-muted-foreground">Clicked</th>
                  </tr>
                </thead>
                <tbody>
                  {clickedNotInterviewed.map((c, i) => (
                    <tr key={i} className="border-b border-border/50 last:border-0">
                      <td className="py-2 font-medium">{c.first_name}</td>
                      <td className="py-2 text-muted-foreground text-xs">{c.role}</td>
                      <td className="py-2 text-muted-foreground text-xs">{c.office}</td>
                      <td className="py-2 text-muted-foreground font-mono text-xs">{c.first_click}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="text-sm text-green-500 flex items-center gap-1">
              <CheckCircle className="h-4 w-4" /> All clickers have completed interviews
            </p>
          )}
        </CardContent>
      </Card>
    </div>
  );
}

function FunnelCard({ label, value, sublabel, color }: {
  label: string;
  value: number;
  sublabel: string;
  color: string;
}) {
  return (
    <Card className="bg-card border-border">
      <CardContent className="pt-6">
        <p className="text-sm text-muted-foreground">{label}</p>
        <p className={`text-3xl font-bold mt-1 ${color}`}>{value}</p>
        <p className="text-xs text-muted-foreground mt-1">{sublabel}</p>
      </CardContent>
    </Card>
  );
}
