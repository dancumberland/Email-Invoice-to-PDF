import { getDashboardData } from "@/lib/data";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { daysUntil, daysSince, formatCurrency } from "@/lib/constants";
import { Clock, AlertTriangle, CheckCircle, Calendar, Mic, MousePointer } from "lucide-react";

export const dynamic = "force-dynamic";

export default function CockpitPage() {
  const data = getDashboardData();
  const { engagement, interviews, pipeline, actions, comms, invoices, agent, clicks } = data;

  const lastComm = comms.length > 0 ? comms[0] : null;
  const lastCommDays = lastComm ? daysSince(lastComm.date) : null;

  return (
    <div className="space-y-6">
      <h1 className="text-2xl font-bold">Command Center</h1>

      {/* Countdown Strip */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <MetricCard
          title="Days to May 7"
          value={String(daysUntil(engagement.end_date))}
          subtitle="Shareholder Meeting"
          urgent={daysUntil(engagement.end_date) <= 14}
        />
        <MetricCard
          title="Days to Discovery Close"
          value={String(daysUntil("2026-04-22"))}
          subtitle="April 22"
        />
        <MetricCard
          title="Current Week"
          value={`${engagement.current_week} of 7`}
          subtitle={engagement.phase.charAt(0).toUpperCase() + engagement.phase.slice(1)}
        />
        <MetricCard
          title="Interviews Done"
          value={`${data.people.filter((p) => p.interview_status !== "scheduled").length}/${data.people.length}`}
          subtitle="People interviewed"
        />
      </div>

      {/* AI Agent Status Strip */}
      {(agent || clicks) && (
        <Card className="bg-card border-border">
          <CardHeader className="pb-3">
            <CardTitle className="text-lg flex items-center gap-2">
              <Mic className="h-4 w-4 text-blue-400" />
              AI Interview Agent
              {agent?.phone_active && (
                <Badge variant="outline" className="text-[10px] border-green-500 text-green-500 ml-1">📞 Phone active</Badge>
              )}
            </CardTitle>
          </CardHeader>
          <CardContent>
            <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
              <div className="text-center">
                <p className="text-2xl font-bold">{agent?.total_conversations ?? "—"}</p>
                <p className="text-xs text-muted-foreground">Conversations</p>
              </div>
              <div className="text-center">
                <p className="text-2xl font-bold text-green-500">{agent?.good_conversations ?? "—"}</p>
                <p className="text-xs text-muted-foreground">Completed (3+ min)</p>
              </div>
              <div className="text-center">
                <p className="text-2xl font-bold text-amber-500">{(agent?.total_conversations ?? 0) - (agent?.good_conversations ?? 0)}</p>
                <p className="text-xs text-muted-foreground">Short / Dropped</p>
              </div>
              <div className="text-center">
                <p className="text-2xl font-bold text-blue-400 flex items-center justify-center gap-1">
                  <MousePointer className="h-4 w-4" />{clicks?.unique_clickers ?? "—"}
                </p>
                <p className="text-xs text-muted-foreground">Link clicks</p>
              </div>
              <div className="text-center">
                <p className="text-2xl font-bold text-muted-foreground">
                  {clicks ? `${Math.round((clicks.unique_clickers / clicks.employees_total) * 100)}%` : "—"}
                </p>
                <p className="text-xs text-muted-foreground">of 226 employees</p>
              </div>
            </div>
            {agent?.phone_active && (
              <p className="text-xs text-muted-foreground mt-3 text-center">
                Phone: {agent.phone_number} · Web + phone available
              </p>
            )}
          </CardContent>
        </Card>
      )}

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Left Column */}
        <div className="space-y-6">
          {/* Today's Schedule */}
          <Card className="bg-card border-border">
            <CardHeader>
              <CardTitle className="text-lg flex items-center gap-2">
                <Calendar className="h-4 w-4" />
                Upcoming Interviews
              </CardTitle>
            </CardHeader>
            <CardContent>
              {interviews.length > 0 ? (
                <div className="space-y-3">
                  {interviews.map((interview, i) => (
                    <div key={i} className="flex items-center justify-between bg-muted rounded-lg px-3 py-2">
                      <div>
                        <p className="text-sm font-medium">{interview.person}</p>
                        <p className="text-xs text-muted-foreground">{interview.role}</p>
                      </div>
                      <div className="text-right">
                        <p className="text-sm font-mono">{interview.time_mt} MT</p>
                        <p className="text-xs text-muted-foreground">{interview.date}</p>
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-sm text-muted-foreground">No upcoming interviews</p>
              )}
            </CardContent>
          </Card>

          {/* Pending Extractions */}
          <Card className="bg-card border-border">
            <CardHeader>
              <CardTitle className="text-lg flex items-center gap-2">
                <AlertTriangle className="h-4 w-4 text-amber-500" />
                Pending Extractions
              </CardTitle>
            </CardHeader>
            <CardContent>
              {pipeline.length > 0 ? (
                <div className="space-y-2">
                  {pipeline.map((item, i) => (
                    <div key={i} className="flex items-center justify-between bg-muted rounded-lg px-3 py-2">
                      <div>
                        <p className="text-sm font-medium">{item.person}</p>
                        <p className="text-xs text-muted-foreground">{item.transcript_date}</p>
                      </div>
                      <Badge variant="outline" className={`text-xs ${item.stage === "transcript" ? "border-amber-500 text-amber-500" : "border-purple-500 text-purple-500"}`}>
                        {item.stage}
                      </Badge>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-sm text-green-500 flex items-center gap-1">
                  <CheckCircle className="h-4 w-4" /> All caught up
                </p>
              )}
            </CardContent>
          </Card>
        </div>

        {/* Right Column */}
        <div className="space-y-6">
          {/* Action Items */}
          <Card className="bg-card border-border">
            <CardHeader>
              <CardTitle className="text-lg">Action Items</CardTitle>
            </CardHeader>
            <CardContent>
              {actions.length > 0 ? (
                <div className="space-y-2">
                  {actions.map((action, i) => (
                    <div key={i} className="flex items-start gap-2 bg-muted rounded-lg px-3 py-2">
                      <div className={`mt-1 w-2 h-2 rounded-full flex-shrink-0 ${
                        action.priority === "high" ? "bg-red-500" :
                        action.priority === "medium" ? "bg-amber-500" : "bg-gray-500"
                      }`} />
                      <div>
                        <p className="text-sm">{action.description}</p>
                        {action.waiting_on && (
                          <p className="text-xs text-muted-foreground mt-0.5">Waiting on: {action.waiting_on}</p>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-sm text-muted-foreground">No action items</p>
              )}
            </CardContent>
          </Card>

          {/* Last Client Touchpoint */}
          {lastComm && (
            <Card className="bg-card border-border">
              <CardHeader>
                <CardTitle className="text-lg">Last Client Contact</CardTitle>
              </CardHeader>
              <CardContent>
                <div className="bg-muted rounded-lg px-3 py-2">
                  <div className="flex items-center justify-between">
                    <p className="text-sm font-medium">{lastComm.with_person}</p>
                    <span className={`text-xs ${lastCommDays !== null && lastCommDays > 5 ? "text-amber-500" : "text-muted-foreground"}`}>
                      {lastCommDays === 0 ? "Today" : lastCommDays === 1 ? "Yesterday" : `${lastCommDays}d ago`}
                    </span>
                  </div>
                  <p className="text-xs text-muted-foreground mt-1">{lastComm.summary}</p>
                  <Badge variant="outline" className="text-[10px] mt-1">{lastComm.type}</Badge>
                </div>
              </CardContent>
            </Card>
          )}

          {/* Invoice Status */}
          <Card className="bg-card border-border">
            <CardHeader>
              <CardTitle className="text-lg">Invoice Status</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="space-y-2">
                {invoices.map((inv, i) => (
                  <div key={i} className="flex items-center justify-between text-sm bg-muted rounded-lg px-3 py-2">
                    <div>
                      <span className="font-medium">M{inv.milestone}</span>
                      <span className="text-muted-foreground ml-2">{inv.description}</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <span className="font-mono">{formatCurrency(inv.amount)}</span>
                      <Badge variant="outline" className={`text-[10px] ${
                        inv.status === "paid" ? "border-green-500 text-green-500" :
                        inv.status === "sent" ? "border-blue-500 text-blue-500" :
                        "border-gray-500 text-gray-500"
                      }`}>
                        {inv.status}
                      </Badge>
                    </div>
                  </div>
                ))}
              </div>
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  );
}

function MetricCard({ title, value, subtitle, urgent }: { title: string; value: string; subtitle?: string; urgent?: boolean }) {
  return (
    <Card className="bg-card border-border">
      <CardContent className="pt-6">
        <p className="text-sm text-muted-foreground">{title}</p>
        <p className={`text-3xl font-bold mt-1 ${urgent ? "text-red-500" : "text-foreground"}`}>{value}</p>
        {subtitle && <p className="text-xs text-muted-foreground mt-1">{subtitle}</p>}
      </CardContent>
    </Card>
  );
}
