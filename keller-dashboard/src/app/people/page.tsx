import { getDashboardData } from "@/lib/data";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { statusColor, statusLabel, INTERVIEW_STATUS, READINESS_TIERS } from "@/lib/constants";

export const dynamic = "force-dynamic";

export default function PeoplePage() {
  const data = getDashboardData();
  const { people, coverage, readiness, waves } = data;

  const completed = people.filter((p) => p.interview_status !== "scheduled").length;
  const extracted = people.filter((p) => ["extracted", "routed"].includes(p.interview_status)).length;
  const routed = people.filter((p) => p.interview_status === "routed").length;

  return (
    <div className="space-y-6">
      <h1 className="text-2xl font-bold">People & Pipeline</h1>

      {/* Stats Strip */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <StatCard label="Total Scheduled" value={people.length} />
        <StatCard label="Interviews Done" value={completed} />
        <StatCard label="Briefs Extracted" value={extracted} />
        <StatCard label="Routed to Analysis" value={routed} />
      </div>

      {/* Pipeline Table */}
      <Card className="bg-card border-border">
        <CardHeader>
          <CardTitle className="text-lg">Interview Pipeline</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-border text-left">
                  <th className="pb-2 font-medium text-muted-foreground">Name</th>
                  <th className="pb-2 font-medium text-muted-foreground">Role</th>
                  <th className="pb-2 font-medium text-muted-foreground">Office</th>
                  <th className="pb-2 font-medium text-muted-foreground">Discipline</th>
                  <th className="pb-2 font-medium text-muted-foreground">Wave</th>
                  <th className="pb-2 font-medium text-muted-foreground">Date</th>
                  <th className="pb-2 font-medium text-muted-foreground">Status</th>
                </tr>
              </thead>
              <tbody>
                {people.map((person, i) => (
                  <tr key={i} className="border-b border-border/50 last:border-0">
                    <td className="py-2 font-medium">{person.name}</td>
                    <td className="py-2 text-muted-foreground">{person.role}</td>
                    <td className="py-2 text-muted-foreground">{person.office}</td>
                    <td className="py-2 text-muted-foreground">{person.discipline}</td>
                    <td className="py-2">
                      <Badge variant="outline" className="text-[10px]">W{person.wave}</Badge>
                    </td>
                    <td className="py-2 text-muted-foreground font-mono text-xs">{person.interview_date || "TBD"}</td>
                    <td className="py-2">
                      <Badge className={`text-[10px] text-white ${statusColor(person.interview_status, INTERVIEW_STATUS)}`}>
                        {statusLabel(person.interview_status, INTERVIEW_STATUS)}
                      </Badge>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </CardContent>
      </Card>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Coverage Matrix */}
        <Card className="bg-card border-border">
          <CardHeader>
            <CardTitle className="text-lg">Coverage Matrix</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="border-b border-border">
                    <th className="pb-2 text-left font-medium text-muted-foreground">Office</th>
                    {coverage.disciplines.map((d) => {
                      const abbrevs: Record<string, string> = {
                        "Water/Wastewater": "W/WW",
                        "Transportation": "Trans",
                        "Structural": "Struct",
                        "Electrical": "Elec",
                        "Civil/Site": "Civil",
                        "Construction Mgmt": "Const",
                        "Engineering": "Eng",
                        "Accounting": "Acct",
                        "Marketing": "Mktg",
                        "Leadership": "Lead",
                      };
                      return (
                        <th key={d} className="pb-2 font-medium text-muted-foreground text-center px-1" title={d}>
                          {abbrevs[d] || d.substring(0, 5)}
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody>
                  {coverage.offices.map((office) => (
                    <tr key={office} className="border-b border-border/50 last:border-0">
                      <td className="py-1.5 font-medium whitespace-nowrap">{office}</td>
                      {coverage.disciplines.map((disc) => {
                        const names = coverage.matrix[office]?.[disc] || [];
                        return (
                          <td key={disc} className="py-1.5 text-center px-1">
                            {names.length > 0 ? (
                              <span className="text-green-500" title={names.join(", ")}>{names.length}</span>
                            ) : (
                              <span className="text-muted-foreground/30">-</span>
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </CardContent>
        </Card>

        {/* AI Readiness Spectrum */}
        <Card className="bg-card border-border">
          <CardHeader>
            <CardTitle className="text-lg">AI Readiness Spectrum</CardTitle>
          </CardHeader>
          <CardContent>
            {READINESS_TIERS.map((tier) => {
              const tierPeople = readiness.filter((r) => r.tier === tier.value);
              if (tierPeople.length === 0) return null;
              return (
                <div key={tier.value} className="mb-4 last:mb-0">
                  <div className="flex items-center gap-2 mb-2">
                    <div className={`w-3 h-3 rounded-full ${tier.color}`} />
                    <span className="text-sm font-medium">{tier.label}</span>
                    <span className="text-xs text-muted-foreground">({tierPeople.length})</span>
                  </div>
                  <div className="flex flex-wrap gap-1.5">
                    {tierPeople.map((person) => (
                      <Badge key={person.name} variant="outline" className="text-xs">
                        {person.name}
                      </Badge>
                    ))}
                  </div>
                </div>
              );
            })}
            {readiness.length === 0 && (
              <p className="text-sm text-muted-foreground">Readiness data not yet populated</p>
            )}
          </CardContent>
        </Card>
      </div>

      {/* Wave Progress */}
      <Card className="bg-card border-border">
        <CardHeader>
          <CardTitle className="text-lg">Wave Progress</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            {waves.map((wave) => {
              const pct = wave.total > 0 ? Math.round((wave.completed / wave.total) * 100) : 0;
              return (
                <div key={wave.number} className="bg-muted rounded-lg p-4">
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-sm font-medium">Wave {wave.number}</span>
                    <span className="text-xs text-muted-foreground">{wave.date_range}</span>
                  </div>
                  <div className="w-full bg-background rounded-full h-2 mb-2">
                    <div
                      className={`h-2 rounded-full ${pct === 100 ? "bg-green-500" : "bg-blue-500"}`}
                      style={{ width: `${pct}%` }}
                    />
                  </div>
                  <div className="flex items-center justify-between">
                    <span className="text-xs text-muted-foreground">{wave.completed}/{wave.total} complete</span>
                    <span className="text-xs font-mono">{pct}%</span>
                  </div>
                  <div className="flex flex-wrap gap-1 mt-2">
                    {wave.people.map((name) => (
                      <span key={name} className="text-[10px] text-muted-foreground">{name.split(" ")[0]}</span>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}

function StatCard({ label, value }: { label: string; value: number }) {
  return (
    <Card className="bg-card border-border">
      <CardContent className="pt-6">
        <p className="text-sm text-muted-foreground">{label}</p>
        <p className="text-3xl font-bold text-foreground mt-1">{value}</p>
      </CardContent>
    </Card>
  );
}
