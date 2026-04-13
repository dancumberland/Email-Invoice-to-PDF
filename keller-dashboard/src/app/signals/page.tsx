import { getDashboardData } from "@/lib/data";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Lightbulb, MessageSquareQuote, Zap, HelpCircle, AlertTriangle } from "lucide-react";

export const dynamic = "force-dynamic";

export default function SignalsPage() {
  const data = getDashboardData();
  const { findings, themes, quickWins, contradictions, quotes, gaps } = data;

  return (
    <div className="space-y-6">
      <h1 className="text-2xl font-bold">Signals</h1>

      {/* Top Findings */}
      {findings.length > 0 && (
        <div className="space-y-4">
          <h2 className="text-lg font-semibold flex items-center gap-2">
            <Lightbulb className="h-5 w-5 text-amber-500" />
            Top Findings
          </h2>
          <div className="space-y-3">
            {findings.map((finding) => (
              <Card key={finding.id} className="bg-card border-border">
                <CardContent className="pt-4">
                  <div className="flex items-start justify-between gap-4">
                    <div className="flex-1">
                      <p className="text-sm font-medium">{finding.statement}</p>
                      <p className="text-xs text-muted-foreground mt-1">{finding.roadmap_implication}</p>
                      <div className="flex flex-wrap gap-1 mt-2">
                        {finding.departments.map((dept) => (
                          <Badge key={dept} variant="outline" className="text-[10px]">{dept}</Badge>
                        ))}
                      </div>
                    </div>
                    <div className="text-right flex-shrink-0">
                      <div className="text-2xl font-bold text-foreground">{finding.evidence_count}</div>
                      <div className="text-[10px] text-muted-foreground">sources</div>
                    </div>
                  </div>
                </CardContent>
              </Card>
            ))}
          </div>
        </div>
      )}

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Theme Clusters */}
        {themes.length > 0 && (
          <Card className="bg-card border-border">
            <CardHeader>
              <CardTitle className="text-lg">Theme Clusters</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              {themes.map((theme) => {
                const themeFindings = findings.filter((f) => theme.finding_ids.includes(f.id));
                return (
                  <div key={theme.slug} className="bg-muted rounded-lg p-3">
                    <p className="text-sm font-medium">{theme.name}</p>
                    <p className="text-xs text-muted-foreground mt-1">{theme.description}</p>
                    {themeFindings.length > 0 && (
                      <div className="mt-2 space-y-1">
                        {themeFindings.map((f) => (
                          <p key={f.id} className="text-xs text-muted-foreground pl-2 border-l-2 border-border">
                            {f.statement}
                          </p>
                        ))}
                      </div>
                    )}
                  </div>
                );
              })}
            </CardContent>
          </Card>
        )}

        {/* Quick Wins */}
        {quickWins.length > 0 && (
          <Card className="bg-card border-border">
            <CardHeader>
              <CardTitle className="text-lg flex items-center gap-2">
                <Zap className="h-4 w-4 text-green-500" />
                Quick Wins
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="space-y-3">
                {quickWins.map((win, i) => (
                  <div key={i} className="bg-muted rounded-lg p-3">
                    <div className="flex items-start justify-between">
                      <p className="text-sm font-medium">{win.title}</p>
                      <Badge variant="outline" className={`text-[10px] ${
                        win.effort === "trivial" ? "border-green-500 text-green-500" :
                        win.effort === "low" ? "border-blue-500 text-blue-500" :
                        "border-amber-500 text-amber-500"
                      }`}>
                        {win.effort}
                      </Badge>
                    </div>
                    <p className="text-xs text-muted-foreground mt-1">{win.description}</p>
                    <p className="text-xs text-muted-foreground mt-1">Benefits: {win.who_benefits}</p>
                    {win.prerequisite && (
                      <p className="text-xs text-amber-500 mt-1">Prereq: {win.prerequisite}</p>
                    )}
                  </div>
                ))}
              </div>
            </CardContent>
          </Card>
        )}
      </div>

      {/* Quotes Board */}
      {quotes.length > 0 && (
        <Card className="bg-card border-border">
          <CardHeader>
            <CardTitle className="text-lg flex items-center gap-2">
              <MessageSquareQuote className="h-4 w-4" />
              Quotes Board
            </CardTitle>
          </CardHeader>
          <CardContent>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
              {quotes.map((quote, i) => (
                <div key={i} className="bg-muted rounded-lg p-3 border-l-2 border-blue-500">
                  <p className="text-sm italic">&ldquo;{quote.text}&rdquo;</p>
                  <div className="flex items-center justify-between mt-2">
                    <span className="text-xs text-muted-foreground">&mdash; {quote.attribution}</span>
                    {quote.usable_in_presentation && (
                      <Badge variant="outline" className="text-[9px] border-green-500 text-green-500">keynote</Badge>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </CardContent>
        </Card>
      )}

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Contradictions */}
        {contradictions.length > 0 && (
          <Card className="bg-card border-border">
            <CardHeader>
              <CardTitle className="text-lg flex items-center gap-2">
                <AlertTriangle className="h-4 w-4 text-amber-500" />
                Contradictions
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-3">
              {contradictions.map((c, i) => (
                <div key={i} className="bg-muted rounded-lg p-3">
                  <p className="text-sm">{c.statement}</p>
                  <p className="text-xs text-muted-foreground mt-1">
                    Sources: {c.sources.join(", ")}
                  </p>
                </div>
              ))}
            </CardContent>
          </Card>
        )}

        {/* Unanswered Questions */}
        {gaps.length > 0 && (
          <Card className="bg-card border-border">
            <CardHeader>
              <CardTitle className="text-lg flex items-center gap-2">
                <HelpCircle className="h-4 w-4 text-blue-500" />
                Unanswered Questions
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-3">
              {gaps.map((gap, i) => (
                <div key={i} className="bg-muted rounded-lg p-3">
                  <div className="flex items-start justify-between">
                    <p className="text-sm">{gap.question}</p>
                    <Badge variant="outline" className={`text-[10px] flex-shrink-0 ml-2 ${
                      gap.priority === "critical" ? "border-red-500 text-red-500" :
                      gap.priority === "important" ? "border-amber-500 text-amber-500" :
                      "border-gray-500 text-gray-500"
                    }`}>
                      {gap.priority}
                    </Badge>
                  </div>
                  <p className="text-xs text-muted-foreground mt-1">{gap.plan_to_fill}</p>
                </div>
              ))}
            </CardContent>
          </Card>
        )}
      </div>

      {/* Empty State */}
      {findings.length === 0 && themes.length === 0 && quickWins.length === 0 && (
        <Card className="bg-card border-border">
          <CardContent className="py-12 text-center">
            <p className="text-muted-foreground">Signals data not yet populated.</p>
            <p className="text-sm text-muted-foreground/70 mt-2">
              Edit signals.json to add findings, themes, quotes, and quick wins.
            </p>
          </CardContent>
        </Card>
      )}
    </div>
  );
}
