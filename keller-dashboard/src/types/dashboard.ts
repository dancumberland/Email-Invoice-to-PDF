export interface DashboardData {
  generated_at: string;
  engagement: {
    name: string;
    value: number;
    start_date: string;
    end_date: string;
    current_week: number;
    days_remaining: number;
    phase: string;
  };
  people: Person[];
  coverage: {
    offices: string[];
    disciplines: string[];
    matrix: Record<string, Record<string, string[]>>;
  };
  readiness: ReadinessEntry[];
  waves: Wave[];
  deliverables: Deliverable[];
  research: ResearchItem[];
  interviews: ScheduledInterview[];
  pipeline: PipelineItem[];
  actions: ActionItem[];
  comms: CommEntry[];
  invoices: Invoice[];
  findings: Finding[];
  themes: Theme[];
  quickWins: QuickWin[];
  contradictions: Contradiction[];
  quotes: Quote[];
  gaps: Gap[];
  agent: AgentStatus | null;
  clicks: ClicksData | null;
}

export interface Person {
  name: string;
  role: string;
  office: string;
  discipline: string;
  wave: number;
  interview_date: string | null;
  interview_status: "scheduled" | "completed" | "extracted" | "routed";
  brief_path: string | null;
  profile_path: string | null;
  questions_path: string | null;
  readiness_tier: "champion" | "early_adopter" | "pragmatist" | "skeptic" | null;
  key_insight: string | null;
}

export interface ReadinessEntry {
  name: string;
  tier: "champion" | "early_adopter" | "pragmatist" | "skeptic";
  role: string;
}

export interface Wave {
  number: number;
  date_range: string;
  people: string[];
  completed: number;
  total: number;
}

export interface Deliverable {
  id: number;
  title: string;
  sow_spec: string;
  status: "not_started" | "inputs_gathering" | "drafting" | "review" | "final";
  target_date: string;
  inputs_needed: InputItem[];
  inputs_available: number;
  inputs_total: number;
  depends_on: number[];
  notes: string | null;
}

export interface InputItem {
  label: string;
  type: "interview" | "research" | "analysis" | "other";
  available: boolean;
  source_path: string | null;
}

export interface ResearchItem {
  topic: string;
  priority: "P1" | "P2" | "P3";
  status: "complete" | "in_progress" | "queued";
  file_path: string | null;
  feeds_deliverable: number[];
}

export interface PipelineItem {
  person: string;
  transcript_date: string;
  stage: "transcript" | "extracted" | "routed";
  brief_path: string | null;
}

export interface ScheduledInterview {
  person: string;
  role: string;
  date: string;
  time: string;
  time_mt: string;
  questions_path: string | null;
  profile_path: string | null;
}

export interface Invoice {
  milestone: number;
  amount: number;
  description: string;
  due_description: string;
  status: "paid" | "sent" | "upcoming";
}

export interface Finding {
  id: number;
  statement: string;
  evidence_count: number;
  departments: string[];
  roadmap_implication: string;
  theme: string;
}

export interface Theme {
  name: string;
  slug: string;
  description: string;
  finding_ids: number[];
}

export interface QuickWin {
  title: string;
  description: string;
  who_benefits: string;
  effort: "trivial" | "low" | "medium";
  prerequisite: string | null;
}

export interface Quote {
  text: string;
  attribution: string;
  theme: string;
  usable_in_presentation: boolean;
}

export interface ActionItem {
  description: string;
  owner: string;
  waiting_on: string | null;
  priority: "high" | "medium" | "low";
}

export interface CommEntry {
  date: string;
  type: "email" | "meeting" | "call";
  with_person: string;
  summary: string;
}

export interface Contradiction {
  statement: string;
  sources: string[];
}

export interface Gap {
  question: string;
  priority: "critical" | "important" | "nice_to_have";
  plan_to_fill: string;
}

export interface AgentCompletedPerson {
  first_name: string;
  role: string;
  office: string;
}

export interface AgentStatus {
  total_conversations: number;
  good_conversations: number;
  completed_people: AgentCompletedPerson[];
  person_attempts: Record<string, number>;
  last_check: string;
  phone_number: string;
  phone_active: boolean;
}

export interface Clicker {
  first_name: string;
  role: string;
  office: string;
  first_click: string;
  interviewed: boolean;
  issue?: string;
}

export interface ClicksData {
  last_updated: string;
  employees_total: number;
  unique_clickers: number;
  clickers: Clicker[];
}
