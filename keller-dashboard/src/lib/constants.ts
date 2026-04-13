export const INTERVIEW_STATUS = [
  { value: "scheduled", label: "Scheduled", color: "bg-slate-500" },
  { value: "completed", label: "Completed", color: "bg-blue-500" },
  { value: "extracted", label: "Extracted", color: "bg-purple-500" },
  { value: "routed", label: "Routed", color: "bg-green-500" },
] as const;

export const DELIVERABLE_STATUS = [
  { value: "not_started", label: "Not Started", color: "bg-slate-500" },
  { value: "inputs_gathering", label: "Gathering Inputs", color: "bg-blue-500" },
  { value: "drafting", label: "Drafting", color: "bg-purple-500" },
  { value: "review", label: "In Review", color: "bg-amber-500" },
  { value: "final", label: "Final", color: "bg-green-500" },
] as const;

export const READINESS_TIERS = [
  { value: "champion", label: "Champion", color: "bg-green-500" },
  { value: "early_adopter", label: "Early Adopter", color: "bg-blue-500" },
  { value: "pragmatist", label: "Pragmatist", color: "bg-amber-500" },
  { value: "skeptic", label: "Skeptic", color: "bg-red-400" },
] as const;

export const RESEARCH_PRIORITY = [
  { value: "P1", label: "P1", color: "text-red-500" },
  { value: "P2", label: "P2", color: "text-amber-500" },
  { value: "P3", label: "P3", color: "text-gray-400" },
] as const;

export const KELLER_OFFICES = [
  "Meridian", "Idaho Falls", "Pocatello", "Coeur d'Alene",
  "Salem", "Beaverton", "Bend",
  "Richland", "Kent", "Clarkston",
  "Reno", "Provo",
] as const;

export const DISCIPLINES = [
  "Water/Wastewater", "Transportation", "Structural",
  "Electrical", "Civil/Site", "Construction Mgmt", "Survey",
] as const;

export const ENGAGEMENT_MILESTONES = {
  discovery_close: "2026-04-22",
  synthesis_start: "2026-04-23",
  larry_review: "2026-04-30",
  shareholder_meeting: "2026-05-07",
} as const;

export function formatCurrency(value: number | null): string {
  if (!value) return "\u2014";
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(value);
}

export function daysUntil(dateStr: string): number {
  const diff = new Date(dateStr).getTime() - Date.now();
  return Math.ceil(diff / (1000 * 60 * 60 * 24));
}

export function daysSince(dateStr: string | null): number | null {
  if (!dateStr) return null;
  const diff = Date.now() - new Date(dateStr).getTime();
  return Math.floor(diff / (1000 * 60 * 60 * 24));
}

export function statusColor(statusValue: string, statusList: readonly { value: string; color: string }[]): string {
  return statusList.find((s) => s.value === statusValue)?.color || "bg-slate-500";
}

export function statusLabel(statusValue: string, statusList: readonly { value: string; label: string }[]): string {
  return statusList.find((s) => s.value === statusValue)?.label || statusValue;
}
