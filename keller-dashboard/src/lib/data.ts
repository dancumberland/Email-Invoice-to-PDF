import { readFileSync } from "fs";
import { join } from "path";
import type { DashboardData } from "@/types/dashboard";

const DATA_PATH = join(process.cwd(), "data", "data.json");

export function getDashboardData(): DashboardData {
  const raw = readFileSync(DATA_PATH, "utf-8");
  return JSON.parse(raw);
}
