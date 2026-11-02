export type ColType = "Integer" | "Decimal" | "Date" | "DateTime" | "Text" | "Mixed";
export type DateOrder = "auto" | "dmy" | "mdy";
export type OutputFormat = "mm/dd/yyyy" | "dd/mm/yyyy" | "yyyy-mm-dd";
export type Overrides = Record<string, Record<string, ColType | "Auto">>;

export interface Column {
  index: number;
  letter: string;
  header: string | null;
  detected_type: ColType;
  effective_type: ColType;
  overridden: boolean;
  confidence: number;
  non_empty: number;
  converted: number;
  warning_count: number;
  notes: string[];
  samples: string[];
}

export interface Sheet {
  name: string;
  rows: number;
  columns: Column[];
}

export interface WarningGroup {
  ambiguous?: boolean;
  sheet: string;
  column: string;
  header: string | null;
  reason: string;
  count: number;
  locations: { cell: string; value: string }[];
}

export interface Result {
  job_id: string;
  filename: string;
  file_type: "xlsx" | "csv";
  download_name: string;
  download_url: string;
  sheets: Sheet[];
  warnings: WarningGroup[];
  warning_total: number;
  ambiguous_total?: number;
  converted_total: number;
  notices: string[];
  csv: { encoding: string; delimiter: string; note: string } | null;
  options: { date_order: DateOrder; output_format: OutputFormat };
}
