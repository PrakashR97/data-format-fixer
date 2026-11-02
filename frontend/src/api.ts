import type { DateOrder, OutputFormat, Overrides, Result } from "./types";

async function parse(res: Response): Promise<Result> {
  if (!res.ok) {
    let message = `Request failed (${res.status}).`;
    try {
      const body = await res.json();
      if (typeof body.detail === "string") message = body.detail;
    } catch {
      /* keep default message */
    }
    throw new Error(message);
  }
  return res.json();
}

export async function uploadFile(file: File, dateOrder: DateOrder, outputFormat: OutputFormat) {
  const form = new FormData();
  form.append("file", file);
  form.append("date_order", dateOrder);
  form.append("output_format", outputFormat);
  return parse(await fetch("/api/jobs", { method: "POST", body: form }));
}

export async function regenerate(
  jobId: string,
  dateOrder: DateOrder,
  outputFormat: OutputFormat,
  overrides: Overrides,
) {
  return parse(
    await fetch(`/api/jobs/${jobId}/regenerate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ date_order: dateOrder, output_format: outputFormat, overrides }),
    }),
  );
}

export function discard(jobId: string) {
  return fetch(`/api/jobs/${jobId}`, { method: "DELETE" }).catch(() => undefined);
}

export function triggerDownload(url: string, name: string) {
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
}
