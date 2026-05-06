import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx-js-style";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Progress } from "@/components/ui/progress";
import { toast } from "sonner";
import { Upload, FileSpreadsheet, Download, X, Sparkles } from "lucide-react";

type FileData = {
  name: string;
  freq: (string | number | null)[];
  cp: (string | number | null)[];
  thickness: [string, string, string];
};

const RANGE_ROWS = 101; // A5:B105 → 101 rows

// Column index for a file's Cp / filename / average / formula cell.
// file 0 → col 1 (B); file i>=1 → col i+1
const cpCol = (i: number) => (i === 0 ? 1 : i + 1);
const nameCol = (i: number) => (i === 0 ? 1 : i + 1);

const avgOf = (t: [string, string, string]) => {
  const nums = t.map((v) => parseFloat(v)).filter((n) => !isNaN(n));
  if (nums.length === 0) return null;
  return nums.reduce((a, b) => a + b, 0) / nums.length - 0.3;
};

const Index = () => {
  const [files, setFiles] = useState<FileData[]>([]);
  const [progress, setProgress] = useState(0);
  const [loading, setLoading] = useState(false);
  const [outputName, setOutputName] = useState("merged_data");

  const handleUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const list = e.target.files;
    if (!list || list.length === 0) return;
    setLoading(true);
    setProgress(0);
    const next: FileData[] = [];
    for (let i = 0; i < list.length; i++) {
      const file = list[i];
      try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const freq: (string | number | null)[] = [];
        const cp: (string | number | null)[] = [];
        for (let r = 5; r <= 105; r++) {
          const aCell = ws[XLSX.utils.encode_cell({ c: 0, r: r - 1 })];
          const bCell = ws[XLSX.utils.encode_cell({ c: 1, r: r - 1 })];
          freq.push(aCell ? (aCell.v as string | number) : null);
          cp.push(bCell ? (bCell.v as string | number) : null);
        }
        next.push({ name: file.name, freq, cp, thickness: ["", "", ""] });
      } catch (err) {
        toast.error(`${file.name} 처리 실패`);
      }
      setProgress(Math.round(((i + 1) / list.length) * 100));
    }
    setFiles((prev) => [...prev, ...next]);
    setLoading(false);
    toast.success(`${next.length}개 파일 통합 완료`);
    e.target.value = "";
  }, []);

  const removeFile = (idx: number) => {
    setFiles((prev) => prev.filter((_, i) => i !== idx));
  };

  const clearAll = () => setFiles([]);

  const updateThickness = (fileIdx: number, tIdx: 0 | 1 | 2, value: string) => {
    setFiles((prev) =>
      prev.map((f, i) => {
        if (i !== fileIdx) return f;
        const t = [...f.thickness] as [string, string, string];
        t[tIdx] = value;
        return { ...f, thickness: t };
      }),
    );
  };

  // Total columns needed: max(nameCol, cpCol) + 1 across all files
  const totalCols = useMemo(() => {
    if (files.length === 0) return 0;
    let max = 0;
    files.forEach((_, i) => {
      max = Math.max(max, nameCol(i), cpCol(i));
    });
    return max + 1;
  }, [files]);

  // Build the merged AOA (105 rows tall)
  const mergedAoa = useMemo(() => {
    if (files.length === 0) return [] as (string | number | null)[][];
    const aoa: (string | number | null)[][] = [];
    for (let r = 0; r < 105; r++) aoa.push(new Array(totalCols).fill(null));

    files.forEach((f, i) => {
      // Row 1: filename
      aoa[0][nameCol(i)] = f.name;
      // Row 2: average thickness (in cpCol)
      const avg = avgOf(f.thickness);
      aoa[1][cpCol(i)] = avg;
      // Row 3: formula value = Cp(row80) * avg * 1e12 * 8.854 / 78.5
      const cp80 = f.cp[75]; // Excel row 80 = data index 75
      if (avg !== null && typeof cp80 === "number") {
        aoa[2][cpCol(i)] = (cp80 * avg * 1e12 * 8.854) / 78.5;
      }
      // Frequency only from first file (column A, rows 5-105)
      if (i === 0) {
        for (let r = 0; r < RANGE_ROWS; r++) {
          aoa[4 + r][0] = f.freq[r];
        }
      }
      // Cp data (rows 5-105) in cpCol
      for (let r = 0; r < RANGE_ROWS; r++) {
        aoa[4 + r][cpCol(i)] = f.cp[r];
      }
    });

    return aoa;
  }, [files, totalCols]);

  const handleExport = () => {
    if (files.length === 0) {
      toast.error("파일을 먼저 업로드해주세요");
      return;
    }

    // Merged sheet
    const ws = XLSX.utils.aoa_to_sheet(mergedAoa);
    // Apply yellow fill to entire row 3
    for (let c = 0; c < totalCols; c++) {
      const addr = XLSX.utils.encode_cell({ c, r: 2 });
      if (!ws[addr]) ws[addr] = { t: "z", v: null };
      ws[addr].s = {
        ...(ws[addr].s || {}),
        fill: { patternType: "solid", fgColor: { rgb: "FFFF00" } },
      };
    }
    if (!ws["!ref"]) ws["!ref"] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: totalCols - 1, r: mergedAoa.length - 1 } });

    // Thickness raw sheet
    const tAoa: (string | number | null)[][] = [];
    for (let r = 0; r < 6; r++) tAoa.push(new Array(totalCols).fill(null));
    tAoa[0][0] = "파일명";
    tAoa[1][0] = "두께1";
    tAoa[2][0] = "두께2";
    tAoa[3][0] = "두께3";
    tAoa[4][0] = "평균";
    tAoa[5][0] = "보정값";
    files.forEach((f, i) => {
      const c = cpCol(i);
      tAoa[0][c] = f.name;
      const nums: number[] = [];
      [0, 1, 2].forEach((k) => {
        const v = parseFloat(f.thickness[k]);
        if (!isNaN(v)) nums.push(v);
        tAoa[1 + k][c] = isNaN(v) ? f.thickness[k] || null : v;
      });
      if (nums.length > 0) {
        const rawAvg = nums.reduce((a, b) => a + b, 0) / nums.length;
        tAoa[4][c] = rawAvg;
        tAoa[5][c] = rawAvg - 0.3;
      }
    });
    const tWs = XLSX.utils.aoa_to_sheet(tAoa);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Merged");
    XLSX.utils.book_append_sheet(wb, tWs, "두께 Raw");
    const filename = (outputName.trim() || "merged_data") + ".xlsx";
    XLSX.writeFile(wb, filename);
    toast.success(`${filename} 저장 완료`);
  };

  // For preview, derive header row meta
  const excelRowLabel = (previewIdx: number) => {
    // previewIdx 0 → Excel row 1
    // previewIdx 1 → Excel row 2
    // previewIdx 2 → Excel row 3
    // previewIdx 3 → Excel row 5 (skip row 4)
    if (previewIdx <= 2) return previewIdx + 1;
    return previewIdx + 2;
  };

  // Build preview: rows 1,2,3 then data rows starting Excel row 5
  const previewRows = useMemo(() => {
    if (mergedAoa.length === 0) return [];
    const rows: (string | number | null)[][] = [];
    rows.push(mergedAoa[0]); // row 1 filenames
    rows.push(mergedAoa[1]); // row 2 averages
    rows.push(mergedAoa[2]); // row 3 formula
    // skip row 4 (empty / removed X data Y data)
    for (let r = 4; r < mergedAoa.length; r++) rows.push(mergedAoa[r]);
    return rows;
  }, [mergedAoa]);

  return (
    <div className="min-h-screen" style={{ background: "var(--gradient-subtle)" }}>
      <div className="container max-w-7xl py-12">
        <header className="mb-10 text-center">
          <div className="inline-flex items-center gap-2 rounded-full bg-accent px-4 py-1.5 text-sm font-medium text-accent-foreground mb-4">
            <Sparkles className="h-4 w-4" />
            Excel Range Merger
          </div>
          <h1 className="text-4xl md:text-5xl font-bold tracking-tight text-foreground mb-3">
            여러 엑셀을 하나로 통합
          </h1>
          <p className="text-muted-foreground text-lg">
            각 파일의 A5:B105 범위를 추출해 가로로 정렬합니다
          </p>
        </header>

        <Card
          className="p-8 mb-6 border-2 border-dashed"
          style={{ boxShadow: "var(--shadow-card)" }}
        >
          <label className="flex flex-col items-center justify-center cursor-pointer gap-3 py-8">
            <div
              className="h-16 w-16 rounded-2xl flex items-center justify-center text-primary-foreground"
              style={{ background: "var(--gradient-primary)", boxShadow: "var(--shadow-elegant)" }}
            >
              <Upload className="h-8 w-8" />
            </div>
            <div className="text-center">
              <div className="font-semibold text-foreground">엑셀 파일 업로드</div>
              <div className="text-sm text-muted-foreground">.xlsx, .xls 다중 선택 가능</div>
            </div>
            <input
              type="file"
              multiple
              accept=".xlsx,.xls"
              onChange={handleUpload}
              className="hidden"
              disabled={loading}
            />
          </label>
          {loading && (
            <div className="mt-4">
              <Progress value={progress} />
              <div className="text-xs text-muted-foreground mt-2 text-center">{progress}%</div>
            </div>
          )}
        </Card>

        {files.length > 0 && (
          <Card className="p-6 mb-6" style={{ boxShadow: "var(--shadow-card)" }}>
            <div className="flex items-center justify-between mb-4">
              <h2 className="font-semibold text-foreground">업로드된 파일 ({files.length})</h2>
              <Button variant="ghost" size="sm" onClick={clearAll}>
                전체 삭제
              </Button>
            </div>
            <ul className="space-y-2">
              {files.map((f, i) => {
                const avg = avgOf(f.thickness);
                return (
                  <li
                    key={i}
                    className="flex flex-wrap items-center gap-3 rounded-lg bg-secondary px-4 py-2.5"
                  >
                    <FileSpreadsheet className="h-4 w-4 text-primary shrink-0" />
                    <span className="text-sm font-medium truncate max-w-[260px]">{f.name}</span>
                    <div className="flex items-center gap-2">
                      {[0, 1, 2].map((k) => (
                        <Input
                          key={k}
                          type="number"
                          step="any"
                          placeholder={`두께${k + 1}`}
                          value={f.thickness[k]}
                          onChange={(e) => updateThickness(i, k as 0 | 1 | 2, e.target.value)}
                          className="h-8 w-24 bg-background"
                        />
                      ))}
                    </div>
                    <span className="text-xs text-muted-foreground">
                      평균: {avg !== null ? avg.toFixed(4) : "—"}
                    </span>
                    <button
                      onClick={() => removeFile(i)}
                      className="ml-auto text-muted-foreground hover:text-destructive transition-colors"
                    >
                      <X className="h-4 w-4" />
                    </button>
                  </li>
                );
              })}
            </ul>
          </Card>
        )}

        {files.length > 0 && (
          <Card className="p-6 mb-6" style={{ boxShadow: "var(--shadow-card)" }}>
            <h2 className="font-semibold text-foreground mb-4">통합 데이터 미리보기</h2>
            <div className="overflow-auto max-h-[500px] rounded-lg border">
              <table className="text-xs">
                <tbody>
                  {previewRows.slice(0, 50).map((row, ri) => (
                    <tr key={ri} className={ri < 3 ? "bg-accent font-semibold" : "even:bg-muted/40"}>
                      <td className="px-2 py-1 text-muted-foreground border-r font-mono w-12 text-center">
                        {excelRowLabel(ri)}
                      </td>
                      {row.map((cell, ci) => (
                        <td
                          key={ci}
                          className="px-3 py-1 border-r whitespace-nowrap"
                          style={{ minWidth: 100 }}
                        >
                          {typeof cell === "number"
                            ? Number.isInteger(cell)
                              ? cell
                              : cell.toPrecision(6)
                            : cell ?? ""}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {previewRows.length > 50 && (
              <div className="text-xs text-muted-foreground mt-2 text-center">
                미리보기는 상위 50행만 표시 · 전체 {previewRows.length}행은 다운로드 파일에 포함
              </div>
            )}
          </Card>
        )}

        {files.length > 0 && (
          <Card className="p-6" style={{ boxShadow: "var(--shadow-card)" }}>
            <h2 className="font-semibold text-foreground mb-4">엑셀로 저장</h2>
            <div className="flex flex-col sm:flex-row gap-3">
              <Input
                value={outputName}
                onChange={(e) => setOutputName(e.target.value)}
                placeholder="파일명"
                className="flex-1"
              />
              <Button onClick={handleExport} size="lg" className="gap-2">
                <Download className="h-4 w-4" />
                다운로드
              </Button>
            </div>
          </Card>
        )}
      </div>
    </div>
  );
};

export default Index;
