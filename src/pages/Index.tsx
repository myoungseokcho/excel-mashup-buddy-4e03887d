import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Progress } from "@/components/ui/progress";
import { toast } from "sonner";
import { Upload, FileSpreadsheet, Download, X, Sparkles } from "lucide-react";

type FileData = {
  name: string;
  xData: (string | number | null)[];
  yData: (string | number | null)[];
};

const RANGE_ROWS = 101; // A5:B105 → 101 rows

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
        const xData: (string | number | null)[] = [];
        const yData: (string | number | null)[] = [];
        for (let r = 5; r <= 105; r++) {
          const aCell = ws[XLSX.utils.encode_cell({ c: 0, r: r - 1 })];
          const bCell = ws[XLSX.utils.encode_cell({ c: 1, r: r - 1 })];
          xData.push(aCell ? (aCell.v as string | number) : null);
          yData.push(bCell ? (bCell.v as string | number) : null);
        }
        next.push({ name: file.name, xData, yData });
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

  const previewRows = useMemo(() => {
    // Build preview matrix: row 1 = filename (every 2 cols), row 2 = X/Y headers, rows 3+ = data
    const cols = files.length * 2;
    if (cols === 0) return [];
    const rows: (string | number | null)[][] = [];
    // header 1: filename
    const h1: (string | number | null)[] = [];
    files.forEach((f) => {
      h1.push(f.name, "");
    });
    rows.push(h1);
    // header 2: X data / Y data
    const h2: (string | number | null)[] = [];
    files.forEach(() => h2.push("X data", "Y data"));
    rows.push(h2);
    for (let r = 0; r < RANGE_ROWS; r++) {
      const row: (string | number | null)[] = [];
      files.forEach((f) => {
        row.push(f.xData[r], f.yData[r]);
      });
      rows.push(row);
    }
    return rows;
  }, [files]);

  const handleExport = () => {
    if (files.length === 0) {
      toast.error("파일을 먼저 업로드해주세요");
      return;
    }
    // Build sheet matching spec: A1 filename, A5:A105 X, B5:B105 Y, with "X data"/"Y data" labels above
    const totalCols = files.length * 2;
    const aoa: (string | number | null)[][] = [];
    // Initialize 105 rows
    for (let r = 0; r < 105; r++) aoa.push(new Array(totalCols).fill(null));

    files.forEach((f, idx) => {
      const xCol = idx * 2;
      const yCol = idx * 2 + 1;
      aoa[0][xCol] = f.name; // row 1
      aoa[3][xCol] = "X data"; // row 4 (header above row 5)
      aoa[3][yCol] = "Y data";
      for (let r = 0; r < RANGE_ROWS; r++) {
        aoa[4 + r][xCol] = f.xData[r]; // row 5 onward
        aoa[4 + r][yCol] = f.yData[r];
      }
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Merged");
    const filename = (outputName.trim() || "merged_data") + ".xlsx";
    XLSX.writeFile(wb, filename);
    toast.success(`${filename} 저장 완료`);
  };

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
              {files.map((f, i) => (
                <li
                  key={i}
                  className="flex items-center justify-between rounded-lg bg-secondary px-4 py-2.5"
                >
                  <div className="flex items-center gap-3 min-w-0">
                    <FileSpreadsheet className="h-4 w-4 text-primary shrink-0" />
                    <span className="text-sm font-medium truncate">{f.name}</span>
                    <span className="text-xs text-muted-foreground shrink-0">
                      열 {String.fromCharCode(65 + i * 2)} / {String.fromCharCode(65 + i * 2 + 1)}
                    </span>
                  </div>
                  <button
                    onClick={() => removeFile(i)}
                    className="text-muted-foreground hover:text-destructive transition-colors"
                  >
                    <X className="h-4 w-4" />
                  </button>
                </li>
              ))}
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
                    <tr key={ri} className={ri < 2 ? "bg-accent font-semibold" : "even:bg-muted/40"}>
                      <td className="px-2 py-1 text-muted-foreground border-r font-mono w-12 text-center">
                        {ri === 0 ? 1 : ri === 1 ? 4 : ri + 3}
                      </td>
                      {row.map((cell, ci) => (
                        <td
                          key={ci}
                          className="px-3 py-1 border-r whitespace-nowrap"
                          style={{ minWidth: 100 }}
                        >
                          {cell ?? ""}
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
