import React, { useState, useRef, useEffect, useMemo } from "react";
import {
  UploadCloud,
  Trash2,
  TrendingDown,
  Store,
  AlertCircle,
  FileSpreadsheet,
  Loader2,
  CheckCircle2,
  Calendar,
  Filter,
  XCircle,
  ArrowRight,
  ShieldCheck,
  ShieldAlert,
  Minus,
  ArrowUpDown,
} from "lucide-react";

export default function App() {
  const [files, setFiles] = useState([]);
  const [parsedData, setParsedData] = useState([]);
  const [isDragging, setIsDragging] = useState(false);
  const [isLoadingXlsx, setIsLoadingXlsx] = useState(true);

  // Progress and Processing State
  const [isProcessing, setIsProcessing] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);

  // Filter & Sort States
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [showOnlyWins, setShowOnlyWins] = useState(false);
  const [sortBy, setSortBy] = useState("date-desc");

  const fileInputRef = useRef(null);

  useEffect(() => {
    // Inject Custom Browser Tab Icon (Favicon)
    let link = document.querySelector("link[rel~='icon']");
    if (!link) {
      link = document.createElement("link");
      link.rel = "icon";
      document.head.appendChild(link);
    }
    link.href =
      'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text y=".9em" font-size="90">🎯</text></svg>';
    document.title = "MSM Markup Tracker";

    // Load Excel Engine
    if (window.XLSX) {
      setIsLoadingXlsx(false);
      return;
    }
    const script = document.createElement("script");
    script.src =
      "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    script.async = true;
    script.onload = () => setIsLoadingXlsx(false);
    document.body.appendChild(script);
  }, []);

  // --- Core Parsing: Extracting Timeline ---
  const processSheetData = (json, filename) => {
    let type = "Unknown";
    if (filename.toLowerCase().includes("del")) type = "Delivery";
    if (filename.toLowerCase().includes("pick")) type = "Pickup";

    let headerRowIdx = -1;
    for (let i = 0; i < json.length; i++) {
      if (
        json[i] &&
        json[i][0] &&
        String(json[i][0]).toLowerCase().includes("store id")
      ) {
        headerRowIdx = i;
        break;
      }
    }

    if (headerRowIdx === -1) return [];
    const headers = json[headerRowIdx].map((h) => (h ? String(h).trim() : ""));

    const defDateIdx = headers.findIndex(
      (h) =>
        h.toLowerCase().includes("defilation date") ||
        h.toLowerCase().includes("deflation date"),
    );
    const data = [];

    if (defDateIdx !== -1) {
      // FORMAT A: Legacy Format
      const defRateIdx = headers.findIndex(
        (h) =>
          h.toLowerCase().includes("defilated rate") ||
          h.toLowerCase().includes("deflated rate"),
      );

      for (let i = headerRowIdx + 1; i < json.length; i++) {
        const row = json[i];
        if (!row) continue;
        const storeId = row[0];
        if (!storeId || storeId === "0" || storeId === "NaN" || storeId === "-")
          continue;

        let rawDefRate =
          defRateIdx !== -1 ? parseFloat(row[defRateIdx]) || 0 : 0;
        let defDate =
          defDateIdx !== -1 && row[defDateIdx]
            ? String(row[defDateIdx])
            : "2000-01-01";
        if (defDate === "-" || defDate === "NaN") defDate = "2000-01-01";

        rawDefRate = Math.max(0, Math.round(rawDefRate));

        const timeline = [];
        if (rawDefRate > 0) {
          timeline.push({ date: "1999-01-01", rate: rawDefRate });
          timeline.push({ date: defDate, rate: 0 });
        } else {
          timeline.push({ date: "1999-01-01", rate: 0 });
          timeline.push({ date: defDate, rate: 0 });
        }

        data.push({
          id: Math.random().toString(36).substr(2, 9),
          filename,
          type,
          storeId: String(storeId),
          timeline,
        });
      }
    } else {
      // FORMAT B: Sigma Daily Export (Extract full timeline)
      const dateCols = [];
      for (let c = 1; c < headers.length; c++) {
        const h = headers[c];
        if (h && isNaN(h) && !isNaN(Date.parse(h)))
          dateCols.push({ index: c, dateStr: h });
      }
      dateCols.sort((a, b) => new Date(a.dateStr) - new Date(b.dateStr));

      for (let i = headerRowIdx + 1; i < json.length; i++) {
        const row = json[i];
        if (!row) continue;
        const storeId = row[0];
        if (!storeId || storeId === "0" || storeId === "NaN" || storeId === "-")
          continue;

        let lastKnownRate = null;
        const timeline = dateCols.map((dc) => {
          const rawVal = row[dc.index];
          const isBlank =
            rawVal === null ||
            rawVal === undefined ||
            rawVal === "" ||
            rawVal === "-" ||
            String(rawVal).toLowerCase() === "nan";

          let parsedRate = 0;
          if (isBlank) {
            // LOCF (Last Observation Carried Forward)
            parsedRate = lastKnownRate !== null ? lastKnownRate : 0;
          } else {
            let val = String(rawVal).replace("%", "").trim();
            parsedRate = parseFloat(val) || 0;
            if (parsedRate > 0 && parsedRate <= 2) parsedRate *= 100;
            lastKnownRate = parsedRate;
          }

          const finalRate = Math.max(0, Math.round(parsedRate));
          return { date: dc.dateStr, rate: finalRate };
        });

        if (timeline.length > 0) {
          data.push({
            id: Math.random().toString(36).substr(2, 9),
            filename,
            type,
            storeId: String(storeId),
            timeline,
          });
        }
      }
    }
    return data;
  };

  const handleFileUpload = async (uploadedFiles) => {
    if (!window.XLSX)
      return alert("Excel engine is still loading. Please wait a second.");

    setIsProcessing(true);
    setUploadProgress(0);

    const newFiles = Array.from(uploadedFiles);
    setFiles((prev) => [...prev, ...newFiles]);
    const allNewData = [];

    for (let i = 0; i < newFiles.length; i++) {
      const file = newFiles[i];
      const data = await new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(new Uint8Array(e.target.result));
        reader.readAsArrayBuffer(file);
      });

      setUploadProgress(((i + 0.5) / newFiles.length) * 100);
      await new Promise((r) => setTimeout(r, 50));

      const wb = window.XLSX.read(data, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = window.XLSX.utils.sheet_to_json(ws, {
        header: 1,
        defval: null,
        raw: false,
      });

      const processed = processSheetData(json, file.name);
      allNewData.push(...processed);

      setUploadProgress(((i + 1) / newFiles.length) * 100);
      await new Promise((r) => setTimeout(r, 20));
    }

    setParsedData((prev) => [...prev, ...allNewData]);

    setTimeout(() => {
      setIsProcessing(false);
      setUploadProgress(0);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }, 500);
  };

  // --- Rule Evaluation & Sorting Engine ---
  const evaluatedStores = useMemo(() => {
    const parseAsLocal = (dStr) => {
      if (dStr.includes("-") && !dStr.includes("T"))
        return new Date(dStr + "T00:00:00");
      return new Date(dStr);
    };

    let processed = parsedData
      .map((store) => {
        let relevantTimeline = [];
        let baselineRate = null;

        if (startDate || endDate) {
          for (let i = 0; i < store.timeline.length; i++) {
            const pt = store.timeline[i];
            const d = parseAsLocal(pt.date);
            d.setHours(0, 0, 0, 0);

            const startLimit = startDate
              ? parseAsLocal(startDate).setHours(0, 0, 0, 0)
              : null;
            const endLimit = endDate
              ? parseAsLocal(endDate).setHours(0, 0, 0, 0)
              : null;

            if (startLimit && d < startLimit) {
              baselineRate = pt.rate;
              // Removed capturing baselineDate to stop the 1-day offset UI bug
              continue;
            }
            if (endLimit && d > endLimit) {
              continue;
            }
            relevantTimeline.push(pt);
          }

          // Apply LOCF rate to the exact start date requested by the user
          if (
            startDate &&
            baselineRate !== null &&
            relevantTimeline.length > 0
          ) {
            relevantTimeline.unshift({
              date: startDate,
              rate: baselineRate,
              isBaseline: true,
            });
          }
        } else {
          relevantTimeline = [...store.timeline];
        }

        if (relevantTimeline.length === 0) return null;

        let beforeRate = relevantTimeline[0].rate;
        let beforeDate = relevantTimeline[0].date;

        for (let i = 0; i < relevantTimeline.length; i++) {
          if (relevantTimeline[i].rate > beforeRate) {
            beforeRate = relevantTimeline[i].rate;
            beforeDate = relevantTimeline[i].date;
          }
        }

        const afterRate = relevantTimeline[relevantTimeline.length - 1].rate;
        let afterDate = relevantTimeline[relevantTimeline.length - 1].date;

        let foundBefore = false;
        for (let i = 0; i < relevantTimeline.length; i++) {
          if (relevantTimeline[i].date === beforeDate) foundBefore = true;
          if (foundBefore && relevantTimeline[i].rate === afterRate) {
            afterDate = relevantTimeline[i].date;
            break;
          }
        }

        // Evaluate DoorDash Rules
        const B = beforeRate;
        const A = afterRate;
        let isClosedWon = false;
        let ruleMatched = "No Change";
        let statusType = "neutral";

        if (B > A) {
          if (store.type === "Delivery") {
            if (B - A >= 5) {
              isClosedWon = true;
              ruleMatched = "Defilated by 5%+";
              statusType = "won";
            } else if (A === 0) {
              isClosedWon = true;
              ruleMatched = "Defilated to 0%";
              statusType = "won";
            } else if (B > 20 && A === 20) {
              isClosedWon = true;
              ruleMatched = "Defilated to exactly 20%";
              statusType = "won";
            } else {
              ruleMatched = "Natural Defilation (Criteria not met)";
              statusType = "natural";
            }
          } else if (store.type === "Pickup") {
            if (B > 1 && A === 0) {
              isClosedWon = true;
              ruleMatched = "Defilated to 0% (From >1%)";
              statusType = "won";
            } else {
              ruleMatched = "Natural Defilation (Criteria not met)";
              statusType = "natural";
            }
          }
        } else if (A > B) {
          ruleMatched = "Infilated (Prices increased)";
          statusType = "inflated";
        }

        const displayBeforeDate = beforeDate.includes("1999")
          ? "N/A"
          : beforeDate;
        const displayAfterDate = afterDate.includes("2000") ? "N/A" : afterDate;

        return {
          ...store,
          beforeRate: B,
          beforeDate: displayBeforeDate,
          afterRate: A,
          afterDate: displayAfterDate,
          isClosedWon,
          ruleMatched,
          statusType,
        };
      })
      .filter(Boolean);

    if (showOnlyWins) {
      processed = processed.filter((s) => s.isClosedWon);
    }

    // Apply Sorting logic
    processed.sort((a, b) => {
      const getTime = (dStr) => {
        const t = new Date(dStr).getTime();
        return isNaN(t) ? 0 : t;
      };

      const dateA = getTime(a.afterDate);
      const dateB = getTime(b.afterDate);
      const dropA = a.beforeRate - a.afterRate;
      const dropB = b.beforeRate - b.afterRate;

      switch (sortBy) {
        case "date-desc":
          return dateB - dateA;
        case "date-asc":
          return dateA - dateB;
        case "drop-desc":
          return dropB - dropA;
        case "drop-asc":
          return dropA - dropB;
        default:
          return 0;
      }
    });

    return processed;
  }, [parsedData, startDate, endDate, showOnlyWins, sortBy]);

  const totalWins = useMemo(
    () => evaluatedStores.filter((s) => s.isClosedWon).length,
    [evaluatedStores],
  );

  const exportToExcel = () => {
    if (evaluatedStores.length === 0 || !window.XLSX) return;

    const wsData = [
      [
        "Store ID",
        "Type",
        "Date (Before)",
        "Initial Markup",
        "Date (After)",
        "Adjusted Markup",
        "Status",
        "Rule Matched",
        "Source File",
      ],
    ];

    evaluatedStores.forEach((row) => {
      wsData.push([
        row.storeId,
        row.type,
        row.beforeDate,
        `${row.beforeRate}%`,
        row.afterDate,
        `${row.afterRate}%`,
        row.isClosedWon
          ? "Closed Won"
          : row.statusType === "natural"
            ? "Natural Defilation"
            : row.statusType === "inflated"
              ? "Infilated"
              : "No Change",
        row.ruleMatched,
        row.filename,
      ]);
    });

    const ws = window.XLSX.utils.aoa_to_sheet(wsData);
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, "Markup Report");

    let filename = `Markup_Audit_Report_${new Date().toISOString().split("T")[0]}`;
    if (showOnlyWins) filename += "_WINS_ONLY";
    filename += ".xlsx";

    window.XLSX.writeFile(wb, filename);
  };

  const clearData = () => {
    setFiles([]);
    setParsedData([]);
    setStartDate("");
    setEndDate("");
  };
  const clearFilters = () => {
    setStartDate("");
    setEndDate("");
  };
  const onDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };
  const onDragLeave = (e) => {
    e.preventDefault();
    setIsDragging(false);
  };
  const onDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files) handleFileUpload(e.dataTransfer.files);
  };

  if (isLoadingXlsx) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50">
        <Loader2 className="w-8 h-8 text-[#eb1700] animate-spin" />
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900 pb-12 relative">
      <input
        type="file"
        multiple
        accept=".xlsx,.xls,.csv"
        className="hidden"
        ref={fileInputRef}
        onChange={(e) => handleFileUpload(e.target.files)}
      />

      {isProcessing && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center">
          <div className="bg-white p-8 rounded-2xl shadow-xl w-full max-w-md">
            <h3 className="text-lg font-bold text-slate-800 mb-2 flex items-center gap-2">
              <Loader2 className="w-5 h-5 text-[#eb1700] animate-spin" />{" "}
              Processing Files...
            </h3>
            <p className="text-sm text-slate-500 mb-6">
              Enforcing LOCF and timelines...
            </p>
            <div className="w-full bg-slate-100 rounded-full h-3 mb-2 overflow-hidden">
              <div
                className="bg-[#eb1700] h-3 rounded-full transition-all duration-300"
                style={{ width: `${uploadProgress}%` }}
              ></div>
            </div>
            <div className="text-right text-xs font-bold text-slate-400">
              {Math.round(uploadProgress)}% Complete
            </div>
          </div>
        </div>
      )}

      <header className="bg-white border-b border-slate-200 px-8 py-4 flex items-center justify-between sticky top-0 z-10 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="bg-[#eb1700] p-2.5 rounded-xl shadow-sm">
            <TrendingDown className="text-white w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-800 leading-tight">
              MSM Markup Tracker
            </h1>
            <p className="text-xs text-slate-500 font-medium">
              Rule-Based Audit & CW Engine
            </p>
          </div>
        </div>

        {parsedData.length > 0 && (
          <div className="flex gap-3">
            <button
              onClick={clearData}
              className="flex items-center gap-2 px-4 py-2 text-sm font-semibold text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors"
            >
              <Trash2 className="w-4 h-4" /> Clear Data
            </button>
            <button
              onClick={exportToExcel}
              disabled={evaluatedStores.length === 0}
              className="flex items-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-[#eb1700] hover:bg-[#d11500] disabled:bg-slate-300 rounded-lg transition-colors shadow-sm"
            >
              <FileSpreadsheet className="w-4 h-4" /> Export Report
            </button>
          </div>
        )}
      </header>

      <main className="max-w-7xl mx-auto px-8 py-8 space-y-8">
        {parsedData.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-16 animate-in fade-in duration-500">
            <div
              onDragOver={onDragOver}
              onDragLeave={onDragLeave}
              onDrop={onDrop}
              className={`w-full max-w-3xl border-2 border-dashed rounded-3xl p-16 text-center transition-all duration-200 bg-white ${isDragging ? "border-[#eb1700] bg-red-50/50 scale-[1.02]" : "border-slate-300 hover:border-[#eb1700] shadow-sm"}`}
            >
              <div className="mx-auto w-20 h-20 mb-6 rounded-full bg-red-50 flex items-center justify-center">
                <UploadCloud className="w-10 h-10 text-[#eb1700]" />
              </div>
              <h3 className="text-2xl font-bold text-slate-800 mb-3">
                Drag and drop your Excel files
              </h3>
              <p className="text-slate-500 mb-8 max-w-md mx-auto leading-relaxed">
                Upload{" "}
                <strong className="text-slate-700">Sigma Export (.xlsx)</strong>{" "}
                files. We map the timeline, execute LOCF for blanks, and apply
                the Playbook CW rules.
              </p>
              <button
                onClick={() => fileInputRef.current?.click()}
                className="px-8 py-3.5 bg-[#eb1700] text-white font-semibold rounded-xl shadow-sm hover:bg-[#d11500] transition-colors"
              >
                Browse Files
              </button>
            </div>
          </div>
        ) : (
          <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center gap-4">
                <div className="bg-blue-50 p-4 rounded-xl text-blue-600">
                  <Store className="w-6 h-6" />
                </div>
                <div>
                  <p className="text-sm font-semibold text-slate-500">
                    Stores Analyzed
                  </p>
                  <p className="text-3xl font-bold text-slate-800">
                    {parsedData.length}
                  </p>
                </div>
              </div>
              <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center gap-4">
                <div className="bg-green-50 p-4 rounded-xl text-green-600">
                  <CheckCircle2 className="w-6 h-6" />
                </div>
                <div>
                  <p className="text-sm font-semibold text-slate-500">
                    Closed Wons Confirmed
                  </p>
                  <p className="text-3xl font-bold text-slate-800">
                    {totalWins}
                  </p>
                </div>
              </div>
              <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center gap-4">
                <div className="bg-amber-50 p-4 rounded-xl text-amber-600">
                  <FileSpreadsheet className="w-6 h-6" />
                </div>
                <div className="w-full">
                  <p className="text-sm font-semibold text-slate-500">
                    Files Processed
                  </p>
                  <div className="flex items-center justify-between">
                    <p className="text-3xl font-bold text-slate-800">
                      {files.length}
                    </p>
                    <button
                      onClick={() => fileInputRef.current?.click()}
                      className="text-sm text-white bg-[#eb1700] hover:bg-[#d11500] px-3 py-1.5 rounded-lg font-medium transition-colors shadow-sm"
                    >
                      + Add More
                    </button>
                  </div>
                </div>
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden flex flex-col">
              <div className="px-6 py-4 border-b border-slate-200 bg-slate-50 flex flex-col xl:flex-row items-start xl:items-center justify-between gap-4">
                <div className="flex items-center gap-4">
                  <div className="flex items-center gap-2">
                    <Filter className="w-5 h-5 text-slate-400" />
                    <h2 className="text-lg font-bold text-slate-800">
                      Timeline Auditor
                    </h2>
                  </div>

                  {/* View Toggle */}
                  <div className="flex bg-slate-200/60 p-1 rounded-lg">
                    <button
                      onClick={() => setShowOnlyWins(false)}
                      className={`px-3 py-1 text-sm font-semibold rounded-md transition-all ${!showOnlyWins ? "bg-white shadow-sm text-slate-800" : "text-slate-500 hover:text-slate-700"}`}
                    >
                      Show All
                    </button>
                    <button
                      onClick={() => setShowOnlyWins(true)}
                      className={`px-3 py-1 text-sm font-semibold rounded-md transition-all ${showOnlyWins ? "bg-white shadow-sm text-green-700" : "text-slate-500 hover:text-slate-700"}`}
                    >
                      Wins Only
                    </button>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-3 w-full xl:w-auto">
                  {/* Sorting Control */}
                  <div className="flex items-center bg-white border border-slate-300 rounded-lg px-2 py-1.5 focus-within:border-[#eb1700] transition-all">
                    <ArrowUpDown className="w-4 h-4 text-slate-400 ml-1 mr-2" />
                    <select
                      value={sortBy}
                      onChange={(e) => setSortBy(e.target.value)}
                      className="text-sm text-slate-700 outline-none bg-transparent cursor-pointer font-medium pr-1"
                    >
                      <option value="date-desc">Newest Date First</option>
                      <option value="date-asc">Oldest Date First</option>
                      <option value="drop-desc">Largest Drop (%) First</option>
                      <option value="drop-asc">Smallest Drop (%) First</option>
                    </select>
                  </div>

                  <div className="w-px h-6 bg-slate-300 hidden sm:block mx-1"></div>

                  <div className="flex items-center bg-white border border-slate-300 rounded-lg px-3 py-1.5 focus-within:border-[#eb1700] transition-all">
                    <Calendar className="w-4 h-4 text-slate-400 mr-2" />
                    <input
                      type="date"
                      value={startDate}
                      onChange={(e) => setStartDate(e.target.value)}
                      className="text-sm text-slate-700 outline-none bg-transparent"
                      title="Filter by Defilation Date"
                    />
                  </div>
                  <span className="text-slate-400 text-sm font-medium">to</span>
                  <div className="flex items-center bg-white border border-slate-300 rounded-lg px-3 py-1.5 focus-within:border-[#eb1700] transition-all">
                    <Calendar className="w-4 h-4 text-slate-400 mr-2" />
                    <input
                      type="date"
                      value={endDate}
                      onChange={(e) => setEndDate(e.target.value)}
                      className="text-sm text-slate-700 outline-none bg-transparent"
                    />
                  </div>
                  {(startDate || endDate) && (
                    <button
                      onClick={clearFilters}
                      className="p-1.5 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-md transition-colors"
                    >
                      <XCircle className="w-5 h-5" />
                    </button>
                  )}
                </div>
              </div>

              <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
                <table className="w-full text-left border-collapse">
                  <thead className="sticky top-0 bg-white shadow-sm z-10">
                    <tr className="bg-white text-slate-500 text-xs uppercase tracking-wider">
                      <th className="px-6 py-4 font-semibold border-b border-slate-200">
                        Store
                      </th>
                      <th className="px-6 py-4 font-semibold border-b border-slate-200 bg-slate-50/50">
                        Initial Markup
                      </th>
                      <th className="px-4 py-4 font-semibold border-b border-slate-200 text-center text-slate-300">
                        <ArrowRight className="w-4 h-4 mx-auto" />
                      </th>
                      <th className="px-6 py-4 font-semibold border-b border-slate-200 bg-slate-50/50">
                        Adjusted Markup
                      </th>
                      <th className="px-6 py-4 font-semibold border-b border-slate-200">
                        Audit Status
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {evaluatedStores.length === 0 ? (
                      <tr>
                        <td
                          colSpan="5"
                          className="px-6 py-20 text-center text-slate-500 bg-slate-50/50"
                        >
                          <CheckCircle2 className="w-16 h-16 mx-auto text-slate-200 mb-4" />
                          <p className="text-xl font-medium text-slate-700">
                            No records found
                          </p>
                        </td>
                      </tr>
                    ) : (
                      evaluatedStores.map((row) => (
                        <tr
                          key={row.id}
                          className="hover:bg-slate-50 transition-colors"
                        >
                          <td className="px-6 py-4">
                            <div className="font-semibold text-slate-800">
                              {row.storeId}
                            </div>
                            <div
                              className={`text-xs font-semibold mt-1 inline-block px-1.5 py-0.5 rounded ${row.type === "Delivery" ? "bg-purple-100 text-purple-700" : "bg-orange-100 text-orange-700"}`}
                            >
                              {row.type}
                            </div>
                          </td>
                          <td className="px-6 py-4 bg-slate-50/30">
                            <div className="text-lg font-bold text-slate-700">
                              {row.beforeRate}%
                            </div>
                            <div className="text-xs text-slate-500 flex items-center gap-1 mt-1">
                              <Calendar className="w-3 h-3" /> Date:{" "}
                              {row.beforeDate}
                            </div>
                          </td>
                          <td className="px-4 py-4 text-center">
                            {row.beforeRate > row.afterRate ? (
                              <TrendingDown
                                className="w-5 h-5 text-green-500 mx-auto"
                                title={`Drop of ${row.beforeRate - row.afterRate}%`}
                              />
                            ) : (
                              <Minus className="w-5 h-5 text-slate-300 mx-auto" />
                            )}
                          </td>
                          <td className="px-6 py-4 bg-slate-50/30">
                            <div className="text-lg font-bold text-slate-700">
                              {row.afterRate}%
                            </div>
                            <div className="text-xs text-slate-500 flex items-center gap-1 mt-1">
                              <Calendar className="w-3 h-3" /> Date:{" "}
                              {row.afterDate}
                            </div>
                          </td>
                          <td className="px-6 py-4">
                            {row.statusType === "won" && (
                              <div className="flex items-start gap-2">
                                <ShieldCheck className="w-5 h-5 text-green-500 mt-0.5 shrink-0" />
                                <div>
                                  <div className="text-sm font-bold text-green-700">
                                    Closed Won
                                  </div>
                                  <div className="text-xs text-green-600/80 font-medium">
                                    {row.ruleMatched}
                                  </div>
                                </div>
                              </div>
                            )}
                            {row.statusType === "natural" && (
                              <div className="flex items-start gap-2">
                                <ShieldAlert className="w-5 h-5 text-amber-500 mt-0.5 shrink-0" />
                                <div>
                                  <div className="text-sm font-bold text-amber-700">
                                    Natural Defilation
                                  </div>
                                  <div className="text-xs text-amber-600/80 font-medium">
                                    {row.ruleMatched}
                                  </div>
                                </div>
                              </div>
                            )}
                            {row.statusType === "inflated" && (
                              <div className="text-sm font-medium text-red-500">
                                {row.ruleMatched}
                              </div>
                            )}
                            {row.statusType === "neutral" && (
                              <div className="text-sm font-medium text-slate-400">
                                {row.ruleMatched}
                              </div>
                            )}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
              <div className="px-6 py-3 border-t border-slate-200 bg-slate-50 text-right">
                <p className="text-xs font-semibold text-slate-500">
                  Showing {evaluatedStores.length} results
                </p>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
