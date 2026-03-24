import React, { useState, useRef, useEffect, useMemo } from "react";
import { Analytics } from "@vercel/analytics/react";
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
  Search,
  CheckSquare,
  Square,
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
  const [sortBy, setSortBy] = useState("date-desc");
  const [searchTerm, setSearchTerm] = useState("");

  // Checkbox Filters (30 Days Win Added)
  const [statusFilters, setStatusFilters] = useState({
    thirtyDays: false,
    won: true,
    natural: true,
    neutral: false,
  });

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

  // --- Core Parsing & NLP Regex Engine ---
  const processSheetData = (json, filename) => {
    let type = "Delivery"; // Default fallback

    // Clean up filename for NLP
    let cleanName = filename.replace(/\.[a-zA-Z0-9]+$/, "");
    cleanName = cleanName.replace(/[\s.,)\]]+$/, "");

    const deliveryRegex = /[\\/|_\-=(.,\s\[]+(del(ivery)?|d)$/i;
    const pickupRegex = /[\\/|_\-=(.,\s\[]+(pick(up)?|p)$/i;

    if (deliveryRegex.test(cleanName)) {
      type = "Delivery";
      cleanName = cleanName.replace(deliveryRegex, "");
    } else if (pickupRegex.test(cleanName)) {
      type = "Pickup";
      cleanName = cleanName.replace(pickupRegex, "");
    } else if (cleanName.toLowerCase().includes("del")) {
      type = "Delivery";
    } else if (cleanName.toLowerCase().includes("pick")) {
      type = "Pickup";
    }

    let sellerName = cleanName
      .replace(/[-_/\\[\]()=|.,]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    if (!sellerName) sellerName = "Unknown Rep";

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
      const pastTimestamp = new Date("1999-01-01T00:00:00").setHours(
        0,
        0,
        0,
        0,
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

        const dTimestamp = (
          defDate.includes("-") && !defDate.includes("T")
            ? new Date(defDate + "T00:00:00")
            : new Date(defDate)
        ).setHours(0, 0, 0, 0);
        rawDefRate = Math.max(0, Math.round(rawDefRate));

        const timeline = [];
        if (rawDefRate > 0) {
          timeline.push({
            date: "1999-01-01",
            timestamp: pastTimestamp,
            rate: rawDefRate,
          });
          timeline.push({ date: defDate, timestamp: dTimestamp, rate: 0 });
        } else {
          timeline.push({
            date: "1999-01-01",
            timestamp: pastTimestamp,
            rate: 0,
          });
          timeline.push({ date: defDate, timestamp: dTimestamp, rate: 0 });
        }

        data.push({
          id: Math.random().toString(36).substr(2, 9),
          filename,
          sellerName,
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
        if (h && isNaN(h) && !isNaN(Date.parse(h))) {
          const ts = (
            h.includes("-") && !h.includes("T")
              ? new Date(h + "T00:00:00")
              : new Date(h)
          ).setHours(0, 0, 0, 0);
          dateCols.push({ index: c, dateStr: h, timestamp: ts });
        }
      }
      dateCols.sort((a, b) => a.timestamp - b.timestamp);

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
            rawVal == null ||
            rawVal === "" ||
            rawVal === "-" ||
            String(rawVal).trim().toLowerCase() === "nan";

          let parsedRate = 0;
          if (isBlank) {
            parsedRate = lastKnownRate !== null ? lastKnownRate : 0;
          } else {
            let val = String(rawVal).replace("%", "").trim();
            parsedRate = parseFloat(val) || 0;
            if (parsedRate > 0 && parsedRate <= 2) parsedRate *= 100;
            lastKnownRate = parsedRate;
          }

          const finalRate = Math.max(0, Math.round(parsedRate));
          return { date: dc.dateStr, timestamp: dc.timestamp, rate: finalRate };
        });

        if (timeline.length > 0) {
          data.push({
            id: Math.random().toString(36).substr(2, 9),
            filename,
            sellerName,
            type,
            storeId: String(storeId),
            timeline,
          });
        }
      }
    }
    return data;
  };

  // --- Base Evaluation Engine (Now with TWO separate engines) ---
  const evaluatedStoresBase = useMemo(() => {
    const thirtyDaysMs = 30 * 24 * 60 * 60 * 1000;
    const startLimit = startDate
      ? new Date(startDate + "T00:00:00").setHours(0, 0, 0, 0)
      : null;
    const endLimit = endDate
      ? new Date(endDate + "T00:00:00").setHours(0, 0, 0, 0)
      : null;

    const searchStart = startLimit;

    return parsedData
      .map((store) => {
        let relevantTimeline = [];
        let baselineRate = null;
        let baselineDate = null;

        // 1. Establish the Timeline based on UI Filters
        if (searchStart || endLimit) {
          for (let i = 0; i < store.timeline.length; i++) {
            const pt = store.timeline[i];
            const t = pt.timestamp;

            if (searchStart && t < searchStart) {
              baselineRate = pt.rate;
              baselineDate = pt.date;
              continue;
            }
            if (endLimit && t > endLimit) {
              continue;
            }
            relevantTimeline.push(pt);
          }

          if (
            searchStart &&
            baselineRate !== null &&
            relevantTimeline.length > 0
          ) {
            relevantTimeline.unshift({
              date: baselineDate,
              timestamp: searchStart - 1,
              rate: baselineRate,
              isBaseline: true,
            });
          }
        } else {
          relevantTimeline = store.timeline;
        }

        if (relevantTimeline.length === 0) return null;

        // =========================================================================
        // ENGINE A: STANDARD CLOSED WON LOGIC (Backward Scan - Option B for Display)
        // =========================================================================
        const finalRate = relevantTimeline[relevantTimeline.length - 1].rate;
        let cwAfterRate = finalRate;

        // We find the exact drop point (Option B) for accurate date filtering/sorting AND display
        let cwDropTimestamp =
          relevantTimeline[relevantTimeline.length - 1].timestamp;
        let cwDropIndex = relevantTimeline.length - 1;

        for (let i = relevantTimeline.length - 1; i >= 0; i--) {
          if (relevantTimeline[i].rate !== finalRate) break;
          cwDropTimestamp = relevantTimeline[i].timestamp;
          cwDropIndex = i;
        }

        // Option B: Grab the exact date the new low rate started for the DISPLAY date
        let cwAfterDateOptionB = relevantTimeline[cwDropIndex].date;

        let cwBeforeRate = relevantTimeline[0].rate;
        let cwBeforeDate = relevantTimeline[0].date;

        // To get the tightest window (Option B) for the Before Date, we look at the peak rate right before the drop.
        // The >= ensures we grab the latest possible date (e.g. Feb 9 instead of Feb 1) if the rate was flat.
        if (cwDropIndex > 0) {
          cwBeforeRate = -1;
          for (let i = 0; i < cwDropIndex; i++) {
            if (relevantTimeline[i].rate >= cwBeforeRate) {
              cwBeforeRate = relevantTimeline[i].rate;
              cwBeforeDate = relevantTimeline[i].date;
            }
          }
        }

        let cwIsClosedWon = false;
        let cwRuleMatched = "No Change";
        let cwStatusType = "neutral";
        const B = cwBeforeRate;
        const A = cwAfterRate;

        if (B > A) {
          if (store.type === "Delivery") {
            if (B - A >= 5) {
              cwIsClosedWon = true;
              cwRuleMatched = "Deflated by 5%+";
              cwStatusType = "won";
            } else if (A === 0) {
              cwIsClosedWon = true;
              cwRuleMatched = "Deflated to 0%";
              cwStatusType = "won";
            } else if (B > 20 && A <= 20) {
              cwIsClosedWon = true;
              cwRuleMatched = "Deflated to <= 20% (From >20%)";
              cwStatusType = "won";
            } else {
              cwRuleMatched = "Natural Deflation (Criteria not met)";
              cwStatusType = "natural";
            }
          } else if (store.type === "Pickup") {
            if (B > 1 && A === 0) {
              cwIsClosedWon = true;
              cwRuleMatched = "Deflated to 0% (From >1%)";
              cwStatusType = "won";
            } else {
              cwRuleMatched = "Natural Deflation (Criteria not met)";
              cwStatusType = "natural";
            }
          }
        } else if (A > B) {
          cwRuleMatched = "No Change";
          cwStatusType = "neutral";
        }

        // =========================================================================
        // ENGINE B: MAINTAINED 30 DAYS LOGIC (Forward Scan - Option C)
        // =========================================================================
        let m30PeakRate = -1;
        let m30PeakDate = null;
        let m30DropRate = null;
        let m30DropTimestamp = null;
        let m30RuleMatched = "No Change";

        // Forward scan to find the VERY FIRST qualifying drop
        for (let i = 0; i < relevantTimeline.length; i++) {
          const current = relevantTimeline[i];

          if (!m30DropTimestamp) {
            if (current.rate >= m30PeakRate) {
              m30PeakRate = current.rate;
              m30PeakDate = current.date;
            }

            let isValidDrop = false;
            if (store.type === "Delivery") {
              if (
                m30PeakRate - current.rate >= 5 ||
                current.rate === 0 ||
                (m30PeakRate > 20 && current.rate <= 20)
              ) {
                isValidDrop = true;
              }
            } else if (store.type === "Pickup") {
              if (m30PeakRate > 1 && current.rate === 0) {
                isValidDrop = true;
              }
            }

            if (isValidDrop) {
              m30DropRate = current.rate;
              m30DropTimestamp = current.timestamp;

              if (store.type === "Delivery") {
                if (m30PeakRate - current.rate >= 5)
                  m30RuleMatched = "Deflated by 5%+";
                else if (current.rate === 0) m30RuleMatched = "Deflated to 0%";
                else m30RuleMatched = "Deflated to <= 20% (From >20%)";
              } else {
                m30RuleMatched = "Deflated to 0% (From >1%)";
              }
              break; // Stop looking! We only care about the FIRST win for 30 Days.
            }
          }
        }

        let m30HeldFor30Days = false;
        let m30AfterDateOptionC = null;

        if (m30DropTimestamp !== null) {
          m30HeldFor30Days = true;
          let foundSufficientData = false;

          // Scan the FULL global timeline to verify 30-day maturity
          for (let i = 0; i < store.timeline.length; i++) {
            const t = store.timeline[i].timestamp;
            if (t > m30DropTimestamp && t <= m30DropTimestamp + thirtyDaysMs) {
              if (store.timeline[i].rate > m30DropRate) {
                m30HeldFor30Days = false; // It didn't hold!
                break;
              }
            }
            if (t >= m30DropTimestamp + thirtyDaysMs) {
              foundSufficientData = true;
            }
          }

          if (!foundSufficientData) {
            const latestT = store.timeline[store.timeline.length - 1].timestamp;
            if (latestT < m30DropTimestamp + thirtyDaysMs) {
              m30HeldFor30Days = false;
            }
          }

          // Option C logic: Grab the very last date in the filtered timeline
          m30AfterDateOptionC =
            relevantTimeline[relevantTimeline.length - 1].date;
        }

        return {
          ...store,
          // Engine A: Standard Data (Using Option B for the display date)
          cwBeforeRate: cwBeforeRate,
          cwBeforeDate:
            cwBeforeDate && cwBeforeDate.includes("1999")
              ? "N/A"
              : cwBeforeDate,
          cwAfterRate: cwAfterRate,
          cwAfterDate:
            cwAfterDateOptionB && cwAfterDateOptionB.includes("2000")
              ? "N/A"
              : cwAfterDateOptionB,
          cwAfterTimestamp: cwDropTimestamp, // Keeps accurate filter logic tied strictly to the exact drop date
          cwIsClosedWon,
          cwRuleMatched,
          cwStatusType,

          // Engine B: Maintained 30 Days Data
          m30BeforeRate: m30PeakRate,
          m30BeforeDate:
            m30PeakDate && m30PeakDate.includes("1999") ? "N/A" : m30PeakDate,
          m30AfterRate: m30DropRate,
          m30AfterDate:
            m30AfterDateOptionC && m30AfterDateOptionC.includes("2000")
              ? "N/A"
              : m30AfterDateOptionC,
          m30Timestamp: m30DropTimestamp,
          m30HeldFor30Days,
          m30RuleMatched,
        };
      })
      .filter(Boolean);
  }, [parsedData, startDate, endDate]);

  // Total Wins KPI Calculation (Uses Standard Engine A)
  const totalWins = useMemo(() => {
    const startLimit = startDate
      ? new Date(startDate + "T00:00:00").setHours(0, 0, 0, 0)
      : null;
    const endLimit = endDate
      ? new Date(endDate + "T00:00:00").setHours(0, 0, 0, 0)
      : null;

    return evaluatedStoresBase.filter((s) => {
      const dropDateValid =
        (!startLimit || s.cwAfterTimestamp >= startLimit) &&
        (!endLimit || s.cwAfterTimestamp <= endLimit);
      return s.cwIsClosedWon && dropDateValid;
    }).length;
  }, [evaluatedStoresBase, startDate, endDate]);

  const handleFileUpload = async (uploadedFiles) => {
    if (!window.XLSX)
      return alert("Excel engine is still loading. Please wait a second.");

    setIsProcessing(true);
    setUploadProgress(0);

    const newFiles = Array.from(uploadedFiles);
    setFiles((prev) => [...prev, ...newFiles]);
    let allNewData = [];

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

      allNewData = allNewData.concat(processed);

      setUploadProgress(((i + 1) / newFiles.length) * 100);
      await new Promise((r) => setTimeout(r, 20));
    }

    setParsedData((prev) => prev.concat(allNewData));

    setTimeout(() => {
      setIsProcessing(false);
      setUploadProgress(0);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }, 500);
  };

  // --- Display Engine (Combines Engines based on Checkboxes) ---
  const displayStores = useMemo(() => {
    const startLimit = startDate
      ? new Date(startDate + "T00:00:00").setHours(0, 0, 0, 0)
      : null;
    const endLimit = endDate
      ? new Date(endDate + "T00:00:00").setHours(0, 0, 0, 0)
      : null;

    let processed = evaluatedStoresBase
      .map((s) => {
        let activeStatus = null;
        let displayRule = "";
        let displayBeforeRate,
          displayBeforeDate,
          displayAfterRate,
          displayAfterDate;

        // Date Check Filter logic
        let cwDropDateValid =
          (!startLimit || s.cwAfterTimestamp >= startLimit) &&
          (!endLimit || s.cwAfterTimestamp <= endLimit);
        let m30DropDateValid =
          (!startLimit || s.m30Timestamp >= startLimit) &&
          (!endLimit || s.m30Timestamp <= endLimit);

        // Cascade Priority Logic
        if (
          statusFilters.thirtyDays &&
          s.m30HeldFor30Days &&
          m30DropDateValid
        ) {
          activeStatus = "thirtyDays";
          displayRule = "Maintained 30+ Days: " + s.m30RuleMatched;
          displayBeforeRate = s.m30BeforeRate;
          displayBeforeDate = s.m30BeforeDate;
          displayAfterRate = s.m30AfterRate;
          displayAfterDate = s.m30AfterDate;
        } else if (statusFilters.won && s.cwIsClosedWon && cwDropDateValid) {
          activeStatus = "won";
          displayRule = s.cwRuleMatched;
          displayBeforeRate = s.cwBeforeRate;
          displayBeforeDate = s.cwBeforeDate;
          displayAfterRate = s.cwAfterRate;
          displayAfterDate = s.cwAfterDate;
        } else if (
          statusFilters.natural &&
          s.cwStatusType === "natural" &&
          cwDropDateValid
        ) {
          activeStatus = "natural";
          displayRule = s.cwRuleMatched;
          displayBeforeRate = s.cwBeforeRate;
          displayBeforeDate = s.cwBeforeDate;
          displayAfterRate = s.cwAfterRate;
          displayAfterDate = s.cwAfterDate;
        } else if (
          statusFilters.neutral &&
          s.cwStatusType === "neutral" &&
          cwDropDateValid
        ) {
          activeStatus = "neutral";
          displayRule = s.cwRuleMatched;
          displayBeforeRate = s.cwBeforeRate;
          displayBeforeDate = s.cwBeforeDate;
          displayAfterRate = s.cwAfterRate;
          displayAfterDate = s.cwAfterDate;
        }

        if (!activeStatus) return null;
        return {
          ...s,
          activeStatus,
          ruleMatched: displayRule,
          displayBeforeRate,
          displayBeforeDate,
          displayAfterRate,
          displayAfterDate,
          sortTimestamp:
            activeStatus === "thirtyDays" ? s.m30Timestamp : s.cwAfterTimestamp,
        };
      })
      .filter(Boolean);

    // Apply Search Filter
    if (searchTerm) {
      processed = processed.filter((s) =>
        s.storeId.toLowerCase().includes(searchTerm.toLowerCase()),
      );
    }

    // Apply Sorting logic
    processed.sort((a, b) => {
      const dateA = a.sortTimestamp || 0;
      const dateB = b.sortTimestamp || 0;
      const dropA = a.displayBeforeRate - a.displayAfterRate;
      const dropB = b.displayBeforeRate - b.displayAfterRate;

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
  }, [
    evaluatedStoresBase,
    statusFilters,
    sortBy,
    searchTerm,
    startDate,
    endDate,
  ]);

  // --- Export Excel Logic ---
  const exportToExcel = () => {
    if (displayStores.length === 0 || !window.XLSX) return;

    const wsData = [
      [
        "Which MSM team are you on?",
        "Seller Name / Source",
        "Store ID",
        "Type",
        "Date (Before)",
        "Initial Markup",
        "Date (After)",
        "Adjusted Markup",
        "Status",
        "Rule Matched",
        "Source File",
        "Inflation Type",
      ],
    ];

    displayStores.forEach((row) => {
      wsData.push([
        "Concentrix",
        row.sellerName,
        row.storeId,
        row.type,
        row.displayBeforeDate,
        `${row.displayBeforeRate}%`,
        row.displayAfterDate,
        `${row.displayAfterRate}%`,
        row.activeStatus === "thirtyDays"
          ? "Maintained 30 Days"
          : row.activeStatus === "won"
            ? "Closed Won"
            : row.activeStatus === "natural"
              ? "Natural Deflation"
              : "No Change",
        row.ruleMatched,
        row.filename,
        row.type,
      ]);
    });

    const ws = window.XLSX.utils.aoa_to_sheet(wsData);
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, "Markup Report");

    let filename = `Markup_Audit_Report_${new Date().toISOString().split("T")[0]}`;
    filename += ".xlsx";

    window.XLSX.writeFile(wb, filename);
  };

  const clearData = () => {
    setFiles([]);
    setParsedData([]);
    setStartDate("");
    setEndDate("");
    setSearchTerm("");
  };

  const toggleAllFilters = () => {
    const allChecked =
      statusFilters.thirtyDays &&
      statusFilters.won &&
      statusFilters.natural &&
      statusFilters.neutral;
    setStatusFilters({
      thirtyDays: !allChecked,
      won: !allChecked,
      natural: !allChecked,
      neutral: !allChecked,
    });
  };

  const clearFilters = () => {
    setStartDate("");
    setEndDate("");
    setSearchTerm("");
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
              Parsing names, enforcing LOCF, and calculating rules...
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
              disabled={displayStores.length === 0}
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
                files. Name your file using your name and type (e.g.{" "}
                <span className="font-mono text-xs bg-slate-100 px-1 py-0.5 rounded">
                  Mohamed Zeitoun-d.xlsx
                </span>
                ).
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
            {/* KPI Cards */}
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
                    Total Valid Wins
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

            {/* Checkbox Filter Ribbon */}
            <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-4 flex flex-wrap items-center gap-4">
              <button
                onClick={toggleAllFilters}
                className="text-sm font-semibold text-slate-500 hover:text-slate-800 transition-colors flex items-center gap-1.5 border-r border-slate-200 pr-4"
              >
                {statusFilters.thirtyDays &&
                statusFilters.won &&
                statusFilters.natural &&
                statusFilters.neutral ? (
                  <CheckSquare className="w-4 h-4 text-slate-400" />
                ) : (
                  <Square className="w-4 h-4 text-slate-400" />
                )}
                Toggle All
              </button>

              <label className="flex items-center gap-2 cursor-pointer text-sm font-semibold text-blue-700 bg-blue-50 px-3 py-1.5 rounded-lg hover:bg-blue-100 transition-colors">
                <input
                  type="checkbox"
                  checked={statusFilters.thirtyDays}
                  onChange={(e) =>
                    setStatusFilters({
                      ...statusFilters,
                      thirtyDays: e.target.checked,
                    })
                  }
                  className="rounded text-blue-600 focus:ring-blue-500 w-4 h-4"
                />
                Maintained 30 Days
              </label>
              <label className="flex items-center gap-2 cursor-pointer text-sm font-semibold text-green-700 bg-green-50 px-3 py-1.5 rounded-lg hover:bg-green-100 transition-colors">
                <input
                  type="checkbox"
                  checked={statusFilters.won}
                  onChange={(e) =>
                    setStatusFilters({
                      ...statusFilters,
                      won: e.target.checked,
                    })
                  }
                  className="rounded text-green-600 focus:ring-green-500 w-4 h-4"
                />
                Closed Won
              </label>
              <label className="flex items-center gap-2 cursor-pointer text-sm font-semibold text-amber-700 bg-amber-50 px-3 py-1.5 rounded-lg hover:bg-amber-100 transition-colors">
                <input
                  type="checkbox"
                  checked={statusFilters.natural}
                  onChange={(e) =>
                    setStatusFilters({
                      ...statusFilters,
                      natural: e.target.checked,
                    })
                  }
                  className="rounded text-amber-600 focus:ring-amber-500 w-4 h-4"
                />
                Natural Deflation
              </label>
              <label className="flex items-center gap-2 cursor-pointer text-sm font-semibold text-slate-700 bg-slate-100 px-3 py-1.5 rounded-lg hover:bg-slate-200 transition-colors">
                <input
                  type="checkbox"
                  checked={statusFilters.neutral}
                  onChange={(e) =>
                    setStatusFilters({
                      ...statusFilters,
                      neutral: e.target.checked,
                    })
                  }
                  className="rounded text-slate-600 focus:ring-slate-500 w-4 h-4"
                />
                No Change
              </label>
            </div>

            {/* Main Table Container */}
            <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden flex flex-col">
              <div className="px-6 py-4 border-b border-slate-200 bg-slate-50 flex flex-col xl:flex-row items-start xl:items-center justify-between gap-4">
                <div className="flex flex-wrap items-center gap-4">
                  <div className="flex items-center gap-2">
                    <Filter className="w-5 h-5 text-slate-400" />
                    <h2 className="text-lg font-bold text-slate-800">
                      Timeline Auditor
                    </h2>
                  </div>

                  {/* Search Bar */}
                  <div className="flex items-center bg-white border border-slate-300 rounded-lg px-3 py-1.5 focus-within:border-[#eb1700] focus-within:ring-1 focus-within:ring-[#eb1700] transition-all">
                    <Search className="w-4 h-4 text-slate-400 mr-2" />
                    <input
                      type="text"
                      placeholder="Search Store ID..."
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      className="text-sm text-slate-700 outline-none bg-transparent w-36"
                    />
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-3 w-full xl:w-auto mt-2 xl:mt-0">
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
                      title="Filter by Deflation Date"
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
                  {(startDate || endDate || searchTerm) && (
                    <button
                      onClick={clearFilters}
                      className="p-1.5 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-md transition-colors"
                      title="Clear Filters"
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
                    {displayStores.length === 0 ? (
                      <tr>
                        <td
                          colSpan="5"
                          className="px-6 py-20 text-center text-slate-500 bg-slate-50/50"
                        >
                          <CheckCircle2 className="w-16 h-16 mx-auto text-slate-200 mb-4" />
                          <p className="text-xl font-medium text-slate-700">
                            No records match filters
                          </p>
                        </td>
                      </tr>
                    ) : (
                      displayStores.map((row) => (
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
                            <div
                              className="text-[10px] text-slate-400 mt-1 uppercase tracking-wider truncate max-w-[120px]"
                              title={row.sellerName}
                            >
                              {row.sellerName}
                            </div>
                          </td>
                          <td className="px-6 py-4 bg-slate-50/30">
                            <div className="text-lg font-bold text-slate-700">
                              {row.displayBeforeRate}%
                            </div>
                            <div className="text-xs text-slate-500 flex items-center gap-1 mt-1">
                              <Calendar className="w-3 h-3" /> Date:{" "}
                              {row.displayBeforeDate}
                            </div>
                          </td>
                          <td className="px-4 py-4 text-center">
                            {row.displayBeforeRate > row.displayAfterRate ? (
                              <TrendingDown
                                className="w-5 h-5 text-green-500 mx-auto"
                                title={`Drop of ${row.displayBeforeRate - row.displayAfterRate}%`}
                              />
                            ) : (
                              <Minus className="w-5 h-5 text-slate-300 mx-auto" />
                            )}
                          </td>
                          <td className="px-6 py-4 bg-slate-50/30">
                            <div className="text-lg font-bold text-slate-700">
                              {row.displayAfterRate}%
                            </div>
                            <div className="text-xs text-slate-500 flex items-center gap-1 mt-1">
                              <Calendar className="w-3 h-3" /> Date:{" "}
                              {row.displayAfterDate}
                            </div>
                          </td>
                          <td className="px-6 py-4">
                            {row.activeStatus === "thirtyDays" && (
                              <div className="flex items-start gap-2">
                                <ShieldCheck className="w-5 h-5 text-blue-500 mt-0.5 shrink-0" />
                                <div>
                                  <div className="text-sm font-bold text-blue-700">
                                    Maintained 30 Days
                                  </div>
                                  <div className="text-xs text-blue-600/80 font-medium">
                                    {row.ruleMatched}
                                  </div>
                                </div>
                              </div>
                            )}
                            {row.activeStatus === "won" && (
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
                            {row.activeStatus === "natural" && (
                              <div className="flex items-start gap-2">
                                <ShieldAlert className="w-5 h-5 text-amber-500 mt-0.5 shrink-0" />
                                <div>
                                  <div className="text-sm font-bold text-amber-700">
                                    Natural Deflation
                                  </div>
                                  <div className="text-xs text-amber-600/80 font-medium">
                                    {row.ruleMatched}
                                  </div>
                                </div>
                              </div>
                            )}
                            {row.activeStatus === "neutral" && (
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
                  Showing {displayStores.length} results
                </p>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
