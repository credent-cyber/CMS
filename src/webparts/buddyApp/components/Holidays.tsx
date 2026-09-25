/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-void */
/* eslint-disable @rushstack/no-new-null */
/* eslint-disable no-useless-escape */
/* eslint-disable no-empty */
import * as React from "react";
import { useState, useEffect, useCallback, useMemo, useRef } from "react";
import { LIST_CONFIG } from "../../../config/spListConfig";
import {
  Chip,
  Drawer,
  IconButton,
  Tooltip,
  TextField,
  InputAdornment,
  Select,
  MenuItem,
  SelectChangeEvent,
  CircularProgress,
  Switch,
} from "@mui/material";
import { DataGrid, GridColDef, GridRenderCellParams } from "@mui/x-data-grid";
import DeleteOutlineIcon         from "@mui/icons-material/DeleteOutline";
import EditOutlinedIcon          from "@mui/icons-material/EditOutlined";
import SaveIcon                  from "@mui/icons-material/Save";
import CloseIcon                 from "@mui/icons-material/Close";
import AddIcon                   from "@mui/icons-material/Add";
import SearchIcon                from "@mui/icons-material/Search";
import BeachAccessIcon           from "@mui/icons-material/BeachAccess";
import FileDownloadOutlinedIcon  from "@mui/icons-material/FileDownloadOutlined";
import KeyboardArrowDownIcon     from "@mui/icons-material/KeyboardArrowDown";
import UploadFileOutlinedIcon    from "@mui/icons-material/UploadFileOutlined";
import CheckCircleOutlineIcon    from "@mui/icons-material/CheckCircleOutline";
import ErrorOutlineIcon          from "@mui/icons-material/ErrorOutline";
import WarningAmberIcon          from "@mui/icons-material/WarningAmber";
import CloudUploadOutlinedIcon   from "@mui/icons-material/CloudUploadOutlined";
import * as XLSX                 from "xlsx";
import { sp }                    from "@pnp/sp";
import { useSnackbar }           from "./Snackbar";
import { BuddyLoader }           from "./Buddyloader";
import styles                    from "./Holidays.module.scss";
import type { IBuddyAppProps }   from "./IBuddyAppProps";

const BRAND = {
  darkGreen:    "#004632",
  classicGreen: "#00D264",
  lightGreen:   "#8CFF8C",
  white:        "#FFFFFF",
  lightStone:   "#F5F2ED",
  darkStone:    "#78736E",
} as const;

const LOCATION_COLORS: Record<string, { bg: string; text: string; border: string }> = {
  "Corporate Office": { bg: "#E3F2FD", text: "#1976D2", border: "#1976D2" },
  "R&D":              { bg: "#F3E5F5", text: "#7B1FA2", border: "#7B1FA2" },
  "Unit 1":           { bg: "#FFE0B2", text: "#E65100", border: "#E65100" },
  "Unit 2":           { bg: "#E0F2F1", text: "#00695C", border: "#00695C" },
  "Unit 3":           { bg: "#FCE4EC", text: "#C2185B", border: "#C2185B" },
};

const GLOBAL_OPTION = "Global";

const DRAWER_SX = {
  zIndex: 99999,
  "& .MuiBackdrop-root": { zIndex: 99998 },
  "& .MuiDrawer-paper": {
    top: "0 !important",
    height: "100vh !important",
    maxHeight: "100vh !important",
    zIndex: 100000,
    display: "flex",
    flexDirection: "column" as const,
  },
} as const;

interface Holiday {
  id: number;
  occasion: string;
  date: string;
  location: string;
}

interface HolidayForm {
  occasion: string;
  date: string;
  location: string;
}

interface BulkRow {
  rowNum:   number;
  occasion: string;
  date:     string;
  location: string;
}

type RowStatus = "valid" | "error" | "duplicate" | "skipped";
interface BulkRowResult extends BulkRow {
  status:   RowStatus;
  errors:   string[];
  expanded: BulkRow[];
}

const EMPTY_FORM: HolidayForm = { occasion: "", date: "", location: "" };

const formatDate = (iso?: string): string => {
  if (!iso) return "—";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "—";
  return d.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
};

const toInputDate = (iso?: string): string => {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
};

const exportToCsv = (holidays: Holiday[]): void => {
  const escape = (v: any) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const header = ["Occasion", "Date", "Location"].map(escape).join(",");
  const body   = holidays.map((h) =>
    [h.occasion, formatDate(h.date), h.location || "Global"].map(escape).join(","),
  ).join("\n");
  const blob = new Blob([header + "\n" + body], { type: "text/csv;charset=utf-8;" });
  const a    = document.createElement("a");
  a.href     = URL.createObjectURL(blob);
  a.download = "holidays.csv";
  a.click();
  URL.revokeObjectURL(a.href);
};

const toYMD = (d: Date): string =>
  `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;

const parseExcelDate = (raw: any): string => {
  if (raw === null || raw === undefined || raw === "") return "";

  if (typeof raw === "number" && raw > 0) {
    try {
      const excelEpoch = new Date(1899, 11, 30); 
      const d = new Date(excelEpoch.getTime() + Math.floor(raw) * 86400000);
      if (!Number.isNaN(d.getTime()) && d.getFullYear() > 1900) return toYMD(d);
    } catch (_) {}
  }

  if (raw instanceof Date) {
    if (!Number.isNaN(raw.getTime())) {
      const y = raw.getUTCFullYear();
      const m = String(raw.getUTCMonth() + 1).padStart(2, "0");
      const d = String(raw.getUTCDate()).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }
  }

  const s = String(raw).trim();

  const ddmmyyyy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (ddmmyyyy) {
    const day   = ddmmyyyy[1].padStart(2, "0");
    const month = ddmmyyyy[2].padStart(2, "0");
    const year  = ddmmyyyy[3];
    const d = new Date(Number(year), Number(month) - 1, Number(day));
    if (!Number.isNaN(d.getTime()) && d.getFullYear() > 1900) return `${year}-${month}-${day}`;
  }

  const yyyymmdd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (yyyymmdd) {
    const y = yyyymmdd[1];
    const m = yyyymmdd[2].padStart(2, "0");
    const d = yyyymmdd[3].padStart(2, "0");
    const dt = new Date(Number(y), Number(m) - 1, Number(d));
    if (!Number.isNaN(dt.getTime()) && dt.getFullYear() > 1900) return `${y}-${m}-${d}`;
  }

  const dt = new Date(s);
  if (!Number.isNaN(dt.getTime()) && dt.getFullYear() > 1900) return toYMD(dt);

  return "";
};

const SELECT_SX = {
  width: "100%",
  fontSize: "0.775rem",
  background: "#fff",
  fontFamily: "'Inter', sans-serif",
  "& .MuiOutlinedInput-notchedOutline": { borderColor: "#e5e7eb" },
  "&:hover .MuiOutlinedInput-notchedOutline": { borderColor: BRAND.classicGreen },
  "&.Mui-focused .MuiOutlinedInput-notchedOutline": { borderColor: BRAND.darkGreen, borderWidth: "1.5px" },
};

const GRID_SX = {
  border: "none", width: "100%", height: "100%",
  fontFamily: "'Inter', sans-serif",
  "& .MuiDataGrid-columnHeaders": {
    backgroundColor: BRAND.darkGreen, color: BRAND.white,
    fontSize: "0.72rem", fontWeight: 700,
    minHeight: "40px !important", maxHeight: "40px !important",
  },
  "& .MuiDataGrid-columnHeaderTitle": {
    fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em",
    color: BRAND.white, fontFamily: "'Inter', sans-serif",
  },
  "& .MuiDataGrid-sortIcon":        { color: `${BRAND.classicGreen} !important` },
  "& .MuiDataGrid-menuIconButton":  { color: `${BRAND.white} !important` },
  "& .MuiDataGrid-columnSeparator": { color: "rgba(255,255,255,0.15)" },
  "& .MuiDataGrid-row:hover":       { backgroundColor: `${BRAND.lightStone} !important` },
  "& .MuiDataGrid-cell": {
    fontSize: "0.775rem", padding: "0 0.625rem", borderColor: "#f3f4f6",
    display: "flex", alignItems: "center", fontFamily: "'Inter', sans-serif",
  },
  "& .MuiDataGrid-footerContainer": { borderTop: "1px solid #e5e7eb", minHeight: "42px", maxHeight: "42px" },
  "& .MuiTablePagination-root": { fontSize: "0.725rem", fontFamily: "'Inter', sans-serif" },
  "& .MuiTablePagination-selectLabel, & .MuiTablePagination-displayedRows": { fontSize: "0.725rem" },
};

export const Holidays: React.FC<IBuddyAppProps> = (props) => {
  const containerRef  = useRef<HTMLDivElement>(null);
  const isMountedRef  = useRef(true);
  const fileInputRef  = useRef<HTMLInputElement>(null);
  const locationsRef  = useRef<string[]>([]);
  const holidaysRef   = useRef<Holiday[]>([]);
  const { showSnackbar } = useSnackbar();

  const [holidays,  setHolidays]  = useState<Holiday[]>([]);
  const [locations, setLocations] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [loadError, setLoadError] = useState("");

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editingId,  setEditingId]  = useState<number | null>(null);
  const [form,       setForm]       = useState<HolidayForm>(EMPTY_FORM);
  const [isSaving,   setIsSaving]   = useState(false);

  const [bulkOpen,        setBulkOpen]        = useState(false);
  const [bulkFile,        setBulkFile]        = useState<File | null>(null);
  const [bulkRows,        setBulkRows]        = useState<BulkRowResult[]>([]);
  const [bulkParsed,      setBulkParsed]      = useState(false);
  const [bulkParsing,     setBulkParsing]     = useState(false);
  const [bulkUploading,   setBulkUploading]   = useState(false);
  const [bulkUploadDone,  setBulkUploadDone]  = useState(false);
  const [bulkUploadStats, setBulkUploadStats] = useState({ added: 0, skipped: 0, failed: 0 });

  const [loaderVisible, setLoaderVisible] = useState(false);
  const [loaderMessage, setLoaderMessage] = useState("Processing…");

  const [searchText,       setSearchText]       = useState("");
  const [filterLocation,   setFilterLocation]   = useState("");
  const [showPastHolidays, setShowPastHolidays] = useState(false);
  const [filterYear,       setFilterYear]       = useState<string>(String(new Date().getFullYear()));

  useEffect(() => () => { isMountedRef.current = false; }, []);

  const loadLocations = useCallback(async () => {
    if (!props.context) return;
    try {
      const items = await sp.web.lists
        .getByTitle(LIST_CONFIG.LISTS.Location_Master)
        .items.select("Title").orderBy("Title", true).top(500).get();
      if (isMountedRef.current) {
        const locs = (items || []).map((i: any) => String(i.Title || "").trim()).filter(Boolean);
        locationsRef.current = locs;
        setLocations(locs);
      }
    } catch (e) { console.error("Holidays: loadLocations", e); }
  }, [props.context]);

  const loadHolidays = useCallback(async () => {
    if (!props.context) return;
    setIsLoading(true); setLoadError("");
    setLoaderVisible(true); setLoaderMessage("Loading holidays…");
    try {
      const items = await sp.web.lists
        .getByTitle(LIST_CONFIG.LISTS.Holidays)
        .items.select("ID,Occasion,Date,Location").orderBy("Date", true).top(500).get();
      if (!isMountedRef.current) return;
      const mapped: Holiday[] = (items || []).map((i: any) => ({
        id:       i.ID,
        occasion: i.Occasion || "",
        date:     i.Date     || "",
        location: String(i.Location || "").trim(),
      }));
      holidaysRef.current = mapped;
      setHolidays(mapped);
    } catch (e) {
      console.error("Holidays: loadHolidays", e);
      if (isMountedRef.current) setLoadError("Failed to load holidays. Please retry.");
    } finally {
      if (isMountedRef.current) { setIsLoading(false); setLoaderVisible(false); }
    }
  }, [props.context]);

  useEffect(() => { void loadLocations(); }, [loadLocations]);
  useEffect(() => { void loadHolidays(); }, [loadHolidays]);

  const openAdd  = () => { setEditingId(null); setForm(EMPTY_FORM); setDrawerOpen(true); };
  const openEdit = (h: Holiday) => {
    setEditingId(h.id);
    setForm({ occasion: h.occasion, date: toInputDate(h.date), location: h.location });
    setDrawerOpen(true);
  };
  const closeDrawer = () => { setDrawerOpen(false); setEditingId(null); setForm(EMPTY_FORM); };

  const handleSave = async () => {
    if (!form.occasion.trim() || !form.date) {
      showSnackbar("Occasion and Date are required.", "error", 3000); return;
    }
    setIsSaving(true); setLoaderVisible(true);
    setLoaderMessage(editingId !== null ? "Updating holiday…" : "Adding holiday…");
    try {
      const normalizedOccasion           = form.occasion.trim().replace(/\s+/g, " ");
      const normalizedOccasionForCompare = form.occasion.trim().replace(/\s+/g, "").toLowerCase();
      const isoDate  = new Date(form.date).toISOString();
      const formDate = form.date;

      if (form.location) {
        const isDup = holidays.some(
          (h) =>
            h.occasion.trim().replace(/\s+/g, "").toLowerCase() === normalizedOccasionForCompare &&
            h.location === form.location &&
            toInputDate(h.date) === formDate &&
            h.id !== editingId,
        );
        if (isDup) {
          showSnackbar(
            `"${normalizedOccasion}" already exists in "${form.location}" on ${formatDate(isoDate)}.`,
            "error", 4000,
          );
          if (isMountedRef.current) { setIsSaving(false); setLoaderVisible(false); }
          return;
        }
      } else if (editingId === null) {
        const isDup = locations.some((loc) =>
          holidays.some(
            (h) =>
              h.occasion.trim().replace(/\s+/g, "").toLowerCase() === normalizedOccasionForCompare &&
              h.location === loc &&
              toInputDate(h.date) === formDate,
          ),
        );
        if (isDup) {
          showSnackbar(
            `"${normalizedOccasion}" already exists in one or more locations on ${formatDate(isoDate)}.`,
            "error", 4000,
          );
          if (isMountedRef.current) { setIsSaving(false); setLoaderVisible(false); }
          return;
        }
      }

      if (!form.location) {
        if (locations.length === 0) {
          showSnackbar("No locations available.", "error", 3000);
          if (isMountedRef.current) setIsSaving(false);
          return;
        }
        if (editingId !== null) {
          await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.Holidays).items.getById(editingId)
            .update({ Occasion: normalizedOccasion, Date: isoDate, Location: locations[0] });
          showSnackbar("Holiday updated.", "success", 3000);
        } else {
          for (const loc of locations) {
            await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.Holidays).items
              .add({ Occasion: normalizedOccasion, Date: isoDate, Location: loc });
          }
          showSnackbar(`Holiday added for all ${locations.length} location(s).`, "success", 3000);
        }
      } else {
        const payload = { Occasion: normalizedOccasion, Date: isoDate, Location: form.location };
        if (editingId !== null) {
          await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.Holidays).items.getById(editingId).update(payload);
          showSnackbar("Holiday updated.", "success", 3000);
        } else {
          await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.Holidays).items.add(payload);
          showSnackbar("Holiday added.", "success", 3000);
        }
      }
      if (!isMountedRef.current) return;
      closeDrawer(); await loadHolidays();
    } catch (e) {
      console.error("Holidays: save", e);
      showSnackbar("Failed to save. Please try again.", "error", 4000);
    } finally {
      if (isMountedRef.current) { setIsSaving(false); setLoaderVisible(false); }
    }
  };

  const handleDelete = useCallback(async (id: number) => {
    if (!window.confirm("Delete this holiday? This cannot be undone.")) return;
    setLoaderVisible(true); setLoaderMessage("Deleting…");
    try {
      await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.Holidays).items.getById(id).delete();
      showSnackbar("Holiday deleted.", "success", 3000);
      await loadHolidays();
    } catch (e) {
      console.error("Holidays: delete", e);
      showSnackbar("Failed to delete.", "error", 4000);
    } finally {
      if (isMountedRef.current) setLoaderVisible(false);
    }
  }, [loadHolidays, showSnackbar]);

  const closeBulkDrawer = () => setBulkOpen(false);

  const [templateUrl,     setTemplateUrl]     = useState<string>("");
  const [templateLoading, setTemplateLoading] = useState(false);
  const [templateDownloading, setTemplateDownloading] = useState(false);

  const fetchTemplateDownloadUrl = useCallback(async () => {
    if (!props.context) return;
    setTemplateLoading(true);
    try {
      const items = await sp.web.lists
        .getByTitle(LIST_CONFIG.LIBRARIES.HolidaysTemplate)
        .items
        .select("ID,FileLeafRef,FileRef,EncodedAbsUrl")
        .orderBy("ID", false)
        .top(1)
        .get();
      if (isMountedRef.current && items?.length > 0) {
        const fileUrl: string =
          items[0].EncodedAbsUrl ||
          `${props.context.pageContext.web.absoluteUrl}${items[0].FileRef}`;
        setTemplateUrl(fileUrl);
      } else if (isMountedRef.current) {
        setTemplateUrl("");
      }
    } catch (e) {
      console.error("Holidays: fetchTemplate", e);
      if (isMountedRef.current) setTemplateUrl("");
    } finally {
      if (isMountedRef.current) setTemplateLoading(false);
    }
  }, [props.context]);

  const handleTemplateDownload = useCallback(async () => {
    if (!templateUrl) return;
    setTemplateDownloading(true);
    try {
      const response = await fetch(templateUrl);
      if (!response.ok) throw new Error("Failed to download template");
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "HolidayTemplate.xlsx";
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (e) {
      console.error("Holidays: handleTemplateDownload", e);
      showSnackbar("Failed to download template. Please try again.", "error", 3000);
    } finally {
      if (isMountedRef.current) setTemplateDownloading(false);
    }
  }, [templateUrl, showSnackbar]);

  const openBulkDrawer = () => {
    setBulkFile(null); setBulkRows([]); setBulkParsed(false); setBulkParsing(false);
    setBulkUploading(false); setBulkUploadDone(false);
    setBulkUploadStats({ added: 0, skipped: 0, failed: 0 });
    setBulkOpen(true);
    void fetchTemplateDownloadUrl();
  };

  const archiveUploadedFile = useCallback(async (file: File): Promise<void> => {
    if (!props.context || !file) return;
    try {
      const ts   = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
      const ext  = file.name.split(".").pop() ?? "xlsx";
      const base = file.name.replace(/\.[^/.]+$/, "");
      const rand = Math.floor(1000 + Math.random() * 9000);
      const name = `${base}_${ts}_${rand}.${ext}`;
      const arrayBuffer = await file.arrayBuffer();
      await sp.web.lists
        .getByTitle(LIST_CONFIG.LIBRARIES.UploadedHolidaysTemplate)
        .rootFolder.files
        .add(name, arrayBuffer, true);
    } catch (e) {
      console.warn("Holidays: archiveUploadedFile failed (non-critical)", e);
    }
  }, [props.context]);

  const parseExcelFile = useCallback((file: File) => {
    setBulkParsing(true);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data  = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb    = XLSX.read(data, { type: "array", cellDates: true });
        const ws    = wb.Sheets[wb.SheetNames[0]];
        const raw   = XLSX.utils.sheet_to_json<any>(ws, { header: 1, defval: "", raw: false });

        let headerIdx = -1;
        for (let i = 0; i < Math.min(raw.length, 10); i++) {
          const row = (raw[i] as any[]).map((c: any) => String(c).trim().toLowerCase());
          if (row.some((c) => c === "occasion")) { headerIdx = i; break; }
        }
        if (headerIdx === -1) {
          if (isMountedRef.current) setBulkParsing(false);
          showSnackbar("Could not find header row with 'Occasion' column.", "error", 5000); return;
        }

        const headers     = (raw[headerIdx] as any[]).map((c: any) => String(c).trim().toLowerCase());
        const colOccasion = headers.findIndex((h) => h === "occasion");
        const colDate     = headers.findIndex((h) => h === "date");
        const colLocation = headers.findIndex((h) => ["location", "loc"].includes(h));

        if (colOccasion === -1 || colDate === -1) {
          showSnackbar("Template must have 'Occasion' and 'Date' columns.", "error", 5000);
          if (isMountedRef.current) setBulkParsing(false);
          return;
        }

        const currentLocations = locationsRef.current;
        const currentHolidays  = holidaysRef.current;

        const existingKeys = new Set(
          currentHolidays.map((h) => {
            const dateStr = toInputDate(h.date);
            return `${h.occasion.trim().replace(/\s+/g, "").toLowerCase()}||${(h.location || "").toLowerCase()}||${dateStr}`;
          }),
        );

        const existingOccasionDateKeys = new Set(
          currentHolidays.map((h) => {
            const dateStr = toInputDate(h.date);
            return `${h.occasion.trim().replace(/\s+/g, "").toLowerCase()}||${dateStr}`;
          }),
        );

        const batchKeys             = new Set<string>();
        const batchOccasionDateKeys = new Set<string>();
        const results: BulkRowResult[] = [];

        const SKIP_PATTERNS = [
          /^required[\s·\-]/i,
          /^valid\s+location/i,
          /^date\s+format/i,
          /^do\s+not\s+edit/i,
          /^fill\s+in/i,
          /^optional/i,
          /^e\.g\./i,
        ];
        const isMetaRow = (occ: string) => SKIP_PATTERNS.some((re) => re.test(occ));

        for (let i = headerIdx + 1; i < raw.length; i++) {
          const rowArr   = raw[i] as any[];
          const occasion = String(rowArr[colOccasion] ?? "").trim();
          const rawDate  = rowArr[colDate];
          const rawLoc   = colLocation !== -1 ? String(rowArr[colLocation] ?? "").trim() : "";

          if (!occasion && !rawDate && !rawLoc) continue;
          if (isMetaRow(occasion)) continue;
          const rawDateStr = String(rawDate ?? "").trim().toLowerCase();
          if (rawDateStr.startsWith("required") || rawDateStr.startsWith("dd/mm") || rawDateStr === "date") continue;

          const errors: string[] = [];

          if (!occasion) errors.push("Occasion is required.");

          const parsedDate = parseExcelDate(rawDate);
          if (!parsedDate) {
            errors.push("Date is missing or invalid (use DD/MM/YYYY or YYYY-MM-DD).");
          }

          const dynamicValid    = [GLOBAL_OPTION, ...currentLocations];
          const locationVal     = rawLoc || GLOBAL_OPTION;
          const matchedLocation = dynamicValid.find(
            (l) => l.toLowerCase() === locationVal.toLowerCase(),
          );
          if (!matchedLocation) {
            if (currentLocations.length === 0) {
              errors.push(`Locations not yet loaded from server. Please close the drawer, wait a moment, and re-upload.`);
            } else {
              errors.push(`Location "${rawLoc}" is not in the active location list. Valid values: ${dynamicValid.join(", ")}.`);
            }
          }

          const resolvedLocation = matchedLocation ?? locationVal;
          const isGlobal         = resolvedLocation === GLOBAL_OPTION;
          const targetLocs       = isGlobal ? (currentLocations.length > 0 ? currentLocations : [GLOBAL_OPTION]) : [resolvedLocation];

          const expanded: BulkRow[] = targetLocs.map((loc) => ({
            rowNum: i + 1, occasion, date: parsedDate, location: loc,
          }));

          const dupErrors: string[] = [];
          if (errors.length === 0 && parsedDate) {

            const occasionDateKey = `${occasion.replace(/\s+/g, "").toLowerCase()}||${parsedDate}`;
            if (existingOccasionDateKeys.has(occasionDateKey)) {
              dupErrors.push(
                `"${occasion}" already exists on ${formatDate(new Date(parsedDate).toISOString())}. Same occasion on the same date is not allowed.`,
              );
            } else if (batchOccasionDateKeys.has(occasionDateKey)) {
              dupErrors.push(
                `"${occasion}" on ${formatDate(new Date(parsedDate).toISOString())} is duplicated within this file.`,
              );
            } else {
              for (const loc of targetLocs) {
                const key = `${occasion.replace(/\s+/g, "").toLowerCase()}||${loc.toLowerCase()}||${parsedDate}`;
                if (existingKeys.has(key)) {
                  dupErrors.push(
                    `"${occasion}" already exists for ${loc} on ${formatDate(new Date(parsedDate).toISOString())}.`,
                  );
                } else if (batchKeys.has(key)) {
                  dupErrors.push(
                    `"${occasion}" is duplicated within this file for ${loc} on ${formatDate(new Date(parsedDate).toISOString())}.`,
                  );
                } else {
                  batchKeys.add(key);
                }
              }
              if (dupErrors.length === 0) {
                batchOccasionDateKeys.add(occasionDateKey);
              }
            }
          }

          const allErrors   = [...errors, ...dupErrors];
          const isDuplicate = errors.length === 0 && dupErrors.length > 0;
          const status: RowStatus = allErrors.length > 0
            ? (isDuplicate ? "duplicate" : "error")
            : "valid";

          results.push({
            rowNum: i + 1, occasion, date: parsedDate,
            location: resolvedLocation, errors: allErrors, status, expanded,
          });
        }

        if (results.length === 0) {
          if (isMountedRef.current) setBulkParsing(false);
          showSnackbar("No data rows found in the file.", "error", 4000); return;
        }

        if (isMountedRef.current) { setBulkRows(results); setBulkParsed(true); setBulkParsing(false); }
      } catch (err) {
        console.error("Bulk parse error", err);
        if (isMountedRef.current) setBulkParsing(false);
        showSnackbar("Failed to parse file. Please use the correct template.", "error", 5000);
      }
    };
    reader.readAsArrayBuffer(file);
  }, [showSnackbar]);

  const handleBulkFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBulkFile(file); setBulkParsed(false); setBulkParsing(false); setBulkRows([]);
    setBulkUploadDone(false);
    parseExcelFile(file);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const handleBulkSubmit = async () => {
    const errorRows     = bulkRows.filter((r) => r.status === "error");
    const duplicateRows = bulkRows.filter((r) => r.status === "duplicate");
    const validRows     = bulkRows.filter((r) => r.status === "valid");

    if (errorRows.length > 0) {
      showSnackbar(
        `Fix ${errorRows.length} error row${errorRows.length !== 1 ? "s" : ""} before uploading. Rows with errors are highlighted in red.`,
        "error", 5000,
      );
      return;
    }
    if (validRows.length === 0 && duplicateRows.length > 0) {
      showSnackbar("All rows are duplicates — nothing to upload.", "error", 3000);
      return;
    }
    if (validRows.length === 0) {
      showSnackbar("No valid rows to upload.", "error", 3000); return;
    }

    setBulkUploading(true);
    setLoaderVisible(true); setLoaderMessage("Uploading holidays in bulk…");

    let added = 0; let skipped = 0; let failed = 0;

    for (const row of validRows) {
      for (const exp of row.expanded) {
        try {
          if (!exp.date) { skipped++; continue; }
          await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.Holidays).items.add({
            Occasion: exp.occasion.trim().replace(/\s+/g, " "),
            Date:     new Date(exp.date).toISOString(),
            Location: exp.location,
          });
          added++;
        } catch (e) {
          console.error("Bulk row error", e, exp);
          failed++;
        }
      }
    }

    if (isMountedRef.current) {
      setBulkUploading(false); setLoaderVisible(false);
      setBulkUploadDone(true);
      setBulkUploadStats({ added, skipped, failed });
      await loadHolidays();
      if (bulkFile) void archiveUploadedFile(bulkFile);
      if (failed === 0)
        showSnackbar(`Bulk upload complete — ${added} row(s) added.`, "success", 4000);
      else
        showSnackbar(`Upload done: ${added} added, ${failed} failed.`, "warning" as any, 5000);
    }
  };

  const today = useMemo(() => {
    const d = new Date(); d.setHours(0, 0, 0, 0); return d;
  }, []);

  const yearOptions = useMemo(() => {
    const y = new Date().getFullYear();
    return [y + 1, y, y - 1, y - 2];
  }, []);

  const filtered = useMemo(() => {
    return holidays.filter((h) => {
      const d = new Date(h.date);
      if (!Number.isNaN(d.getTime()) && d.getFullYear() !== parseInt(filterYear)) return false;
      if (!showPastHolidays && !Number.isNaN(d.getTime()) && d < today) return false;
      if (filterLocation) {
        if (filterLocation === "__global__") { if (h.location) return false; }
        else if ((h.location || "").toLowerCase() !== filterLocation.toLowerCase()) return false;
      }
      if (searchText) {
        const q = searchText.toLowerCase();
        return [h.occasion, h.location, formatDate(h.date)].some((f) => f?.toLowerCase().includes(q));
      }
      return true;
    });
  }, [holidays, filterLocation, searchText, showPastHolidays, today, filterYear]);

  const upcomingCount = useMemo(
    () => filtered.filter((h) => { const d = new Date(h.date); return !Number.isNaN(d.getTime()) && d >= today; }).length,
    [filtered, today],
  );

  const bulkSummary = useMemo(() => {
    const valid      = bulkRows.filter((r) => r.status === "valid").length;
    const errors     = bulkRows.filter((r) => r.status === "error").length;
    const duplicates = bulkRows.filter((r) => r.status === "duplicate").length;
    const totalRows  = bulkRows.reduce((acc, r) => acc + (r.status === "valid" ? r.expanded.length : 0), 0);
    return { valid, errors, duplicates, totalRows };
  }, [bulkRows]);

  const columns: GridColDef[] = [
    {
      field: "actions", headerName: "Actions", width: 90, sortable: false,
      renderCell: (p: GridRenderCellParams) => {
        const holiday = p.row as Holiday;
        const d       = new Date(holiday.date);
        const isPast  = !Number.isNaN(d.getTime()) && d < today;
        return (
          <div style={{ display: "flex", gap: "4px", alignItems: "center" }}>
            <Tooltip title={isPast ? "Cannot edit past occasion" : "Edit holiday"}>
              <span>
                <IconButton size="small" onClick={() => openEdit(holiday)} className={styles.iconBtn} disabled={isPast}>
                  <EditOutlinedIcon sx={{ fontSize: "1rem" }} />
                </IconButton>
              </span>
            </Tooltip>
            <Tooltip title={isPast ? "Cannot delete past occasion" : "Delete holiday"}>
              <span>
                <IconButton size="small" onClick={() => void handleDelete(holiday.id)} className={styles.iconBtnDanger} disabled={isPast}>
                  <DeleteOutlineIcon sx={{ fontSize: "1rem" }} />
                </IconButton>
              </span>
            </Tooltip>
          </div>
        );
      },
    },
    { field: "occasion", headerName: "Occasion", flex: 1, minWidth: 140 },
    {
      field: "date", headerName: "Date", width: 130,
      renderCell: (p) => {
        const d    = new Date(p.value as string);
        const past = !Number.isNaN(d.getTime()) && d < today;
        return (
          <span style={{
            fontSize: "0.775rem", color: past ? BRAND.darkStone : "#111827",
            fontWeight: past ? 400 : 600, fontFamily: "'Inter', sans-serif",
          }}>
            {formatDate(p.value as string)}
          </span>
        );
      },
      sortComparator: (a, b) => new Date(a).getTime() - new Date(b).getTime(),
    },
    {
      field: "location", headerName: "Location", width: 140,
      renderCell: (p) => {
        const loc = p.value as string;
        if (loc && LOCATION_COLORS[loc]) {
          const c = LOCATION_COLORS[loc];
          return (
            <Chip label={loc} size="small" variant="outlined"
              sx={{ fontSize: "0.68rem", height: "21px", fontWeight: 600, fontFamily: "'Inter', sans-serif", background: c.bg, color: c.text, border: `1px solid ${c.border}` }} />
          );
        }
        return loc ? (
          <Chip label={loc} size="small" variant="outlined"
            sx={{ fontSize: "0.68rem", height: "21px", fontWeight: 600, fontFamily: "'Inter', sans-serif", background: BRAND.lightGreen, color: BRAND.darkGreen, border: `1px solid ${BRAND.classicGreen}` }} />
        ) : (
          <Chip label="Global" size="small" variant="outlined"
            sx={{ fontSize: "0.68rem", height: "21px", fontWeight: 600, fontFamily: "'Inter', sans-serif", background: BRAND.lightStone, color: BRAND.darkGreen, border: `1px solid ${BRAND.darkStone}` }} />
        );
      },
    },
    {
      field: "upcoming", headerName: "Status", width: 100, sortable: false,
      valueGetter: (params: any) => {
        const d = new Date(params.row.date);
        return Number.isNaN(d.getTime()) ? "" : d >= today ? "Upcoming" : "Past";
      },
      renderCell: (p) => {
        const v  = p.value as string;
        if (!v) return null;
        const up = v === "Upcoming";
        return (
          <Chip label={v} size="small"
            sx={{ fontSize: "0.68rem", height: "21px", fontWeight: 600, fontFamily: "'Inter', sans-serif",
              background: up ? BRAND.classicGreen : BRAND.lightStone,
              color: up ? BRAND.darkGreen : BRAND.darkStone,
              border: `1px solid ${up ? BRAND.darkGreen : BRAND.darkStone}` }} />
        );
      },
    },
  ];

  return (
    <div className={styles.Container} ref={containerRef}>
      {loaderVisible && (
        <div style={{ position: "fixed", inset: 0, zIndex: 999999 }}>
          <BuddyLoader visible={loaderVisible} variant="default" message={loaderMessage} />
        </div>
      )}

      <div className={styles.pageHeading}>
        <div className={styles.pageHeadingLeft}>
          <div className={styles.pageHeadingIcon}>
            <BeachAccessIcon sx={{ fontSize: "1.1rem" }} />
          </div>
          <div>
            <h2 className={styles.pageTitle}>Holiday Management</h2>
            <p className={styles.pageSubtitle}>Manage organisation holidays and location-specific leave dates</p>
          </div>
        </div>
        <div className={styles.pageHeadingRight}>
          <div className={styles.headingStat}>
            <span className={styles.headingStatVal}>{filtered.length}</span>
            <span className={styles.headingStatLbl}>Total</span>
          </div>
          <div className={styles.headingStatDivider} />
          <div className={styles.headingStat}>
            <span className={styles.headingStatVal}>{upcomingCount}</span>
            <span className={styles.headingStatLbl}>Upcoming</span>
          </div>
          <div className={styles.headingStatDivider} />
          <button className={styles.exportBtn} onClick={() => exportToCsv(filtered)} title="Export CSV">
            <FileDownloadOutlinedIcon sx={{ fontSize: "0.85rem" }} />
            Export
          </button>
          <button className={styles.bulkBtn} onClick={openBulkDrawer}>
            <UploadFileOutlinedIcon sx={{ fontSize: "0.9rem" }} />
            Bulk Upload
          </button>
          <button className={styles.addBtn} onClick={openAdd}>
            <AddIcon sx={{ fontSize: "0.9rem" }} />
            Add Holiday
          </button>
        </div>
      </div>

      {loadError && (
        <div className={styles.errorBanner}>
          {loadError}
          <button className={styles.retryBtn} onClick={() => void loadHolidays()}>Retry</button>
        </div>
      )}

      <div className={styles.filterBar}>
        <div className={styles.filterWrap}>
          <Select value={String(filterYear)} onChange={(e: SelectChangeEvent) => setFilterYear(e.target.value)}
            displayEmpty renderValue={(v) => v ? `Year ${v}` : "Select Year"}
            size="small" IconComponent={KeyboardArrowDownIcon} sx={SELECT_SX}>
            {yearOptions.map((year) => (
              <MenuItem key={year} value={String(year)}
                sx={{ fontSize: "0.775rem", fontFamily: "'Inter', sans-serif" }}
                selected={String(year) === filterYear}>
                {year === new Date().getFullYear()
                  ? `${year} (Current)`
                  : year === new Date().getFullYear() + 1
                    ? `${year} (Next)`
                    : year}
              </MenuItem>
            ))}
          </Select>
        </div>

        <div className={styles.filterWrap}>
          <Select value={filterLocation} onChange={(e: SelectChangeEvent) => setFilterLocation(e.target.value)}
            displayEmpty renderValue={(v) => v === "__global__" ? "Global Only" : v || "All Locations"}
            size="small" IconComponent={KeyboardArrowDownIcon} sx={SELECT_SX}>
            <MenuItem value="" sx={{ fontSize: "0.775rem", color: "#9ca3af", fontFamily: "'Inter', sans-serif" }}><em>All Locations</em></MenuItem>
            {locations.map((loc) => (
              <MenuItem key={loc} value={loc} sx={{ fontSize: "0.775rem", fontFamily: "'Inter', sans-serif" }}>📍 {loc}</MenuItem>
            ))}
          </Select>
        </div>

        <TextField placeholder="Search by occasion, date, location…" value={searchText}
          onChange={(e) => setSearchText(e.target.value)} size="small" variant="outlined"
          InputProps={{
            startAdornment: (<InputAdornment position="start"><SearchIcon sx={{ fontSize: "0.95rem", color: BRAND.darkGreen }} /></InputAdornment>),
            style: { fontFamily: "'Inter', sans-serif" },
          }}
          sx={{
            flex: 1, minWidth: 200,
            "& .MuiOutlinedInput-root": {
              fontSize: "0.775rem", fontFamily: "'Inter', sans-serif",
              "& fieldset": { borderColor: "#e5e7eb" },
              "&:hover fieldset": { borderColor: BRAND.classicGreen },
              "&.Mui-focused fieldset": { borderColor: BRAND.darkGreen, borderWidth: "1.5px" },
            },
          }} />

        <div className={styles.toggleWrap}>
          <span className={styles.toggleLabel}>Show Past</span>
          <Switch checked={showPastHolidays} onChange={(e) => setShowPastHolidays(e.target.checked)} size="small"
            sx={{
              "& .MuiSwitch-switchBase.Mui-checked": { color: BRAND.classicGreen },
              "& .MuiSwitch-switchBase.Mui-checked + .MuiSwitch-track": { backgroundColor: BRAND.classicGreen },
            }} />
        </div>
      </div>

      <div className={styles.gridCard}>
        {isLoading && holidays.length === 0 ? (
          <div className={styles.loadingWrap}>
            <CircularProgress size={28} sx={{ color: BRAND.classicGreen }} />
            <span>Loading holidays…</span>
          </div>
        ) : filtered.length === 0 && !isLoading ? (
          <div className={styles.emptyState}>
            <BeachAccessIcon sx={{ fontSize: "2.5rem", color: BRAND.darkStone, mb: 1 }} />
            <div className={styles.emptyStateTitle}>No holidays found</div>
            <div className={styles.emptyStateDesc}>
              {searchText || filterLocation ? "Try adjusting your filters." : "Click \"Add Holiday\" to create the first entry."}
            </div>
          </div>
        ) : (
          <DataGrid rows={filtered} getRowId={(r: Holiday) => r.id} columns={columns}
            loading={isLoading} disableRowSelectionOnClick
            initialState={{ pagination: { paginationModel: { pageSize: 15 } }, sorting: { sortModel: [{ field: "date", sort: "asc" }] } }}
            pageSizeOptions={[10, 15, 25, 50]} rowHeight={40} columnHeaderHeight={40} sx={GRID_SX} />
        )}
      </div>

      <Drawer
        anchor="right"
        open={drawerOpen}
        onClose={closeDrawer}
        ModalProps={{ disablePortal: true, keepMounted: true, container: containerRef.current }}
        sx={{
          ...DRAWER_SX,
          "& .MuiDrawer-paper": {
            ...DRAWER_SX["& .MuiDrawer-paper"],
            width: { xs: "100%", sm: "620px" },
            boxShadow: "-4px 0 28px rgba(0,0,0,0.18)",
          },
        }}>
        <div className={styles.drawerHeader}>
          <div>
            <h2 className={styles.drawerTitle}>{editingId !== null ? "Edit Holiday" : "Add Holiday"}</h2>
            <p className={styles.drawerSub}>
              {editingId !== null ? "Update the holiday details below" : "Fill in the details for the new holiday"}
            </p>
          </div>
          <IconButton onClick={closeDrawer} className={styles.closeIconBtn}>
            <CloseIcon sx={{ fontSize: "1rem" }} />
          </IconButton>
        </div>

        <div className={styles.drawerBody}>
          <div className={styles.editForm}>
            <div className={styles.formRow3}>
              <div className={styles.editField}>
                <label className={styles.editLabel}>Occasion <span className={styles.required}>*</span></label>
                <input type="text" className={styles.editValue}
                  placeholder="e.g. Festival, National Holiday"
                  value={form.occasion} onChange={(e) => setForm((p) => ({ ...p, occasion: e.target.value }))} />
              </div>
              <div className={styles.editField}>
                <label className={styles.editLabel}>Date <span className={styles.required}>*</span></label>
                <input type="date" className={styles.editValue}
                  value={form.date} onChange={(e) => setForm((p) => ({ ...p, date: e.target.value }))}
                  min={toInputDate(new Date().toISOString())} />
              </div>
              <div className={styles.editField}>
                <label className={styles.editLabel}>Location</label>
                <select className={styles.editValue} value={form.location}
                  onChange={(e) => setForm((p) => ({ ...p, location: e.target.value }))}
                  disabled={editingId !== null}>
                  <option value="">🌐 Global</option>
                  {locations.map((loc) => <option key={loc} value={loc}>📍 {loc}</option>)}
                  {form.location && !locations.includes(form.location) && (
                    <option value={form.location}>📍 {form.location}</option>
                  )}
                </select>
              </div>
            </div>

            {(form.occasion || form.date) && (
              <div className={styles.previewCard}>
                <div className={styles.previewCardHeader}>Preview</div>
                <div className={styles.previewRow}>
                  <span className={styles.previewLabel}>Occasion</span>
                  <span className={styles.previewValue}>{form.occasion || "—"}</span>
                </div>
                <div className={styles.previewRow}>
                  <span className={styles.previewLabel}>Date</span>
                  <span className={styles.previewValue}>{form.date ? formatDate(new Date(form.date).toISOString()) : "—"}</span>
                </div>
                <div className={styles.previewRow}>
                  <span className={styles.previewLabel}>Location</span>
                  <span className={styles.previewValue}>{form.location || "Global"}</span>
                </div>
              </div>
            )}
          </div>
        </div>

        <div className={styles.drawerFooter}>
          <button className={styles.cancelBtn} onClick={closeDrawer} disabled={isSaving}>Cancel</button>
          <button className={styles.saveBtn} onClick={() => void handleSave()}
            disabled={isSaving || !form.occasion.trim() || !form.date}>
            <SaveIcon sx={{ fontSize: "0.9rem" }} />
            {isSaving ? "Saving…" : editingId !== null ? "Save Changes" : "Add Holiday"}
          </button>
        </div>
      </Drawer>

      <Drawer
        anchor="right"
        open={bulkOpen}
        onClose={() => { if (!bulkUploading) closeBulkDrawer(); }}
        ModalProps={{ disablePortal: true, keepMounted: true, container: containerRef.current }}
        sx={{
          ...DRAWER_SX,
          "& .MuiDrawer-paper": {
            ...DRAWER_SX["& .MuiDrawer-paper"],
            width: { xs: "100%", sm: "820px" },
            boxShadow: "-4px 0 28px rgba(0,0,0,0.18)",
          },
        }}>

        <div className={styles.drawerHeader}>
          <div>
            <h2 className={styles.drawerTitle}>Bulk Upload Holidays</h2>
            <p className={styles.drawerSub}>Upload the HolidayTemplate Excel file to add multiple holidays at once</p>
          </div>
          <IconButton onClick={closeBulkDrawer} className={styles.closeIconBtn} disabled={bulkUploading}>
            <CloseIcon sx={{ fontSize: "1rem" }} />
          </IconButton>
        </div>

        <div className={styles.drawerBody}>

          <div className={styles.bulkInstructions}>
            <div className={styles.bulkInstructionsHeader}>
              <div className={styles.bulkInstructionsTitle}>
                <UploadFileOutlinedIcon sx={{ fontSize: "0.95rem" }} />
                How to use bulk upload
              </div>
              {templateLoading ? (
                <div className={styles.templateDownloadLoading}>
                  <CircularProgress size={12} sx={{ color: BRAND.classicGreen }} />
                  <span>Loading template…</span>
                </div>
              ) : templateUrl ? (
                <button
                  onClick={() => void handleTemplateDownload()}
                  disabled={templateDownloading}
                  className={styles.templateDownloadBtn}
                  title="Download the latest HolidayTemplate file"
                >
                  <FileDownloadOutlinedIcon sx={{ fontSize: "0.85rem" }} />
                  {templateDownloading ? "Downloading..." : "Download Template"}
                </button>
              ) : (
                <span className={styles.templateDownloadMissing}>Template not found in library</span>
              )}
            </div>
            <ol className={styles.bulkInstructionsList}>
              <li>Download the <strong>HolidayTemplate</strong> file using the button above.</li>
              <li>Fill in the rows: <code>Occasion</code>, <code>Date</code> (DD/MM/YYYY), <code>Location</code>.</li>
              <li>For <strong>Location</strong> use: <code>Global</code> or any active location (e.g. {locations.length > 0 ? locations.slice(0, 3).join(", ") : "Corporate Office, R&D, Unit 1"}{locations.length > 3 ? "…" : ""}).</li>
              <li>The same occasion name on the same date is <strong>not allowed</strong>, even across different locations.</li>
              <li>If any row has an error, the <strong>entire file is blocked</strong> — fix errors and re-upload.</li>
              <li>Save and upload the filled file below.</li>
            </ol>
          </div>

          {!bulkUploadDone && (
            <div className={styles.bulkDropZone}
              onClick={() => fileInputRef.current?.click()}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                const f = e.dataTransfer.files[0];
                if (f) { setBulkFile(f); setBulkParsed(false); setBulkParsing(false); setBulkRows([]); parseExcelFile(f); }
              }}>
              <input ref={fileInputRef} type="file" accept=".xlsx,.xls"
                style={{ display: "none" }} onChange={handleBulkFileChange} />
              <CloudUploadOutlinedIcon sx={{ fontSize: "2rem", color: BRAND.darkGreen, mb: "0.35rem" }} />
              {bulkFile ? (
                <>
                  <div className={styles.bulkFileName}>{bulkFile.name}</div>
                  <div className={styles.bulkDropHint}>Click to replace file</div>
                </>
              ) : (
                <>
                  <div className={styles.bulkDropHint}>Click or drag &amp; drop your Excel file here</div>
                  <div className={styles.bulkDropSub}>.xlsx or .xls · Max 5 MB</div>
                </>
              )}
            </div>
          )}

          {bulkParsing && (
            <div className={styles.bulkParsingWrap}>
              <CircularProgress size={20} sx={{ color: BRAND.classicGreen }} />
              <span>Reading and validating file…</span>
            </div>
          )}

          {bulkParsed && !bulkUploadDone && (
            <>
              <div className={styles.bulkSummaryRow}>
                <Tooltip
                  title={bulkSummary.totalRows !== bulkSummary.valid
                    ? `${bulkSummary.valid} occasion${bulkSummary.valid !== 1 ? "s" : ""} in the file. Global occasions expand to one entry per location — ${bulkSummary.totalRows} total records will be created in SharePoint.`
                    : `${bulkSummary.valid} occasion${bulkSummary.valid !== 1 ? "s" : ""} will be added to SharePoint.`}
                  placement="bottom">
                  <div className={`${styles.bulkPill} ${styles.bulkPillValid}`} style={{ cursor: "default" }}>
                    <CheckCircleOutlineIcon sx={{ fontSize: "0.85rem" }} />
                    {bulkSummary.valid} occasion{bulkSummary.valid !== 1 ? "s" : ""}
                    {bulkSummary.totalRows !== bulkSummary.valid
                      ? ` → ${bulkSummary.totalRows} SP entries`
                      : ""}
                  </div>
                </Tooltip>
                {bulkSummary.duplicates > 0 && (
                  <div className={`${styles.bulkPill} ${styles.bulkPillDuplicate}`}>
                    <WarningAmberIcon sx={{ fontSize: "0.85rem" }} />
                    {bulkSummary.duplicates} duplicate{bulkSummary.duplicates !== 1 ? "s" : ""}
                  </div>
                )}
                {bulkSummary.errors > 0 && (
                  <div className={`${styles.bulkPill} ${styles.bulkPillError}`}>
                    <ErrorOutlineIcon sx={{ fontSize: "0.85rem" }} />
                    {bulkSummary.errors} error{bulkSummary.errors !== 1 ? "s" : ""}
                  </div>
                )}
              </div>

              <div className={styles.bulkTable}>
                <div className={styles.bulkTableHeader}>
                  <span className={styles.bulkColRow}>#</span>
                  <span className={styles.bulkColOccasion}>Occasion</span>
                  <span className={styles.bulkColDate}>Date</span>
                  <span className={styles.bulkColLoc}>Location</span>
                  <span className={styles.bulkColStatus}>Status</span>
                </div>
                <div className={styles.bulkTableBody}>
                  {bulkRows.map((r, idx) => (
                    <div key={idx} className={`${styles.bulkTableRow} ${styles[`bulkRow_${r.status}`]}`}>
                      <span className={styles.bulkColRow}>{idx + 1}</span>
                      <span className={styles.bulkColOccasion} title={r.occasion}>{r.occasion || "—"}</span>
                      <span className={styles.bulkColDate}>{r.date ? formatDate(new Date(r.date).toISOString()) : "—"}</span>
                      <span className={styles.bulkColLoc}>
                        {r.location === "Global"
                          ? <span className={styles.bulkGlobalTag}>🌐 Global → {locations.length} loc</span>
                          : r.location || "—"}
                      </span>
                      <span className={styles.bulkColStatus}>
                        {r.status === "valid" && (
                          <span className={styles.bulkStatusValid}><CheckCircleOutlineIcon sx={{ fontSize: "0.8rem" }} /> Valid</span>
                        )}
                        {r.status === "duplicate" && (
                          <Tooltip title={r.errors.join(" | ")} placement="left">
                            <span className={styles.bulkStatusDuplicate}><WarningAmberIcon sx={{ fontSize: "0.8rem" }} /> Duplicate</span>
                          </Tooltip>
                        )}
                        {r.status === "error" && (
                          <Tooltip title={r.errors.join(" | ")} placement="left">
                            <span className={styles.bulkStatusError}><ErrorOutlineIcon sx={{ fontSize: "0.8rem" }} /> Error</span>
                          </Tooltip>
                        )}
                      </span>
                    </div>
                  ))}
                </div>
              </div>

              {(bulkSummary.duplicates > 0 || bulkSummary.errors > 0) && (
                <div className={styles.bulkWarningNote}>
                  <WarningAmberIcon sx={{ fontSize: "0.85rem" }} />
                  {bulkSummary.errors > 0
                    ? <span><strong>Rows with errors block the upload.</strong> Fix the highlighted rows and re-upload the file. Duplicate rows will be skipped automatically.</span>
                    : <span>Duplicate rows will be <strong>skipped</strong> during upload. Only valid rows will be processed.</span>
                  }
                </div>
              )}
            </>
          )}

          {bulkUploadDone && (
            <div className={styles.bulkDoneBanner}>
              <CheckCircleOutlineIcon sx={{ fontSize: "2rem", color: BRAND.classicGreen }} />
              <div className={styles.bulkDoneTitle}>Upload Complete</div>
              <div className={styles.bulkDoneStats}>
                <span className={styles.bulkDoneStat}><strong>{bulkUploadStats.added}</strong> rows added</span>
                {bulkUploadStats.skipped > 0 && <span className={styles.bulkDoneStatWarn}><strong>{bulkUploadStats.skipped}</strong> skipped</span>}
                {bulkUploadStats.failed > 0 && <span className={styles.bulkDoneStatErr}><strong>{bulkUploadStats.failed}</strong> failed</span>}
              </div>
              <div className={styles.bulkDoneArchiveNote}>
                <CheckCircleOutlineIcon sx={{ fontSize: "0.8rem" }} />
                File saved to <strong>UploadedHolidaysTemplate</strong> library for records
              </div>
              <button className={styles.bulkUploadAnotherBtn} onClick={() => {
                setBulkFile(null); setBulkRows([]); setBulkParsed(false); setBulkParsing(false); setBulkUploadDone(false);
              }}>
                Upload Another File
              </button>
            </div>
          )}
        </div>

        <div className={styles.drawerFooter}>
          <button className={styles.cancelBtn} onClick={closeBulkDrawer} disabled={bulkUploading}>
            {bulkUploadDone ? "Close" : "Cancel"}
          </button>
          {!bulkUploadDone && (
            <button
              className={`${styles.saveBtn} ${bulkSummary.errors > 0 ? styles.saveBtnError : ""}`}
              onClick={() => void handleBulkSubmit()}
              disabled={!bulkParsed || bulkParsing || bulkSummary.valid === 0 || bulkSummary.errors > 0 || bulkUploading}>
              {bulkUploading
                ? <><CircularProgress size={14} sx={{ color: BRAND.darkGreen }} /> Uploading…</>
                : bulkParsing
                  ? <><CircularProgress size={14} sx={{ color: BRAND.darkGreen }} /> Validating…</>
                  : bulkSummary.errors > 0
                    ? <><ErrorOutlineIcon sx={{ fontSize: "0.9rem" }} /> Fix {bulkSummary.errors} Error{bulkSummary.errors !== 1 ? "s" : ""} First</>
                    : <><UploadFileOutlinedIcon sx={{ fontSize: "0.9rem" }} /> Upload {bulkSummary.valid > 0 ? `${bulkSummary.valid} Occasion${bulkSummary.valid !== 1 ? "s" : ""}${bulkSummary.totalRows !== bulkSummary.valid ? ` (${bulkSummary.totalRows} SP entries)` : ""}` : ""}</>
              }
            </button>
          )}
        </div>
      </Drawer>
    </div>
  );
};

export default Holidays;