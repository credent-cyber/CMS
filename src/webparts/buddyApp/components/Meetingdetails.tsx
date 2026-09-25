/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/self-closing-comp */
/* eslint-disable dot-notation */
/* eslint-disable  @rushstack/no-new-null*/
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-void */
/* eslint-disable eqeqeq */
/* eslint-disable max-lines*/
/* eslint-disable prefer-const */
/* eslint-disable @microsoft/spfx/import-requires-chunk-name */
/* eslint-disable @typescript-eslint/no-use-before-define*/

import * as React from "react";
import { useState, useMemo, useRef, useEffect, useCallback } from "react";
import { DataGrid, GridColDef, GridRenderCellParams } from "@mui/x-data-grid";
import {
  Chip,
  Drawer,
  Select,
  MenuItem,
  SelectChangeEvent,
  Switch,
  IconButton,
  Tooltip,
  InputAdornment,
  TextField,
  CircularProgress,
  // Radio,
  // RadioGroup,
  // FormControlLabel,
  // FormControl,
  // FormLabel,
} from "@mui/material";
import CalendarMonthOutlinedIcon from "@mui/icons-material/CalendarMonthOutlined";
import VisibilityOutlinedIcon from "@mui/icons-material/VisibilityOutlined";
import SyncAltIcon from "@mui/icons-material/SyncAlt";
import PersonOffOutlinedIcon from "@mui/icons-material/PersonOffOutlined";
import FileDownloadOutlinedIcon from "@mui/icons-material/FileDownloadOutlined";
import CloseIcon from "@mui/icons-material/Close";
import SaveIcon from "@mui/icons-material/Save";
import CheckCircleIcon from "@mui/icons-material/CheckCircle";
import RadioButtonUncheckedIcon from "@mui/icons-material/RadioButtonUnchecked";
import HourglassTopIcon from "@mui/icons-material/HourglassTop";
import KeyboardArrowDownIcon from "@mui/icons-material/KeyboardArrowDown";
import WarningAmberIcon from "@mui/icons-material/WarningAmber";
import SearchIcon from "@mui/icons-material/Search";
import SyncIcon from "@mui/icons-material/Sync";
import LockOutlinedIcon from "@mui/icons-material/LockOutlined";
import QuestionAnswerOutlinedIcon from "@mui/icons-material/QuestionAnswerOutlined";
import EditOutlinedIcon from "@mui/icons-material/EditOutlined";
import BookmarkBorderOutlinedIcon from "@mui/icons-material/BookmarkBorderOutlined";
import BookmarkOutlinedIcon from "@mui/icons-material/BookmarkOutlined";
import EventBusyIcon from "@mui/icons-material/EventBusy";
import LocationOnOutlinedIcon from "@mui/icons-material/LocationOnOutlined";
import HistoryOutlinedIcon from "@mui/icons-material/HistoryOutlined";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import ExpandLessIcon from "@mui/icons-material/ExpandLess";
import DescriptionOutlinedIcon from "@mui/icons-material/DescriptionOutlined";
import FeedbackOutlinedIcon from "@mui/icons-material/FeedbackOutlined";
import StarOutlineIcon from "@mui/icons-material/StarOutline";
import ThumbUpOutlinedIcon from "@mui/icons-material/ThumbUpOutlined";
import { sp, Web } from "@pnp/sp/presets/all";
import { useSnackbar } from "./Snackbar";
import { BuddyLoader } from "./Buddyloader";
import LIST_CONFIG, { getThresholdSafeListItems } from "../../../config/spListConfig";
import styles from "./Meetingdetails.module.scss";
import type { IBuddyAppProps } from './IBuddyAppProps';

type SPMeetingStatus = "Pending" | "Completed" | "Skipped" | string;
type EmpStatus = "Active" | "Inactive" | "No";
type OverallStatus = "Assigned" | "Completed" | string;
type DrawerKind = null | "view" | "reassign" | "release" | "feedback";
type FeedbackQuestionType = "yes_no" | "agree_disagree" | "remark" | "rating";

interface BuddyAllocationRow {
  id: string;
  spId: number;
  njGlobalId: string;
  njName: string;
  njEmail: string;
  njDepartment: string;
  njLocation: string;
  doj: string;
  buddyName: string;
  buddyEmail: string;
  buddyDepartment: string;
  buddyGlobalId: string;
  buddyLocation: string;
  allocationCount: number;
  allocatedOn: string;
  status: string;
  overallStatus: OverallStatus;
  currentMeetingStatus: string;
  firstInteractionDate: string;
  secondInteractionDate: string;
  thirdInteractionDate: string;
  fourthInteractionDate: string;
  fifthInteractionDate: string;
  sixthInteractionDate: string;
  firstActualInteractionDate: string;
  secondActualInteractionDate: string;
  thirdActualInteractionDate: string;
  fourthActualInteractionDate: string;
  fifthActualInteractionDate: string;
  sixthActualInteractionDate: string;
  firstMeetingStatus: SPMeetingStatus;
  secondMeetingStatus: SPMeetingStatus;
  thirdMeetingStatus: SPMeetingStatus;
  fourthMeetingStatus: SPMeetingStatus;
  fifthMeetingStatus: SPMeetingStatus;
  sixthMeetingStatus: SPMeetingStatus;
  employeeStatus: EmpStatus;
  njSpId?: number;
  meetingJsonHistory?: string;
  /** JSON string storing NJ feedback per milestone (3rd and 6th meeting) */
  feedbackJsonHistory?: string;
}

interface MeetingSlot {
  no: number;
  label: string;
  scheduledDate: string;
  actualDate: string;
  status: SPMeetingStatus;
}

interface BuddyQuestion {
  id: number;
  meeting: string;
  question: string;
  status: string;
}

interface MeetingAnswer {
  id: number;
  njGlobalId: string;
  meeting: string;
  question: string;
  answer: string;
  njName: string;
  njId?: string;
  status: string;
  answeredBy?: string;
}

interface HolidayRecord {
  date: Date;
  locations: string[];
  occasion: string;
}

interface MeetingHistoryEntry {
  meeting: string;
  action: "date_updated" | "completed" | "draft_saved";
  previousScheduledDate?: string;
  updatedScheduledDate?: string;
  actualDate?: string;
  previousStatus?: string;
  updatedStatus?: string;
  reason?: string;
  updatedBy: string;
  updatedByEmail: string;
  updatedAt: string;
  recalculatedDates?: Record<string, any>;
  locations?: string;
  holidaysConsidered?: number;
}

/** Feedback question definition */
interface FeedbackQuestion {
  id: string;
  milestone: "3rd" | "6th"; // which milestone this question belongs to
  type: FeedbackQuestionType;
  question: string;
  options?: string[]; // for yes_no / agree_disagree
  required: boolean;
  sortOrder: number;
}

/** A single submitted feedback record */
interface FeedbackEntry {
  milestone: "3rd" | "6th";
  submittedAt: string;
  submittedByEmail: string;
  answers: Record<string, string>; // questionId -> answer
  questionSnapshots?: Record<string, {
    question: string;
    type: FeedbackQuestionType;
  }>;
}

const getFeedbackQuestionsForEntry = (
  entry: FeedbackEntry,
  currentQuestions: FeedbackQuestion[],
): FeedbackQuestion[] => {
  const snapshots = entry.questionSnapshots || {};
  const currentById = new Map(
    currentQuestions
      .filter((q) => q.milestone === entry.milestone)
      .map((q) => [q.id, q]),
  );

  return Object.keys(entry.answers || {}).map((id) => {
    const snapshot = snapshots[id];
    if (snapshot) {
      return {
        id,
        milestone: entry.milestone,
        type: snapshot.type || 'remark',
        question: normalizeSingleSpaces(snapshot.question),
        required: false,
        sortOrder: currentById.get(id)?.sortOrder ?? Number.POSITIVE_INFINITY,
      } as FeedbackQuestion;
    }

    const current = currentById.get(id);
    if (current) return current;

    return {
      id,
      milestone: entry.milestone,
      type: 'remark',
      question: id,
      required: false,
      sortOrder: Number.POSITIVE_INFINITY,
    } as FeedbackQuestion;
  }).sort((left, right) =>
    left.sortOrder - right.sortOrder || left.id.localeCompare(right.id, undefined, { numeric: true }),
  );
};

const MEETING_LABELS = [
  "1st Meeting", "2nd Meeting", "3rd Meeting",
  "4th Meeting", "5th Meeting", "6th Meeting",
];

const MEETING_SLOT_KEYS: {
  label: string;
  status: keyof BuddyAllocationRow;
  scheduled: keyof BuddyAllocationRow;
  actual: keyof BuddyAllocationRow;
  spStatusField: string;
  spScheduledField: string;
  spActualField: string;
}[] = [
  { label: "1st Meeting", status: "firstMeetingStatus",  scheduled: "firstInteractionDate",  actual: "firstActualInteractionDate",  spStatusField: "FirstMeetingStatus",  spScheduledField: "FirstInteractionDate",  spActualField: "FirstActualInteractionDate"  },
  { label: "2nd Meeting", status: "secondMeetingStatus", scheduled: "secondInteractionDate", actual: "secondActualInteractionDate", spStatusField: "SecondMeetingStatus", spScheduledField: "SecondInteractionDate", spActualField: "SecondActualInteractionDate" },
  { label: "3rd Meeting", status: "thirdMeetingStatus",  scheduled: "thirdInteractionDate",  actual: "thirdActualInteractionDate",  spStatusField: "ThirdMeetingStatus",  spScheduledField: "ThirdInteractionDate",  spActualField: "ThirdActualInteractionDate"  },
  { label: "4th Meeting", status: "fourthMeetingStatus", scheduled: "fourthInteractionDate", actual: "fourthActualInteractionDate", spStatusField: "FourthMeetingStatus", spScheduledField: "FourthInteractionDate", spActualField: "FourthActualInteractionDate" },
  { label: "5th Meeting", status: "fifthMeetingStatus",  scheduled: "fifthInteractionDate",  actual: "fifthActualInteractionDate",  spStatusField: "FifthMeetingStatus",  spScheduledField: "FifthInteractionDate",  spActualField: "FifthActualInteractionDate"  },
  { label: "6th Meeting", status: "sixthMeetingStatus",  scheduled: "sixthInteractionDate",  actual: "sixthActualInteractionDate",  spStatusField: "SixthMeetingStatus",  spScheduledField: "SixthInteractionDate",  spActualField: "SixthActualInteractionDate"  },
];

/**
 * Parse feedback options from SharePoint.
 * Supports JSON array format: ["Yes","No"]
 * Also supports comma-separated fallback: Yes,No
 */
const parseFeedbackOptions = (raw?: string): string[] | undefined => {
  if (!raw || !raw.trim()) return undefined;

  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      return parsed.map((x) => String(x));
    }
  } catch {
    // Fallback to comma-separated values below.
  }

  return raw
    .split(/[|,]/)
    .map((x) => x.trim())
    .filter(Boolean);
};

const normalizeFeedbackQuestionType = (raw?: string): FeedbackQuestionType => {
  const v = String(raw || "").trim().toLowerCase().replace(/[\s-]/g, "_");
  if (["yes_no", "yesno", "yes/no"].includes(v)) return "yes_no";
  if (["agree_disagree", "likert", "agree/disagree"].includes(v)) return "agree_disagree";
  if (["rating", "star_rating"].includes(v)) return "rating";
  if (["remark", "remarks", "comment", "text"].includes(v)) return "remark";
  return "remark";
};

const normalizeSingleSpaces = (value?: string): string =>
  String(value || "").trim().replace(/\s+/g, " ");

const formatDate = (v?: string): string => {
  if (!v) return "—";
  const d = new Date(v);
  return Number.isNaN(d.getTime())
    ? "—"
    : d.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
};

const formatExportDate = (v?: string): string => {
  if (!v) return "";
  const d = new Date(v);
  return Number.isNaN(d.getTime())
    ? ""
    : d.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
};

const toDateInputValue = (iso?: string): string => {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  const yyyy = d.getFullYear();
  const mm   = String(d.getMonth() + 1).padStart(2, "0");
  const dd   = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
};

const MIN_MEETING_ANSWER_LENGTH = 20;

const getMeetingAnswerValidationMessage = (text?: string): string => {
  const normalized = normalizeSingleSpaces(text);
  if (!normalized) return "";
  if (normalized.length < MIN_MEETING_ANSWER_LENGTH) {
    return `Minimum ${MIN_MEETING_ANSWER_LENGTH} characters required (${normalized.length}/${MIN_MEETING_ANSWER_LENGTH}).`;
  }
  return "";
};

const todayInputValue = (): string => {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
};

const isWeekend = (dateStr: string): boolean => {
  const d = new Date(dateStr);
  if (Number.isNaN(d.getTime())) return false;
  const day = d.getDay();
  return day === 0 || day === 6;
};

const getHolidayMatch = (
  dateStr: string,
  holidays: HolidayRecord[],
  relevantLocations: string[],
): string | null => {
  if (!dateStr || holidays.length === 0) return null;
  const d = new Date(dateStr);
  d.setHours(0, 0, 0, 0);
  const dateString = d.toDateString();
  const lowerLocs  = relevantLocations.map((l) => l.toLowerCase().trim()).filter(Boolean);
  for (const h of holidays) {
    const hd = new Date(h.date);
    hd.setHours(0, 0, 0, 0);
    if (hd.toDateString() !== dateString) continue;
    const isGlobal = h.locations.length === 0;
    const matches  = isGlobal || h.locations.some((hl) => lowerLocs.includes(hl.toLowerCase().trim()));
    if (matches) return h.occasion || "Holiday";
  }
  return null;
};

const parseSpDate = (raw: string): Date | null => {
  if (!raw) return null;
  const isoStr = raw.includes(";") ? raw.split(";")[1].trim() : raw;
  const d = new Date(isoStr);
  if (Number.isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
};
// kxjngvkdfjgk
// const addWorkingDays = (
//   startDate: Date | string,
//   workingDays: number,
//   holidays: HolidayRecord[],
//   relevantLocations: string[],
// ): Date => {
//   const toMidnight = (d: Date): Date => {
//     const n = new Date(d); n.setHours(0, 0, 0, 0); return n;
//   };
//   const lowerLocs = relevantLocations.map((l) => l.toLowerCase().trim()).filter(Boolean);
//   const holidaySet = new Set<string>();
//   holidays.forEach((h) => {
//     const isGlobal   = h.locations.length === 0;
//     const matchesLoc = isGlobal || h.locations.some((hl) => lowerLocs.includes(hl.toLowerCase().trim()));
//     if (matchesLoc) holidaySet.add(toMidnight(h.date).toDateString());
//   });
//   let current = toMidnight(new Date(startDate));
//   let added   = 0;
//   while (added < workingDays) {
//     current.setDate(current.getDate() + 1);
//     const day = current.getDay();
//     if (day !== 0 && day !== 6 && !holidaySet.has(current.toDateString())) added++;
//   }
//   return current;
// };

/**
 * PROPER EXCEL EXPORT — structured with styled headers and multi-level rows.
 * Excludes Q&A answers from NJ exports (handled by caller by passing empty answers).
 */
const exportToExcel = async (
  filename: string,
  visibleRows: BuddyAllocationRow[],
  allCompletedAnswers: { njGlobalId: string; meeting: string; question: string; answer: string }[],
  includeAnswers = true,
  includeFeedback = false,
  feedbackQuestions: FeedbackQuestion[] = [],
): Promise<void> => {
  const XLSX = await import("xlsx");

  type SheetRow = Record<string, string | number>;
  const workbook = XLSX.utils.book_new();

  // ─── SHEET 1: Detail (Meeting Information) ────────────────────────────────
  const detailData: SheetRow[] = [];
  visibleRows.forEach((row) => {
    MEETING_SLOT_KEYS.forEach((slot) => {
      const meetingStatus = String(row[slot.status] || "Pending");
      const scheduledDate = row[slot.scheduled] ? formatExportDate(row[slot.scheduled] as string) : "";
      const actualDate = row[slot.actual] ? formatExportDate(row[slot.actual] as string) : "";

      detailData.push({
        "NJ Name": row.njName,
        "NJ Email": row.njName,
        "NJ Global ID": row.njGlobalId,
        "NJ Department": row.njDepartment,
        "NJ Location": row.njLocation,
        "DOJ": formatExportDate(row.doj),
        "Buddy Name": row.buddyName,
        "Buddy Email": row.buddyEmail,
        "Buddy Global ID": row.buddyGlobalId,
        "Buddy Department": row.buddyDepartment,
        "Buddy Location": row.buddyLocation,
        "Allocated On": formatExportDate(row.allocatedOn),
        "Status": row.overallStatus,
        "Meeting": slot.label,
        "Meeting Status": meetingStatus,
        "Scheduled Date": scheduledDate,
        "Actual Date": actualDate,
      });
    });
  });
  const detailSheet = XLSX.utils.json_to_sheet(detailData);
  XLSX.utils.book_append_sheet(workbook, detailSheet, "Detail");

  // ─── SHEET 2: Buddy Questions & Answers ──────────────────────────────────
  if (includeAnswers) {
    const qaData: SheetRow[] = [];
    const answerMap = new Map<string, { question: string; answer: string }[]>();

    allCompletedAnswers.forEach((a) => {
      const key = `${a.njGlobalId}__${a.meeting}`;
      const list = answerMap.get(key) ?? [];
      list.push({ question: a.question, answer: a.answer });
      answerMap.set(key, list);
    });

    visibleRows.forEach((row) => {
      MEETING_SLOT_KEYS.forEach((slot) => {
        const meetingStatus = String(row[slot.status] || "Pending");
        const key = `${row.njGlobalId}__${slot.label}`;
        const answers = answerMap.get(key) ?? [];

        if (meetingStatus === "Completed" && answers.length > 0) {
          answers.forEach((qa) => {
            qaData.push({
              "NJ Name": row.njName,
              "NJ Email": row.njName,
              "NJ Global ID": row.njGlobalId,
              "Buddy Name": row.buddyName,
              "Buddy Email": row.buddyEmail,
              "Buddy Global ID": row.buddyGlobalId,
              "Meeting": slot.label,
              "Question": qa.question,
              "Answer": qa.answer,
            });
          });
        }
      });
    });

    if (qaData.length > 0) {
      const qaSheet = XLSX.utils.json_to_sheet(qaData);
      XLSX.utils.book_append_sheet(workbook, qaSheet, "Buddy Questions");
    }
  }

  // ─── SHEET 3: Feedback ────────────────────────────────────────────────────
  if (includeFeedback) {
    const feedbackData: SheetRow[] = [];
    const questionMap = new Map<string, string>();
    feedbackQuestions.forEach((q) => {
      questionMap.set(q.id, q.question);
    });

    visibleRows.forEach((row) => {
      const feedbackEntries = parseFeedbackHistory(row.feedbackJsonHistory);

      feedbackEntries.forEach((entry) => {
        const milestone = entry.milestone === "3rd" ? "3rd Milestone" : "6th Milestone";
        const relatedQuestions = getFeedbackQuestionsForEntry(entry, feedbackQuestions);

            relatedQuestions.forEach((q) => {
              const answer = entry.answers[q.id] || "";
              feedbackData.push({
                "NJ Name": row.njName,
                "NJ Email": row.njName,
                "NJ Global ID": row.njGlobalId,
                "Buddy Name": row.buddyName,
                "Buddy Email": row.buddyEmail,
                "Buddy Global ID": row.buddyGlobalId,
                "Milestone": milestone,
                "Question": q.question,
                "Response": q.type === "rating" ? formatStarRating(answer) : answer,
                "Submitted At": formatExportDate(entry.submittedAt),
                "Submitted By": row.njName,
              });
            });
      });
    });

    if (feedbackData.length > 0) {
      const feedbackSheet = XLSX.utils.json_to_sheet(feedbackData);
      XLSX.utils.book_append_sheet(workbook, feedbackSheet, "Feedback");
    }
  }

  XLSX.writeFile(workbook, filename.endsWith(".xlsx") ? filename : `${filename}.xlsx`);
};

type StatusBadgeKind = "overall" | "current" | "meeting" | "employee";

type StatusBadgeColors = { bg: string; fg: string; border: string };

const getStatusBadgeColors = (status?: string, kind: StatusBadgeKind = "meeting"): StatusBadgeColors => {
  const v = String(status || "").trim().toLowerCase();

  if (!v || v === "—" || v === "-") {
    return { bg: "#f3f4f6", fg: "#6b7280", border: "#e5e7eb" };
  }

  if (kind === "employee") {
    if (v === "active") return { bg: "#dcfce7", fg: "#166534", border: "#86efac" };
    if (["inactive", "no", "released"].includes(v)) return { bg: "#fee2e2", fg: "#991b1b", border: "#fca5a5" };
    return { bg: "#f3f4f6", fg: "#374151", border: "#e5e7eb" };
  }

  if (kind === "overall") {
    if (v === "assigned") return { bg: "#dbeafe", fg: "#1d4ed8", border: "#93c5fd" };
    if (v === "completed") return { bg: "#dcfce7", fg: "#166534", border: "#86efac" };
    if (v.includes("progress")) return { bg: "#fef3c7", fg: "#92400e", border: "#fcd34d" };
    if (["pending", "draft"].includes(v)) return { bg: "#fff7ed", fg: "#c2410c", border: "#fdba74" };
    if (["inactive", "released", "cancelled", "canceled"].includes(v)) return { bg: "#fee2e2", fg: "#991b1b", border: "#fca5a5" };
    return { bg: "#f3f4f6", fg: "#374151", border: "#d1d5db" };
  }

  // Current status contains values like "1st Meeting Pending" / "2nd Meeting Completed".
  if (kind === "current") {
    if (v.includes("completed")) return { bg: "#dcfce7", fg: "#166534", border: "#86efac" };
    if (v.includes("pending") || v.includes("scheduled") || v.includes("upcoming")) return { bg: "#fef3c7", fg: "#92400e", border: "#fcd34d" };
    if (v.includes("skipped")) return { bg: "#e0f2fe", fg: "#075985", border: "#7dd3fc" };
    if (v.includes("inactive") || v.includes("released")) return { bg: "#fee2e2", fg: "#991b1b", border: "#fca5a5" };
    return { bg: "#f3f4f6", fg: "#374151", border: "#d1d5db" };
  }

  // Per-meeting status badges.
  if (v === "completed") return { bg: "#dcfce7", fg: "#166534", border: "#86efac" };
  if (v === "pending") return { bg: "#fef3c7", fg: "#92400e", border: "#fcd34d" };
  if (v === "skipped") return { bg: "#e0f2fe", fg: "#075985", border: "#7dd3fc" };
  if (["inactive", "released", "cancelled", "canceled"].includes(v)) return { bg: "#fee2e2", fg: "#991b1b", border: "#fca5a5" };
  if (v === "draft") return { bg: "#f5f3ff", fg: "#6d28d9", border: "#c4b5fd" };
  return { bg: "#f3f4f6", fg: "#374151", border: "#d1d5db" };
};

const statusBadgeSx = (colors: StatusBadgeColors, fontSize = "0.68rem") => ({
  fontSize,
  height: "21px",
  fontWeight: 700,
  backgroundColor: colors.bg,
  color: colors.fg,
  border: `1px solid ${colors.border}`,
  "& .MuiChip-label": { px: "8px" },
});

const normalizeEmpStatus = (raw?: string): EmpStatus => {
  const v = String(raw || "").toLowerCase().trim();
  return v === "active" ? "Active" : "Inactive";
};

const normalizeGridText = (value?: string | number | null): string =>
  String(value ?? "").trim().toLowerCase();

const isInactiveSearchText = (value: string): boolean =>
  normalizeGridText(value) === "inactive";

const normalizeListStatus = (value?: string): string => {
  const v = normalizeGridText(value);
  if (["inactive", "no", "released", "disabled"].includes(v)) return "Inactive";
  if (v === "active") return "Active";
  return String(value || "").trim();
};

const meetingRowSearchValues = (r: BuddyAllocationRow): Array<string | number | undefined> => [
  r.status,
  r.employeeStatus,
  r.overallStatus,
  r.currentMeetingStatus,
  r.buddyName,
  r.buddyEmail,
  r.buddyLocation,
  r.buddyDepartment,
  r.buddyGlobalId,
  r.njName,
  r.njEmail,
  r.njDepartment,
  r.njLocation,
  r.njGlobalId,
  r.doj,
  r.allocatedOn,
  r.firstMeetingStatus,
  r.secondMeetingStatus,
  r.thirdMeetingStatus,
  r.fourthMeetingStatus,
  r.fifthMeetingStatus,
  r.sixthMeetingStatus,
];

const rowMeetingStatuses = (row: BuddyAllocationRow): SPMeetingStatus[] => [
  row.firstMeetingStatus,
  row.secondMeetingStatus,
  row.thirdMeetingStatus,
  row.fourthMeetingStatus,
  row.fifthMeetingStatus,
  row.sixthMeetingStatus,
];

const completedMeetingCountForRow = (row: BuddyAllocationRow): number =>
  rowMeetingStatuses(row).filter((status) => status === "Completed").length;

const isAllMeetingsCompleted = (row: BuddyAllocationRow): boolean =>
  completedMeetingCountForRow(row) === MEETING_SLOT_KEYS.length || row.overallStatus === "Completed";

const isSameBuddyAsCurrent = (
  row: BuddyAllocationRow,
  buddy?: { id?: number; email?: string; globalId?: string } | null,
): boolean => {
  if (!row || !buddy) return false;
  const selectedGlobal = String(buddy.globalId || "").toLowerCase().trim();
  const currentGlobal  = String(row.buddyGlobalId || "").toLowerCase().trim();
  const selectedEmail  = String(buddy.email || "").toLowerCase().trim();
  const currentEmail   = String(row.buddyEmail || "").toLowerCase().trim();
  return (!!selectedGlobal && selectedGlobal === currentGlobal) || (!!selectedEmail && selectedEmail === currentEmail);
};

const parseMeetingHistory = (jsonStr?: string): MeetingHistoryEntry[] => {
  if (!jsonStr) return [];
  try {
    const parsed = JSON.parse(jsonStr);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch { return []; }
};

const parseFeedbackHistory = (jsonStr?: string): FeedbackEntry[] => {
  if (!jsonStr) return [];
  try {
    const parsed = JSON.parse(jsonStr);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch { return []; }
};

const formatStarRating = (ratingValue: string | number): string => {
  const num = Number(ratingValue);
  if (isNaN(num) || num < 1 || num > 5) {
    return String(ratingValue);
  }
  const fullStars = Math.round(num);
  const emptyStars = 5 - fullStars;
  return "★".repeat(fullStars) + "☆".repeat(emptyStars) + ` (${ratingValue})`;
};

const renderFeedbackRatingStars = (ratingValue: string | number): React.ReactElement => {
  const num = Number(ratingValue);
  if (isNaN(num) || num < 1 || num > 5) {
    return <span>{ratingValue}</span>;
  }
  const fullStars = Math.round(num);
  const emptyStars = 5 - fullStars;
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: "2px" }}>
      <span style={{ color: "#f59e0b" }}>{"★".repeat(fullStars)}</span>
      <span style={{ color: "#d1d5db" }}>{"☆".repeat(emptyStars)}</span>
      <span style={{ marginLeft: "4px", color: "inherit" }}>({ratingValue})</span>
    </span>
  );
};


const canNJSubmitFeedback = (milestone: "3rd" | "6th", row: BuddyAllocationRow): boolean => {
  if (milestone === "3rd") {
    // Can submit if 3rd meeting is Completed
    return row.thirdMeetingStatus === "Completed";
  } else if (milestone === "6th") {
    // Can submit if 6th meeting is Completed
    return row.sixthMeetingStatus === "Completed";
  }
  return false;
};

const hasNJSubmittedFeedback = (milestone: "3rd" | "6th", njEmail: string, row: BuddyAllocationRow): boolean => {
  const entries = parseFeedbackHistory(row.feedbackJsonHistory);
  return entries.some((e) => e.milestone === milestone && e.submittedByEmail === njEmail);
};

/**
 * Check if a buddy is available for reassignment.
 * Rule: Buddy can have max 4 active NJs.
 * When buddy already has 4 active assignments, they can accept a reassignment only if
 * at least ONE existing assignment has ALL 6 meetings COMPLETED (meaning that slot is completely freed up).
 *
 * @param buddyGlobalId - Global ID of the buddy to check
 * @param allRows - All buddy allocation rows
 * @returns true if buddy is available for reassignment, false otherwise
 */
const isBuddyAvailableForReassignment = (
  buddyGlobalId: string,
  allRows: BuddyAllocationRow[],
): boolean => {
  // Count only active allocations that are NOT completed
  // Completed allocations free up slots for new assignments
  const buddyRowsActive = allRows.filter(
    (row) =>
      row.buddyGlobalId === buddyGlobalId &&
      row.employeeStatus === "Active" &&
      row.overallStatus !== "Completed"
  );

  // If less than 4 non-completed active assignments, buddy is available
  return buddyRowsActive.length < 4;
};

function StatusIcon({ status }: { status: SPMeetingStatus }) {
  if (status === "Completed") return <CheckCircleIcon sx={{ fontSize: "1rem", color: "#10b981" }} />;
  if (status === "Pending")   return <HourglassTopIcon sx={{ fontSize: "1rem", color: "#f59e0b" }} />;
  if (status === "Skipped")   return <RadioButtonUncheckedIcon sx={{ fontSize: "1rem", color: "#dc2626" }} />;
  return <RadioButtonUncheckedIcon sx={{ fontSize: "1rem", color: "#d1d5db" }} />;
}

function SChip({ s, kind = "meeting", fontSize = "0.68rem" }: { s?: string; kind?: StatusBadgeKind; fontSize?: string }) {
  const colors = getStatusBadgeColors(s, kind);
  return (
    <Chip
      label={s || "—"}
      size="small"
      sx={statusBadgeSx(colors, fontSize)}
    />
  );
}

const GRID_SX = {
  border: "none", width: "100%", height: "100%",
  "& .MuiDataGrid-columnHeaders": { backgroundColor: "#047857", color: "#fff", fontSize: "0.72rem", fontWeight: 700, minHeight: "40px !important", maxHeight: "40px !important" },
  "& .MuiDataGrid-columnHeaderTitle": { fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: "#fff" },
  "& .MuiDataGrid-sortIcon": { color: "#fff !important" },
  "& .MuiDataGrid-menuIconButton": { color: "#fff !important" },
  "& .MuiDataGrid-columnSeparator": { color: "rgba(255,255,255,0.15)" },
  "& .MuiDataGrid-row": { transition: "background 0.15s" },
  "& .MuiDataGrid-row:hover": { backgroundColor: "#ecfdf5 !important" },
  "& .MuiDataGrid-cell": { fontSize: "0.775rem", padding: "0 0.625rem", borderColor: "#f3f4f6", display: "flex", alignItems: "center" },
  "& .MuiDataGrid-footerContainer": { borderTop: "1px solid #e5e7eb", minHeight: "42px", maxHeight: "42px" },
  "& .MuiTablePagination-root": { fontSize: "0.725rem" },
  "& .MuiTablePagination-selectLabel, & .MuiTablePagination-displayedRows": { fontSize: "0.725rem" },
  "& .MuiDataGrid-virtualScroller": { overflowX: "auto" },
  "& .row-inactive": { backgroundColor: "rgba(209,213,219,0.5)", opacity: 0.75, "&:hover": { opacity: 1, backgroundColor: "rgba(209,213,219,0.7)" } },
};

const FILTER_SELECT_SX = {
  width: "100%", fontSize: "0.775rem", background: "#fff",
  "& .MuiOutlinedInput-notchedOutline":             { borderColor: "#e5e7eb" },
  "&:hover .MuiOutlinedInput-notchedOutline":       { borderColor: "#00A859" },
  "&.Mui-focused .MuiOutlinedInput-notchedOutline": { borderColor: "#00A859" },
  "& .MuiSelect-select": { padding: "6.5px 12px", fontSize: "0.775rem" },
};

const DRAWER_OVER_TOPBAR_SX = {
  zIndex: 99999,
  "& .MuiBackdrop-root": {
    zIndex: 99998,
  },
  "& .MuiDrawer-paper": {
    top: "0 !important",
    height: "100vh !important",
    maxHeight: "100vh !important",
    overflowY: "hidden !important",
    zIndex: 100000,
  },
};

export const MeetingDetails: React.FC<IBuddyAppProps> = (props) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const isMountedRef = useRef(true);
  const { showSnackbar } = useSnackbar();

  // ── Data ───────────────────────────────────────────────────────────────────
  const [rows,      setRows]      = useState<BuddyAllocationRow[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [loadError, setLoadError] = useState("");

  const holidaysRef = useRef<HolidayRecord[]>([]);

  // ── Filters ────────────────────────────────────────────────────────────────
  const [filterBuddy,  setFilterBuddy]  = useState("");
  const [filterLoc,    setFilterLoc]    = useState("");
  const [filterNJ,     setFilterNJ]     = useState("");
  const [filterStatus, setFilterStatus] = useState("");
  const [searchText,   setSearchText]   = useState("");
  const [showInactiveMeetings, setShowInactiveMeetings] = useState(false);

  // ── View drawer ────────────────────────────────────────────────────────────
  const [drawerKind, setDrawerKind] = useState<DrawerKind>(null);
  const [activeRow,  setActiveRow]  = useState<BuddyAllocationRow | null>(null);
  const [selMeeting, setSelMeeting] = useState<string>("Summary");
  const [showHistory, setShowHistory] = useState(false);

  // ── Q&A state ──────────────────────────────────────────────────────────────
  const [allQuestions,     setAllQuestions]     = useState<BuddyQuestion[]>([]);
  const [savedAnswers,     setSavedAnswers]     = useState<MeetingAnswer[]>([]);
  const [draftAnswers,     setDraftAnswers]     = useState<Record<string, string>>({});
  const [isLoadingQA,      setIsLoadingQA]      = useState(false);
  const [isSavingComplete, setIsSavingComplete] = useState(false);
  const [isSavingDraft,    setIsSavingDraft]    = useState(false);

  // ── Date editing state ─────────────────────────────────────────────────────
  const [isEditingDates,    setIsEditingDates]    = useState(false);
  const [editScheduledDate, setEditScheduledDate] = useState("");
  const [isSavingDates,     setIsSavingDates]     = useState(false);
  const [editReason,        setEditReason]        = useState("");
  const [dateError,         setDateError]         = useState("");

  // ── Reassign state ─────────────────────────────────────────────────────────
  const [buddySearchTerm,    setBuddySearchTerm]    = useState("");
  const [buddySearchResults, setBuddySearchResults] = useState<{ id: number; name: string; email: string; globalId: string; location: string; department: string }[]>([]);
  const [showBuddyDropdown,  setShowBuddyDropdown]  = useState(false);
  const [selectedNewBuddy,   setSelectedNewBuddy]   = useState<(typeof buddySearchResults)[0] | null>(null);
  const [isSavingReassign,   setIsSavingReassign]   = useState(false);
  const buddyBlurRef = useRef<number | null>(null);

  // ── Release state ──────────────────────────────────────────────────────────
  const [relActive,   setRelActive]   = useState(true);
  const [relConfirm,  setRelConfirm]  = useState(false);
  const [isSavingRel, setIsSavingRel] = useState(false);

  // ── Feedback state ─────────────────────────────────────────────────────────
  const [feedbackMilestone, setFeedbackMilestone] = useState<"3rd" | "6th">("3rd");
  const [feedbackAnswers,   setFeedbackAnswers]   = useState<Record<string, string>>({});
  const [isSavingFeedback,  setIsSavingFeedback]  = useState(false);
  const [showFeedbackView,  setShowFeedbackView]  = useState(true);
  const [feedbackQuestions, setFeedbackQuestions] = useState<FeedbackQuestion[]>([]);
  const [isLoadingFeedbackQuestions, setIsLoadingFeedbackQuestions] = useState(false);
  const [showStatsLoader, setShowStatsLoader] = useState(true);
  const [isUserBootstrapped, setIsUserBootstrapped] = useState(false);

  // Controlled pagination prevents the grid from staying blank on first render.
  // It also gives us a stable debug point when users click next/previous.
  const [paginationModel, setPaginationModel] = useState({ page: 0, pageSize: 10 });

//   const [showInactive,     setShowInactive]     = useState(false);
  const [currentUserName,  setCurrentUserName]  = useState("");
  const [currentUserEmail, setCurrentUserEmail] = useState("");
  const [currentUserRole,  setCurrentUserRole]  = useState("");
  const [currentUserLocation, setCurrentUserLocation] = useState("");

  // ─── Bootstrap current user ────────────────────────────────────────────────

  useEffect(() => {
    if (!props.context) return;

    const bootstrapCurrentUser = async () => {
      setIsUserBootstrapped(false);
      try {
        const email = props.context?.pageContext?.user?.email ||
          props.context?.pageContext?.legacyPageContext?.userEmail || "";
        const name  = props.context?.pageContext?.user?.displayName || "";

        setCurrentUserEmail(String(email || "").toLowerCase());
        setCurrentUserName(name);

        if (!email) {
          console.warn("[MeetingDetails/bootstrap] Login email not found. Loading without user role scope.");
          return;
        }

        const safeEmail = String(email).replace(/'/g, "''");
        // const items = await sp.web.lists
        //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        const items = await getThresholdSafeListItems({
          web: empMasterWeb,
          listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
          select: "ID,EmailID,Role,FullName,Location",
          filter: `EmailID eq '${safeEmail}' and PositionType eq 1`,
          maxItems: 1,
        });

        if (!isMountedRef.current) return;

        const userItem = items?.[0];
        setCurrentUserRole(userItem?.Role || "");
        if (userItem?.FullName) setCurrentUserName(userItem.FullName);
        setCurrentUserLocation(userItem?.Location || "");

        console.info("[MeetingDetails/bootstrap]", {
          loginEmail: email,
          matchedItems: items?.length || 0,
          role: userItem?.Role || "",
          location: userItem?.Location || "",
          employeeMasterId: userItem?.ID,
        });
      } catch (e) {
        console.error("MeetingDetails: bootstrapCurrentUser", e);
      } finally {
        if (isMountedRef.current) setIsUserBootstrapped(true);
      }
    };
    void bootstrapCurrentUser();
  }, [props.context]);

  // ─── Load NeulandHolidaysList ──────────────────────────────────────────────

  useEffect(() => {
    if (!props.context) return;
    const load = async () => {
      try {
        const items = await getThresholdSafeListItems({
          web: sp.web,
          listTitle: LIST_CONFIG.LISTS.Holidays,
          select: "ID,Title,Occasion,Date,Location",
        });

        const records: HolidayRecord[] = (items || []).reduce(
          (acc: HolidayRecord[], i: any) => {
            const raw = String(i.Date || "").trim();
            if (!raw) return acc;
            const isoStr = raw.includes(";") ? raw.split(";")[1].trim() : raw;
            const d = parseSpDate(isoStr);
            if (!d) { console.warn(`NeulandHolidaysList: bad date "${raw}"`); return acc; }
            let locs: string[] = [];
            if (Array.isArray(i.Location)) {
              locs = i.Location.map((l: any) => String(l || "").trim()).filter(Boolean);
            } else if (i.Location) {
              locs = [String(i.Location).trim()].filter(Boolean);
            }
            acc.push({ date: d, locations: locs, occasion: i.Occasion || i.Title || "Holiday" });
            return acc;
          },
          [],
        );
        holidaysRef.current = records;
      } catch (e) {
        console.warn("MeetingDetails: NeulandHolidaysList not loaded:", e);
        holidaysRef.current = [];
      }
    };
    void load();
  }, [props.context]);

  // ─── Cleanup ───────────────────────────────────────────────────────────────

  useEffect(() => {
    return () => {
      isMountedRef.current = false;
      if (buddyBlurRef.current) window.clearTimeout(buddyBlurRef.current);
    };
  }, []);

  // ─── Load BuddyAllocate rows ───────────────────────────────────────────────

  const loadRows = useCallback(async () => {
    if (!props.context || !isUserBootstrapped) {
      console.debug("[MeetingDetails/loadRows] waiting", {
        hasContext: Boolean(props.context),
        isUserBootstrapped,
      });
      return;
    }

    setIsLoading(true);
    setLoadError("");

    const debugStart = Date.now();

    console.info("[MeetingDetails/loadRows] start", {
      role: currentUserRole || "(blank)",
      email: currentUserEmail || "(blank)",
      location: currentUserLocation || "(blank)",
      showInactiveMeetings,
    });

    try {
      const items = await getThresholdSafeListItems({
        web: sp.web,
        listTitle: LIST_CONFIG.LISTS.BuddyAllocate,
        select:
          "ID,Title,Status,OverallStatus,CurrentMeetingStatus,AllocationCount," +
          "GlobalIDofBuddy,LocationofBuddy,BuddyEmailID,BuddyDepartment,BuddyAllocatedOn," +
          "BuddyNameRefChoiceColumn," +
          "GlobalIDofNJ,NJEmailID," +
          "NameOfTheNewJoinerRefColumn," +
          "DateofJoining," +
          "FirstInteractionDate,SecondInteractionDate,ThirdInteractionDate," +
          "FourthInteractionDate,FifthInteractionDate,SixthInteractionDate," +
          "FirstActualInteractionDate,SecondActualInteractionDate,ThirdActualInteractionDate," +
          "FourthActualInteractionDate,FifthActualInteractionDate,SixthActualInteractionDate," +
          "FirstMeetingStatus,SecondMeetingStatus,ThirdMeetingStatus," +
          "FourthMeetingStatus,FifthMeetingStatus,SixthMeetingStatus," +
          "MetingJsonHistory,FeedbackJsonHistory",
        // Do not filter BuddyAllocate by Status here. Released/inactive allocations
        // must still load so the Show Inactive toggle can display them.
      });

      if (!isMountedRef.current) return;

      const mapped: BuddyAllocationRow[] = (items || []).map((i: any) => ({
        id: i.ID.toString(), spId: i.ID,
        njGlobalId: i.GlobalIDofNJ || "",
        njName: normalizeSingleSpaces(i.NameOfTheNewJoinerRefColumn || ""),
        njEmail: i.NJEmailID || "",
        njDepartment: "", njLocation: "",
        buddyName: normalizeSingleSpaces(i.BuddyNameRefChoiceColumn || ""),
        buddyEmail: i.BuddyEmailID || "",
        buddyDepartment: normalizeSingleSpaces(i.BuddyDepartment || ""),
        buddyGlobalId: i.GlobalIDofBuddy || "",
        buddyLocation: normalizeSingleSpaces(i.LocationofBuddy || ""),
        allocationCount: i.AllocationCount ?? 1,
        allocatedOn: i.BuddyAllocatedOn || "",
        status: normalizeListStatus(i.Status),
        overallStatus: i.OverallStatus || "—",
        currentMeetingStatus: i.CurrentMeetingStatus || "—",
        doj: i.DateofJoining || "",
        firstInteractionDate: i.FirstInteractionDate || "",
        secondInteractionDate: i.SecondInteractionDate || "",
        thirdInteractionDate: i.ThirdInteractionDate || "",
        fourthInteractionDate: i.FourthInteractionDate || "",
        fifthInteractionDate: i.FifthInteractionDate || "",
        sixthInteractionDate: i.SixthInteractionDate || "",
        firstActualInteractionDate: i.FirstActualInteractionDate || "",
        secondActualInteractionDate: i.SecondActualInteractionDate || "",
        thirdActualInteractionDate: i.ThirdActualInteractionDate || "",
        fourthActualInteractionDate: i.FourthActualInteractionDate || "",
        fifthActualInteractionDate: i.FifthActualInteractionDate || "",
        sixthActualInteractionDate: i.SixthActualInteractionDate || "",
        firstMeetingStatus: i.FirstMeetingStatus || "Pending",
        secondMeetingStatus: i.SecondMeetingStatus || "Pending",
        thirdMeetingStatus: i.ThirdMeetingStatus || "Pending",
        fourthMeetingStatus: i.FourthMeetingStatus || "Pending",
        fifthMeetingStatus: i.FifthMeetingStatus || "Pending",
        sixthMeetingStatus: i.SixthMeetingStatus || "Pending",
        employeeStatus: "Active",
        njSpId: undefined,
        meetingJsonHistory: i.MetingJsonHistory || "",
        feedbackJsonHistory: i.FeedbackJsonHistory || "",
      }));

      // Enrich with NJ data
      try {
        // const njItems = await sp.web.lists
        //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        const njItems = await getThresholdSafeListItems({
          web: empMasterWeb,
          listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
          select: "ID,GlobalID,Department,Location,EmployeeStatus,Role",
          filter: "PositionType eq 1",
          // Do not filter only Role eq 'New Joinee' here. Some Employee Master
          // records have blank/null Role, but their allocation row still exists.
          // Loading all light-weight employee rows makes the first page enrichment stable.
        });

        const njMap = new Map<string, { id: number; dept: string; loc: string; status: EmpStatus }>();
        for (const nj of njItems || []) {
          if (nj.GlobalID) njMap.set(nj.GlobalID, { id: nj.ID, dept: nj.Department || "", loc: nj.Location || "", status: normalizeEmpStatus(nj.EmployeeStatus) });
        }
        for (const row of mapped) {
          const info = njMap.get(row.njGlobalId);
          if (info) { row.njDepartment = info.dept; row.njLocation = info.loc; row.employeeStatus = info.status; row.njSpId = info.id; }
        }

        const missingEnrichment = mapped.filter((row) => row.njGlobalId && !njMap.has(row.njGlobalId));
        console.info("[MeetingDetails/loadRows] enrichment", {
          employeeRowsFetched: njItems?.length || 0,
          employeeGlobalIdsMapped: njMap.size,
          allocationRowsMissingEmployeeMasterMatch: missingEnrichment.length,
          sampleMissingGlobalIds: missingEnrichment.slice(0, 10).map((row) => row.njGlobalId),
          allNJGlobalIds: mapped.slice(0, 5).map((row) => row.njGlobalId),
          allMapGlobalIds: Array.from(njMap.keys()).slice(0, 5),
        });
      } catch (enrichErr) {
        console.warn("MeetingDetails: NJ enrichment failed (non-fatal)", enrichErr);
      }

      console.info("[MeetingDetails/loadRows] result", {
        fetchedFromBuddyAllocate: items?.length || 0,
        mappedRows: mapped.length,
        durationMs: Date.now() - debugStart,
        firstRow: mapped[0],
        first10Ids: mapped.slice(0, 10).map((row) => row.id),
        statusSummary: mapped.reduce((acc: Record<string, number>, row) => {
          const key = row.status || "(blank)";
          acc[key] = (acc[key] || 0) + 1;
          return acc;
        }, {}),
      });

      setRows(mapped);
      setPaginationModel((prev) => ({ ...prev, page: 0 }));
    } catch (e) {
      console.error("MeetingDetails: loadRows", e);
      if (isMountedRef.current) setLoadError("Failed to load meeting data. Please retry.");
    } finally {
      if (isMountedRef.current) setIsLoading(false);
    }
  }, [props.context, isUserBootstrapped, currentUserRole, currentUserEmail, currentUserLocation, showInactiveMeetings]);

  useEffect(() => { void loadRows(); }, [loadRows]);

  // ─── Hide stats loader after data loads ────────────────────────────────────
  useEffect(() => {
    if (!isLoading && showStatsLoader) {
      const timer = setTimeout(() => setShowStatsLoader(false), 600);
      return () => clearTimeout(timer);
    }
  }, [isLoading, showStatsLoader]);

  // ─── Load Questions ────────────────────────────────────────────────────────

  const loadQuestions = useCallback(async () => {
    try {
      const items = await getThresholdSafeListItems({
        web: sp.web,
        listTitle: LIST_CONFIG.LISTS.BUDDY_QUESTIONS,
        select: "ID,Meeting,Questions,Status",
        filter: "Status eq 'Active'",
      });
      if (!isMountedRef.current) return;
      const mappedQuestions: BuddyQuestion[] = (items || []).map((i: any) => ({
        id: Number(i.ID),
        meeting: i.Meeting || "",
        question: i.Questions || "",
        status: i.Status || "",
      }));
      setAllQuestions(mappedQuestions.sort((left, right) => left.id - right.id));
    } catch (e) { console.error("MeetingDetails: loadQuestions", e); }
  }, []);

  useEffect(() => { void loadQuestions(); }, [loadQuestions]);

  // ─── Load Feedback Questions Master ─────────────────────────────────────────

  const loadFeedbackQuestions = useCallback(async () => {
    try {
      setIsLoadingFeedbackQuestions(true);

      const items = await getThresholdSafeListItems({
        web: sp.web,
        listTitle: LIST_CONFIG.LISTS.FEEDBACK_QUESTIONS_MASTER,
        select: "ID,Title,QuestionId,Milestone,QuestionType,QuestionText,Options,IsRequired,SortOrder,Status",
        filter: "Status eq 'Active'",
        sort: (a: any, b: any) =>
          Number(a.SortOrder || 0) - Number(b.SortOrder || 0) ||
          Number(a.ID || 0) - Number(b.ID || 0),
      });

      if (!isMountedRef.current) return;

      const mapped: FeedbackQuestion[] = (items || [])
        .map((i: any) => ({
          id: i.QuestionId || i.Title || String(i.ID),
          milestone: (i.Milestone === "6th" ? "6th" : "3rd") as "3rd" | "6th",
          type: normalizeFeedbackQuestionType(i.QuestionType),
          question: normalizeSingleSpaces(i.QuestionText),
          options: parseFeedbackOptions(i.Options)?.map((opt) => normalizeSingleSpaces(opt)).filter(Boolean),
          required: i.IsRequired === "Yes",
          sortOrder: Number(i.SortOrder || 0),
        }))
        .filter((q) => q.id && q.question && q.type) as FeedbackQuestion[];

      setFeedbackQuestions(mapped);
    } catch (e) {
      console.error("MeetingDetails: loadFeedbackQuestions", e);
      showSnackbar("Failed to load data.", "error", 4000);
    } finally {
      if (isMountedRef.current) setIsLoadingFeedbackQuestions(false);
    }
  }, [showSnackbar]);

  useEffect(() => { void loadFeedbackQuestions(); }, [loadFeedbackQuestions]);

  // ─── Load saved answers ────────────────────────────────────────────────────

  const loadAnswersForRow = useCallback(async (njGlobalId: string) => {
    if (!njGlobalId) return;
    setIsLoadingQA(true);
    try {
      const normalizedNjGlobalId = njGlobalId.trim().toLowerCase();
      const items = await getThresholdSafeListItems({
        web: sp.web,
        listTitle: LIST_CONFIG.LISTS.MEETING_ANSWERS,
        select: "ID,GlobalIDofNJ,Meetings,Questions,Answers,Status,Author/Title",
        expand: "Author",
        // Filtering locally avoids threshold failures when GlobalIDofNJ is not indexed.
        includeItem: (item: any) =>
          String(item.GlobalIDofNJ || "").trim().toLowerCase() === normalizedNjGlobalId,
      });
      if (!isMountedRef.current) return;
      setSavedAnswers((items || []).map((i: any) => ({
        id: i.ID, njGlobalId: i.GlobalIDofNJ || "", meeting: i.Meetings || "",
        question: i.Questions || "", answer: i.Answers || "",
        njName: "", njId: "", status: i.Status || "",
        answeredBy: i.Author?.Title || "",
      })));
    } catch (e) { console.error("MeetingDetails: loadAnswers", e); }
    finally { if (isMountedRef.current) setIsLoadingQA(false); }
  }, []);

  // ─── Build history entry ───────────────────────────────────────────────────

  const buildUpdatedHistory = useCallback(
    (row: BuddyAllocationRow, newEntry: MeetingHistoryEntry): string => {
      let history: MeetingHistoryEntry[] = [];
      if (row.meetingJsonHistory) {
        try { history = JSON.parse(row.meetingJsonHistory); } catch { history = []; }
      }
      history.push(newEntry);
      if (history.length > 200) history = history.slice(-200);
      return JSON.stringify(history, null, 2);
    },
    [],
  );

  // ─── Role helpers ──────────────────────────────────────────────────────────

  const isNJLogin      = useMemo(() => currentUserRole === "New Joinee",       [currentUserRole]);
  const isBuddyLogin   = useMemo(() => currentUserRole === "Buddy",            [currentUserRole]);
  const isSuperAdmin   = useMemo(() => currentUserRole === "Super Admin",      [currentUserRole]);
  const isLndAdmin     = useMemo(() => currentUserRole === "L&D Site Admin",   [currentUserRole]);
  const isAdminLogin   = useMemo(() => isSuperAdmin || isLndAdmin,             [isSuperAdmin, isLndAdmin]);

  // ─── Base visible rows ─────────────────────────────────────────────────────
  // First apply only user-role / location scoping. Do NOT apply Show Inactive here,
  // because the header must be able to show Active + Inactive counts together.
  const roleScopedRows = useMemo(
    () => rows.filter((r) => {
      // NJ: only their own row (matched by email)
      if (isNJLogin && (r.njEmail || "").toLowerCase() !== currentUserEmail) return false;

      // Buddy: only their assigned NJs
      if (isBuddyLogin && (r.buddyEmail || "").toLowerCase() !== currentUserEmail) return false;

      // L&D Site Admin: only NJs whose location matches admin's location
      if (isLndAdmin && currentUserLocation) {
        if ((r.njLocation || "").toLowerCase().trim() !== currentUserLocation.toLowerCase().trim()) return false;
      }

      return true;
    }),
    [rows, isNJLogin, isBuddyLogin, currentUserEmail, isLndAdmin, currentUserLocation],
  );

  const searchRequestsInactive = useMemo(
    () => isInactiveSearchText(searchText),
    [searchText],
  );

  // Grid rows respect the Show Inactive toggle. When the user searches exactly
  // "Inactive", keep inactive rows in scope so the search can actually match them.
  const baseVisibleRows = useMemo(
    () => roleScopedRows.filter((r) => {
      const allocationStatus = normalizeGridText(r.status);
      const employeeStatus = normalizeGridText(r.employeeStatus);

      if (
        !showInactiveMeetings &&
        !searchRequestsInactive &&
        (allocationStatus === "inactive" || employeeStatus === "inactive")
      ) {
        return false;
      }

      return true;
    }),
    [roleScopedRows, showInactiveMeetings, searchRequestsInactive],
  );

  const buddyOpts  = useMemo(() => {
    let rows = baseVisibleRows;
    // Filter by location if selected
    if (filterLoc) {
      rows = rows.filter((r) => r.buddyLocation === filterLoc);
    }
    // Filter by NJ name if selected
    if (filterNJ) {
      rows = rows.filter((r) => r.njName === filterNJ);
    }
    // Filter by status if selected
    if (filterStatus) {
      rows = rows.filter((r) => r.overallStatus === filterStatus);
    }
    return [...new Set(rows.map((r) => r.buddyName).filter(Boolean))].sort();
  }, [baseVisibleRows, filterLoc, filterNJ, filterStatus]);
  
  const locOpts    = useMemo(() => {
    let rows = baseVisibleRows;
    // Filter by buddy name if selected
    if (filterBuddy) {
      rows = rows.filter((r) => r.buddyName === filterBuddy);
    }
    // Filter by NJ name if selected
    if (filterNJ) {
      rows = rows.filter((r) => r.njName === filterNJ);
    }
    // Filter by status if selected
    if (filterStatus) {
      rows = rows.filter((r) => r.overallStatus === filterStatus);
    }
    return [...new Set(rows.map((r) => r.buddyLocation).filter(Boolean))].sort();
  }, [baseVisibleRows, filterBuddy, filterNJ, filterStatus]);
  
  const njOpts     = useMemo(() => {
    let rows = baseVisibleRows;
    // Filter by location if selected
    if (filterLoc) {
      rows = rows.filter((r) => r.buddyLocation === filterLoc);
    }
    // Filter by buddy name if selected
    if (filterBuddy) {
      rows = rows.filter((r) => r.buddyName === filterBuddy);
    }
    // Filter by status if selected
    if (filterStatus) {
      rows = rows.filter((r) => r.overallStatus === filterStatus);
    }
    return [...new Set(rows.map((r) => r.njName).filter(Boolean))].sort();
  }, [baseVisibleRows, filterLoc, filterBuddy, filterStatus]);
  
  const statusOpts = useMemo(() => {
    let rows = baseVisibleRows;
    // Filter by location if selected
    if (filterLoc) {
      rows = rows.filter((r) => r.buddyLocation === filterLoc);
    }
    // Filter by buddy name if selected
    if (filterBuddy) {
      rows = rows.filter((r) => r.buddyName === filterBuddy);
    }
    // Filter by NJ name if selected
    if (filterNJ) {
      rows = rows.filter((r) => r.njName === filterNJ);
    }
    return [...new Set(rows.map((r) => r.overallStatus).filter(Boolean))].sort();
  }, [baseVisibleRows, filterLoc, filterBuddy, filterNJ]);

  const buddySelfInfo = useMemo(() => {
    if (!isBuddyLogin) return null;
    const myRow = rows.find((r) => (r.buddyEmail || "").toLowerCase() === currentUserEmail);
    if (!myRow) return null;
    return { name: myRow.buddyName || "", location: myRow.buddyLocation || "" };
  }, [isBuddyLogin, rows, currentUserEmail]);

  const canEditRow = useCallback(
    (row?: BuddyAllocationRow | null) => {
      if (!row || !isBuddyLogin) return false;
      return (row.buddyEmail || "").toLowerCase() === currentUserEmail;
    },
    [isBuddyLogin, currentUserEmail],
  );

  const filtered = useMemo(
    () => baseVisibleRows.filter((r) => {
      if (filterBuddy  && r.buddyName       !== filterBuddy)  return false;
      if (filterLoc    && r.buddyLocation   !== filterLoc)    return false;
      if (filterNJ     && r.njName          !== filterNJ)     return false;
      if (filterStatus && r.overallStatus   !== filterStatus) return false;
      if (searchText.trim()) {
        const q = normalizeGridText(searchText);
        return meetingRowSearchValues(r).some((f) => normalizeGridText(f).includes(q));
      }
      return true;
    }),
    [baseVisibleRows, filterBuddy, filterLoc, filterNJ, filterStatus, searchText],
  );

  useEffect(() => {
    setPaginationModel((prev) => ({ ...prev, page: 0 }));
  }, [filterBuddy, filterLoc, filterNJ, filterStatus, searchText, showInactiveMeetings, currentUserRole, currentUserLocation]);

  // ─── Reset buddy/NJ filters when location changes (cascading filter) ─────
  useEffect(() => {
    if (filterLoc) {
      // When location is selected, reset buddy and NJ to avoid invalid selections
      setFilterBuddy("");
      setFilterNJ("");
    }
  }, [filterLoc]);

  // ─── Reset NJ filter when buddy changes (cascading filter) ─────
  useEffect(() => {
    if (filterBuddy) {
      // When buddy is selected, reset NJ to show only NJs for that buddy
      setFilterNJ("");
    }
  }, [filterBuddy]);

  // ─── Reset buddy filter when NJ changes (cascading filter) ─────
  useEffect(() => {
    if (filterNJ) {
      // When NJ is selected, reset buddy to show only buddy for that NJ
      setFilterBuddy("");
    }
  }, [filterNJ]);

  useEffect(() => {
    console.info("[MeetingDetails/grid]", {
      totalRowsLoaded: rows.length,
      roleScopedRows: roleScopedRows.length,
      baseVisibleRows: baseVisibleRows.length,
      filteredRows: filtered.length,
      paginationModel,
      role: currentUserRole || "(blank)",
      email: currentUserEmail || "(blank)",
      location: currentUserLocation || "(blank)",
      showInactiveMeetings,
      filters: { filterBuddy, filterLoc, filterNJ, filterStatus, searchText },
      first10FilteredIds: filtered.slice(0, 10).map((row) => row.id),
    });
  }, [
    rows.length,
    roleScopedRows.length,
    baseVisibleRows.length,
    filtered.length,
    paginationModel,
    currentUserRole,
    currentUserEmail,
    currentUserLocation,
    showInactiveMeetings,
    filterBuddy,
    filterLoc,
    filterNJ,
    filterStatus,
    searchText,
    filtered,
  ]);

  // ─── Feedback helpers ──────────────────────────────────────────────────────

  /**
   * Checks whether the NJ can give feedback for a specific milestone.
   * Condition: the milestone meeting must be "Completed" AND no prior feedback submitted.
   */
  const canGiveFeedback = useCallback(
    (row: BuddyAllocationRow, milestone: "3rd" | "6th"): boolean => {
      if (!isNJLogin) return false;
      const statusField = milestone === "3rd" ? "thirdMeetingStatus" : "sixthMeetingStatus";
      if (row[statusField] !== "Completed") return false;
      const existing = parseFeedbackHistory(row.feedbackJsonHistory);
      return !existing.some((e) => e.milestone === milestone);
    },
    [isNJLogin],
  );

  const hasFeedbackSubmitted = useCallback(
    (row: BuddyAllocationRow, milestone: "3rd" | "6th"): boolean => {
      const existing = parseFeedbackHistory(row.feedbackJsonHistory);
      return existing.some((e) => e.milestone === milestone);
    },
    [],
  );

  // ─── Meeting slots ─────────────────────────────────────────────────────────

  const meetingSlots = useMemo((): MeetingSlot[] => {
    if (!activeRow) return [];
    return MEETING_SLOT_KEYS.map((k, idx) => ({
      no: idx + 1, label: k.label,
      scheduledDate: activeRow[k.scheduled] as string,
      actualDate:    activeRow[k.actual]    as string,
      status:        activeRow[k.status]    as SPMeetingStatus,
    }));
  }, [activeRow]);

  const activeMeetingSlot = useMemo(
    () => selMeeting === "Summary" ? null : (meetingSlots.find((m) => m.label === selMeeting) ?? null),
    [meetingSlots, selMeeting],
  );

  const minScheduledDateForSelectedMeeting = useMemo(() => {
    if (!activeRow || selMeeting === "Summary") return todayInputValue();
    const idx   = MEETING_LABELS.indexOf(selMeeting);
    const today = todayInputValue();
    if (idx <= 0) return today;
    const prevSlot  = MEETING_SLOT_KEYS[idx - 1];
    const prevDate  = activeRow[prevSlot.scheduled] as string;
    const prevInput = toDateInputValue(prevDate);
    if (!prevInput) return today;
    return prevInput > today ? prevInput : today;
  }, [activeRow, selMeeting]);

  // ─── Combine master questions with orphan answers ──────────────────────────
  // For DRAFT meetings: show master + orphan questions
  // For COMPLETED meetings: show ONLY questions that have answers (no unanswered new questions)
  const currentMeetingQuestions = useMemo(() => {
    // Get master questions for this meeting
    const masterQuestions = allQuestions
      .filter((q) => q.meeting === selMeeting)
      .sort((left, right) => left.id - right.id);
    
    // Get all answer questions for this meeting (excluding empty)
    const answerQuestions = savedAnswers
      .filter((a) => a.meeting === selMeeting)
      .map((a) => a.question)
      .filter(Boolean);
    
    // Create a map of master question texts for quick lookup (case-insensitive, trimmed)
    const masterQuestionTexts = new Set(masterQuestions.map((q) => q.question.toLowerCase().trim()));
    
    // Find questions that have answers but are NOT in master list
    const orphanQuestions = Array.from(
      new Set( // Remove duplicates
        answerQuestions
          .filter((q) => !masterQuestionTexts.has(q.toLowerCase().trim()))
      )
    );
    
    // Check if current meeting is completed
    const isMeetingCompleted = meetingSlots.find((m) => m.label === selMeeting)?.status === "Completed";
    
    if (isMeetingCompleted) {
      // For completed meetings: only show questions that have answers
      // Filter master questions to only those with answers
      const masterWithAnswers = masterQuestions.filter((q) => 
        answerQuestions.some((aq) => aq.toLowerCase().trim() === q.question.toLowerCase().trim())
      );
      
      // Add orphan questions (they already have answers)
      const combined: BuddyQuestion[] = [
        ...masterWithAnswers,
        ...orphanQuestions.map((q, idx) => ({
          id: -1000 - idx,
          meeting: selMeeting,
          question: q,
          status: "Active",
        })),
      ];
      
      return combined.sort((left, right) => {
        const leftIsOrphan = left.id < 0;
        const rightIsOrphan = right.id < 0;
        if (leftIsOrphan !== rightIsOrphan) return leftIsOrphan ? 1 : -1;
        return left.id - right.id;
      });
    } else {
      // For draft/pending meetings: show master + orphan questions
      const combined: BuddyQuestion[] = [
        ...masterQuestions,
        ...orphanQuestions.map((q, idx) => ({
          id: -1000 - idx, // Negative IDs for orphan questions
          meeting: selMeeting,
          question: q,
          status: "Active",
        })),
      ];
      
      return combined.sort((left, right) => {
        const leftIsOrphan = left.id < 0;
        const rightIsOrphan = right.id < 0;
        if (leftIsOrphan !== rightIsOrphan) return leftIsOrphan ? 1 : -1;
        return left.id - right.id;
      });
    }
  }, [allQuestions, savedAnswers, selMeeting, meetingSlots]);

  const savedAnswerMap = useMemo(() => {
    const map: Record<string, string> = {};
    // savedAnswers is ID DESC, so retain the first (latest) record per question.
    for (const a of savedAnswers) {
      if (a.meeting === selMeeting && !Object.prototype.hasOwnProperty.call(map, a.question)) {
        map[a.question] = a.answer;
      }
    }
    return map;
  }, [savedAnswers, selMeeting]);

  const savedAnswerStatusMap = useMemo(() => {
    const map: Record<string, string> = {};
    // Keep the status from the same latest answer record used above.
    for (const a of savedAnswers) {
      if (a.meeting === selMeeting && !Object.prototype.hasOwnProperty.call(map, a.question)) {
        map[a.question] = a.status;
      }
    }
    return map;
  }, [savedAnswers, selMeeting]);

  const hasSomeDraftAnswers = useMemo(
    () => Object.values(draftAnswers).some((v) => v && v.trim() !== ""),
    [draftAnswers],
  );

  const isMeetingLocked = useCallback(
    (meetingLabel: string): boolean => {
      if (!activeRow) return false;
      const idx = MEETING_LABELS.indexOf(meetingLabel);
      if (idx <= 0) return false;
      const prevKey    = MEETING_SLOT_KEYS[idx - 1];
      const prevStatus = activeRow[prevKey.status] as SPMeetingStatus;
      return prevStatus !== "Completed";
    },
    [activeRow],
  );

  const allAnswersFilled = useMemo(() => {
    if (currentMeetingQuestions.length === 0) return true;
    return currentMeetingQuestions.every((q) => {
      const saved = savedAnswerMap[q.question];
      const draft = draftAnswers[String(q.id)];
      return (draft && draft.trim() !== "") || (saved && saved.trim() !== "");
    });
  }, [currentMeetingQuestions, savedAnswerMap, draftAnswers]);

  const hasInvalidMeetingAnswers = useMemo(() => {
    return currentMeetingQuestions.some((q) => {
      if (!Object.prototype.hasOwnProperty.call(draftAnswers, String(q.id))) return false;
      const answer = normalizeSingleSpaces(draftAnswers[String(q.id)] ?? "");
      return answer !== "" && answer.length < MIN_MEETING_ANSWER_LENGTH;
    });
  }, [currentMeetingQuestions, draftAnswers]);

  const allAnswersValid = useMemo(() => {
    if (currentMeetingQuestions.length === 0) return true;
    return currentMeetingQuestions.every((q) => {
      const answer = normalizeSingleSpaces(draftAnswers[String(q.id)] ?? savedAnswerMap[q.question] ?? "");
      return answer.length >= MIN_MEETING_ANSWER_LENGTH;
    });
  }, [currentMeetingQuestions, draftAnswers, savedAnswerMap]);

  const isCurrentMeetingCompleted = useMemo(() => {
    if (!activeRow || selMeeting === "Summary") return false;
    return meetingSlots.find((m) => m.label === selMeeting)?.status === "Completed";
  }, [activeRow, meetingSlots, selMeeting]);

  const canCurrentUserEditSelectedMeeting = useMemo(() => canEditRow(activeRow), [canEditRow, activeRow]);

  // ─── Stats ─────────────────────────────────────────────────────────────────

  // Header counters use joinee-level status to avoid confusion:
  // - Completed = joinees whose full meeting cycle is completed
  // - In Progress = active joinees with at least one completed meeting but not all six
  // - Pending = active joinees with no completed meetings yet
  // - Inactive = inactive / released joinees in the current role/location scope
  const activeJoineeCount = roleScopedRows.filter((r) => r.employeeStatus === "Active").length;
//   const inactiveJoineeCount = roleScopedRows.filter((r) => r.employeeStatus !== "Active").length;
  const inactiveMeetingCount = roleScopedRows.filter((r) => r.status === "Inactive").length;
  const completedJoineeCount = roleScopedRows.filter((r) => isAllMeetingsCompleted(r)).length;
  const inProgressJoineeCount = roleScopedRows.filter((r) => {
    const done = completedMeetingCountForRow(r);
    return r.employeeStatus === "Active" && done > 0 && !isAllMeetingsCompleted(r);
  }).length;
  const pendingJoineeCount = roleScopedRows.filter((r) =>
    r.employeeStatus === "Active" && completedMeetingCountForRow(r) === 0,
  ).length;
  const completedMeetingCount = roleScopedRows.reduce(
    (sum, row) => sum + completedMeetingCountForRow(row),
    0,
  );

  const drawerSlotStats = useMemo(() => {
    // Always show summary-level stats (all meetings) regardless of selected meeting
    return {
      completed: meetingSlots.filter((m) => m.status === "Completed").length,
      pending:   meetingSlots.filter((m) => m.status === "Pending").length,
      skipped:   meetingSlots.filter((m) => m.status === "Skipped").length,
      total:     meetingSlots.length,
    };
  }, [meetingSlots]);

  // ─── Feedback questions for current milestone ──────────────────────────────
  const currentFeedbackQuestions = useMemo(
    () => feedbackQuestions
      .filter((q) => q.milestone === feedbackMilestone)
      .sort((a, b) => {
        return a.sortOrder - b.sortOrder || a.id.localeCompare(b.id, undefined, { numeric: true });
      }),
    [feedbackQuestions, feedbackMilestone],
  );

  const feedbackCanSubmit = useMemo(() => {
    return currentFeedbackQuestions.every((q) => {
      const answerValue = (feedbackAnswers[q.id] || "").trim();
      if (q.required && !answerValue) return false;
      if (q.type !== "remark") return true;
      if (!answerValue) return true;
      return normalizeSingleSpaces(answerValue).length >= MIN_MEETING_ANSWER_LENGTH;
    });
  }, [currentFeedbackQuestions, feedbackAnswers]);

  const feedbackValidationMessages = useMemo(() => {
    const messages: Record<string, string> = {};
    currentFeedbackQuestions.forEach((q) => {
      if (q.type !== "remark") return;
      const answerValue = (feedbackAnswers[q.id] || "").trim();
      if (!answerValue) {
        if (q.required) messages[q.id] = "This question is required.";
        return;
      }
      const normalized = normalizeSingleSpaces(answerValue);
      if (normalized.length < MIN_MEETING_ANSWER_LENGTH) {
        messages[q.id] = `Minimum ${MIN_MEETING_ANSWER_LENGTH} characters required (${normalized.length}/${MIN_MEETING_ANSWER_LENGTH}).`;
      }
    });
    return messages;
  }, [currentFeedbackQuestions, feedbackAnswers]);

  // ─── Drawer helpers ────────────────────────────────────────────────────────

  const openView = (row: BuddyAllocationRow) => {
    setActiveRow(row); setSelMeeting("Summary"); setDraftAnswers({}); setSavedAnswers([]);
    setIsEditingDates(false); setEditReason(""); setDateError("");
    setDrawerKind("view");
    void loadAnswersForRow(row.njGlobalId);
  };

  const openReassign = (row: BuddyAllocationRow) => {
    setActiveRow(row); setBuddySearchTerm(""); setBuddySearchResults([]); setSelectedNewBuddy(null);
    setDrawerKind("reassign");
  };

  const openRelease = (row: BuddyAllocationRow) => {
    setActiveRow(row); setRelActive(row.employeeStatus === "Active"); setRelConfirm(false);
    setDrawerKind("release");
  };

  const openFeedback = (row: BuddyAllocationRow, milestone: "3rd" | "6th") => {
    setActiveRow(row);
    setFeedbackMilestone(milestone);
    setFeedbackAnswers({});
    setDrawerKind("feedback");
  };

  const closeDrawer = () => {
    setDrawerKind(null); setActiveRow(null); setSelMeeting("Summary");
    setBuddySearchTerm(""); setBuddySearchResults([]); setSelectedNewBuddy(null);
    setRelConfirm(false); setDraftAnswers({}); setSavedAnswers([]);
    setIsEditingDates(false); setEditReason(""); setDateError("");
    setFeedbackAnswers({});
  };

  const handleMeetingTabSwitch = (label: string) => {
    setSelMeeting(label); setDraftAnswers({}); setIsEditingDates(false);
    setEditReason(""); setDateError("");
  };

  const enterEditDates = () => {
    if (!activeMeetingSlot) return;

    if (activeMeetingSlot.status === "Completed") {
      showSnackbar("Meeting date cannot be updated after the meeting is completed.", "warning", 4000);
      return;
    }

    setEditScheduledDate(toDateInputValue(activeMeetingSlot.scheduledDate));
    setEditReason("");
    setDateError("");
    setIsEditingDates(true);
  };

  const validateDateSelection = useCallback(
    (dateStr: string): string => {
      if (!dateStr || !activeRow) return "";
      if (isWeekend(dateStr)) {
        const d   = new Date(dateStr);
        const day = d.getDay();
        return `${day === 0 ? "Sunday" : "Saturday"} is a non-working day. Please select a weekday.`;
      }
      const relevantLocs = [activeRow.njLocation || "", activeRow.buddyLocation || ""].filter(Boolean);
      const holidayName  = getHolidayMatch(dateStr, holidaysRef.current, relevantLocs);
      if (holidayName) {
        const formattedDate = new Date(dateStr).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
        return `${formattedDate} is a holiday (${holidayName}) for ${relevantLocs.join(" / ")} location(s). Please select a different date.`;
      }
      return "";
    },
    [activeRow],
  );

  // ─── Save meeting dates ────────────────────────────────────────────────────

  const handleSaveDates = async () => {
    if (!activeRow || selMeeting === "Summary" || !editScheduledDate) return;

    if (activeMeetingSlot?.status === "Completed") {
      showSnackbar("Meeting date cannot be updated after the meeting is completed.", "warning", 4000);
      setIsEditingDates(false);
      return;
    }

    if (!editReason.trim()) {
      showSnackbar("Please provide a reason for the date change.", "error", 3000);
      return;
    }
    const validationError = validateDateSelection(editScheduledDate);
    if (validationError) { setDateError(validationError); showSnackbar(validationError, "error", 5000); return; }

    const slotIndex = MEETING_SLOT_KEYS.findIndex((k) => k.label === selMeeting);
    if (slotIndex === -1) return;
    setIsSavingDates(true);
    try {
      const updatePayload: Record<string, any> = {};
      const holidays       = holidaysRef.current;
      const relevantLocs   = [activeRow.njLocation || "", activeRow.buddyLocation || ""].filter(Boolean);
      const anchorDate = new Date(editScheduledDate);
      anchorDate.setHours(0, 0, 0, 0);

      // Only update the specific meeting date that was changed
      const targetKey = MEETING_SLOT_KEYS[slotIndex];
      const iso = anchorDate.toISOString();
      updatePayload[targetKey.spScheduledField] = iso;

      // --- Custom: Update reschedule fields ---
      updatePayload["IsRescheduleMetting"] = true;
      updatePayload["Reschedulemeeting"] = selMeeting; // Internal name for 'Reschedule meeting' Choice field
      // Ensure date columns are set as ISO date strings or null
      updatePayload["RescheduleMeetingDate"] = iso || null;
      updatePayload["OldMeetingDate"] = activeMeetingSlot?.scheduledDate ? new Date(activeMeetingSlot.scheduledDate).toISOString() : null;
      // Textarea field for reason
      updatePayload["RescheduleReason"] = editReason.trim();
      // --- End custom ---

      const historyEntry: MeetingHistoryEntry = {
        meeting: selMeeting, action: "date_updated",
        previousScheduledDate: activeMeetingSlot?.scheduledDate ? formatDate(activeMeetingSlot.scheduledDate) : "—",
        updatedScheduledDate: anchorDate.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" }),
        reason: editReason.trim(),
        updatedBy: currentUserName, updatedByEmail: currentUserEmail,
        updatedAt: new Date().toISOString(),
        locations: relevantLocs.join(", ") as any,
        holidaysConsidered: holidays.filter((h) => {
          const isGlobal = h.locations.length === 0;
          return isGlobal || h.locations.some((hl) => relevantLocs.some((rl) => rl.toLowerCase() === hl.toLowerCase()));
        }).length as any,
      };
      const updatedHistory = buildUpdatedHistory(activeRow, historyEntry);
      updatePayload["MetingJsonHistory"] = updatedHistory;

      const statusFields: (keyof BuddyAllocationRow)[] = ["firstMeetingStatus", "secondMeetingStatus", "thirdMeetingStatus", "fourthMeetingStatus", "fifthMeetingStatus", "sixthMeetingStatus"];
      const pendingIndex = statusFields.findIndex((f) => activeRow[f] === "Pending");
      if (pendingIndex === slotIndex) {
        updatePayload["UpcomingMeetingDate"] = iso;
      }
      updatePayload["RunWF"] = "Yes";

      await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.BuddyAllocate).items.getById(activeRow.spId).update(updatePayload);
      if (!isMountedRef.current) return;

      const updatedRow: BuddyAllocationRow = { ...activeRow, meetingJsonHistory: updatedHistory };
      (updatedRow as Record<string, any>)[targetKey.scheduled as string] = iso;

      setActiveRow(updatedRow);
      setRows((prev) => prev.map((r) => (r.id === activeRow.id ? updatedRow : r)));
      await loadRows();
      setIsEditingDates(false); setEditReason(""); setDateError("");
      showSnackbar(`${selMeeting} date updated successfully!`, "success", 3000);
    } catch (e) {
      console.error("handleSaveDates:", e);
      showSnackbar("Failed to update scheduled dates. Please try again.", "error", 4000);
    } finally {
      if (isMountedRef.current) setIsSavingDates(false);
    }
  };

  // ─── Save answers as Draft ────────────────────────────────────────────────

  const handleSaveDraft = async () => {
    if (!activeRow || selMeeting === "Summary") return;
    if (hasInvalidMeetingAnswers) {
      showSnackbar(`Each answer must be at least ${MIN_MEETING_ANSWER_LENGTH} characters.`, "error", 4000);
      return;
    }
    setIsSavingDraft(true);
    try {
      for (const q of currentMeetingQuestions) {
        const answerText = draftAnswers[String(q.id)]?.trim();
        if (!answerText) continue;
        const existing = savedAnswers.find((a) => a.meeting === selMeeting && a.question === q.question);
        if (existing) {
          await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.MEETING_ANSWERS).items.getById(existing.id).update({ Answers: answerText, Status: "Draft" });
        } else {
          await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.MEETING_ANSWERS).items.add({
            GlobalIDofNJ: activeRow.njGlobalId, Meetings: selMeeting,
            Questions: q.question, Answers: answerText,
            Status: "Draft",
          });
        }
      }
      if (!isMountedRef.current) return;

      const historyEntry: MeetingHistoryEntry = {
        meeting: selMeeting, action: "draft_saved",
        updatedBy: currentUserName, updatedByEmail: currentUserEmail,
        updatedAt: new Date().toISOString(),
      };
      const updatedHistory = buildUpdatedHistory(activeRow, historyEntry);
      await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.BuddyAllocate).items.getById(activeRow.spId).update({ MetingJsonHistory: updatedHistory, RunWF: "Yes" });
      setActiveRow((prev) => prev ? { ...prev, meetingJsonHistory: updatedHistory } : prev);
      setRows((prev) => prev.map((r) => (r.id === activeRow.id ? { ...r, meetingJsonHistory: updatedHistory } : r)));
      await loadAnswersForRow(activeRow.njGlobalId);
      await loadRows();
      setDraftAnswers({});
      showSnackbar(`Answers for ${selMeeting} saved as draft!`, "success", 3000);
    } catch (e) {
      console.error("handleSaveDraft:", e);
      showSnackbar("Failed to save draft. Please try again.", "error", 4000);
    } finally {
      if (isMountedRef.current) setIsSavingDraft(false);
    }
  };

  // ─── Complete meeting ─────────────────────────────────────────────────────

  const handleCompleteMeeting = async () => {
    if (!activeRow || selMeeting === "Summary") return;
    if (!allAnswersValid) {
      showSnackbar(`All answers must be at least ${MIN_MEETING_ANSWER_LENGTH} characters.`, "error", 4000);
      return;
    }
    setIsSavingComplete(true);
    try {
      const slotKeyObj = MEETING_SLOT_KEYS.find((k) => k.label === selMeeting);
      if (!slotKeyObj) throw new Error("Meeting slot key not found");

      for (const q of currentMeetingQuestions) {
        const answerText = draftAnswers[String(q.id)]?.trim() || savedAnswerMap[q.question]?.trim() || "";
        if (!answerText) continue;
        const existing = savedAnswers.find((a) => a.meeting === selMeeting && a.question === q.question);
        if (existing) {
          await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.MEETING_ANSWERS).items.getById(existing.id).update({ Answers: answerText, Status: "Completed" });
        } else {
          await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.MEETING_ANSWERS).items.add({
            GlobalIDofNJ: activeRow.njGlobalId, Meetings: selMeeting,
            Questions: q.question, Answers: answerText,
            Status: "Completed",
          });
        }
      }

      const now = new Date();
      const todayDateOnly = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
      const todayIso = now.toISOString();
      
      const updatePayload: Record<string, string> = {
        [slotKeyObj.spStatusField]: "Completed",
        [slotKeyObj.spActualField]: todayDateOnly,
        CurrentMeetingStatus: `${selMeeting} Completed`,
        RunWF: "Yes",
      };

      const statusFields: (keyof BuddyAllocationRow)[] = ["firstMeetingStatus", "secondMeetingStatus", "thirdMeetingStatus", "fourthMeetingStatus", "fifthMeetingStatus", "sixthMeetingStatus"];
      const updatedStatuses = statusFields.map((f) => f === slotKeyObj.status ? "Completed" : (activeRow[f] as string));
      if (updatedStatuses.every((s) => s === "Completed")) updatePayload["OverallStatus"] = "Completed";

      const nextPendingIndex = updatedStatuses.findIndex((s) => s === "Pending");
      if (nextPendingIndex >= 0) {
        const nextMeetingKey = MEETING_SLOT_KEYS[nextPendingIndex].scheduled;
        const upcomingDate = activeRow[nextMeetingKey];
        if (upcomingDate) {
          updatePayload["UpcomingMeetingDate"] = String(upcomingDate);
        } else {
          updatePayload["UpcomingMeetingDate"] = null as any;
        }
      } else {
        // All meetings completed - clear the UpcomingMeetingDate field by setting to null
        updatePayload["UpcomingMeetingDate"] = null as any;
      }

      const historyEntry: MeetingHistoryEntry = {
        meeting: selMeeting, action: "completed",
        previousStatus: activeMeetingSlot?.status || "Pending", updatedStatus: "Completed",
        actualDate: now.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" }),
        updatedBy: currentUserName, updatedByEmail: currentUserEmail, updatedAt: todayIso,
      };
      const updatedHistory = buildUpdatedHistory(activeRow, historyEntry);
      updatePayload["MetingJsonHistory"] = updatedHistory;

      await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.BuddyAllocate).items.getById(activeRow.spId).update(updatePayload);
      if (!isMountedRef.current) return;

      const patchRow = (r: BuddyAllocationRow) => ({
        ...r,
        [slotKeyObj.status as string]: "Completed",
        [slotKeyObj.actual as string]: todayDateOnly,
        currentMeetingStatus: `${selMeeting} Completed`,
        overallStatus: updatedStatuses.every((s) => s === "Completed") ? "Completed" : r.overallStatus,
        meetingJsonHistory: updatedHistory,
      });

      setRows((prev) => prev.map((r) => (r.id !== activeRow.id ? r : patchRow(r))));
      setActiveRow((prev) => prev ? patchRow(prev) : prev);
      await loadAnswersForRow(activeRow.njGlobalId);
      await loadRows();
      setDraftAnswers({});
      setIsEditingDates(false);
      setEditReason("");
      setDateError("");
      showSnackbar(`${selMeeting} marked as Completed successfully!`, "success", 3000);
    } catch (e) {
      console.error("completeMeeting:", e);
      showSnackbar("Failed to complete meeting. Please try again.", "error", 4000);
    } finally {
      if (isMountedRef.current) setIsSavingComplete(false);
    }
  };

  // ─── Submit NJ Feedback ────────────────────────────────────────────────────

  const handleSubmitFeedback = async () => {
    if (!activeRow) return;
    if (!feedbackCanSubmit) {
      showSnackbar("Please answer all required questions and meet the minimum character requirement for remarks.", "error", 4000);
      return;
    }
    setIsSavingFeedback(true);
    try {
      const cleanedAnswers = Object.fromEntries(
        Object.entries(feedbackAnswers)
          .map(([questionId, answer]) => [questionId, normalizeSingleSpaces(answer)])
          .filter(([, answer]) => Boolean(answer)),
      ) as Record<string, string>;

      const questionSnapshots = Object.fromEntries(
        currentFeedbackQuestions
          .filter((q) => Boolean(cleanedAnswers[q.id]))
          .map((q) => [
            q.id,
            {
              question: normalizeSingleSpaces(q.question),
              type: q.type,
            },
          ]),
      ) as FeedbackEntry['questionSnapshots'];

      const newEntry: FeedbackEntry = {
        milestone: feedbackMilestone,
        submittedAt: new Date().toISOString(),
        submittedByEmail: currentUserEmail,
        answers: cleanedAnswers,
        questionSnapshots,
      };

      let existing: FeedbackEntry[] = [];
      try { existing = JSON.parse(activeRow.feedbackJsonHistory || "[]"); } catch { existing = []; }
      existing.push(newEntry);
      const updatedFeedback = JSON.stringify(existing, null, 2);

      const answeredFeedbackQuestions = currentFeedbackQuestions
        .map((q) => ({
          questionId: q.id,
          question: normalizeSingleSpaces(q.question),
          answer: cleanedAnswers[q.id] || "",
        }))
        .filter((item) => item.answer);

      await sp.web.lists
        .getByTitle(LIST_CONFIG.LISTS.BuddyAllocate)
        .items.getById(activeRow.spId)
        .update({ FeedbackJsonHistory: updatedFeedback, RunWF: "Yes" });

      // Also store submitted feedback question/answer rows in the New Joinee Questionnaire answer list.
      // This keeps the submitted feedback available as normal Q&A records for reporting/export.
      for (const item of answeredFeedbackQuestions) {
        const existingAnswers = await getThresholdSafeListItems({
          web: sp.web,
          listTitle: LIST_CONFIG.LISTS.MEETING_ANSWERS,
          select: "ID,Questions",
          filter:
            `GlobalIDofNJ eq '${activeRow.njGlobalId.replace(/'/g, "''")}' and ` +
            `Meetings eq '${(`${feedbackMilestone} Feedback`).replace(/'/g, "''")}'`,
        });

        const existingAnswer = existingAnswers.filter((a) => a.Questions === item.question);

        if (existingAnswer?.length > 0) {
          await sp.web.lists
            .getByTitle(LIST_CONFIG.LISTS.MEETING_ANSWERS)
            .items.getById(existingAnswer[0].ID)
            .update({ Answers: item.answer, Status: "Completed" });
        } else {
          await sp.web.lists
            .getByTitle(LIST_CONFIG.LISTS.MEETING_ANSWERS)
            .items.add({
              GlobalIDofNJ: activeRow.njGlobalId,
              Meetings: `${feedbackMilestone} Feedback`,
              Questions: item.question,
              Answers: item.answer,
              Status: "Completed",
            });
        }
      }

      if (!isMountedRef.current) return;

      setRows((prev) => prev.map((r) => r.id === activeRow.id ? { ...r, feedbackJsonHistory: updatedFeedback } : r));
      setActiveRow((prev) => prev ? { ...prev, feedbackJsonHistory: updatedFeedback } : prev);
      setFeedbackAnswers(cleanedAnswers);
      await loadAnswersForRow(activeRow.njGlobalId);
      await loadRows();
      showSnackbar(`Feedback for ${feedbackMilestone} meeting milestone submitted successfully!`, "success", 3000);
      closeDrawer();
    } catch (e) {
      console.error("handleSubmitFeedback:", e);
      showSnackbar("Failed to submit feedback. Please try again.", "error", 4000);
    } finally {
      if (isMountedRef.current) setIsSavingFeedback(false);
    }
  };

  // ─── Buddy search ──────────────────────────────────────────────────────────

  const handleBuddySearch = async (term: string) => {
    if (term.trim().length < 2) { setBuddySearchResults([]); return; }
    try {
      const esc = term.replace(/'/g, "''").trim();
      let filter = `PositionType eq 1 and EmployeeStatus eq 'Active' and Role eq 'Buddy' and (substringof('${esc}',FullName) or substringof('${esc}',EmailID) or substringof('${esc}',GlobalID))`;
      if ((isLndAdmin || isSuperAdmin) && activeRow?.njLocation) {
        const njLoc = activeRow.njLocation.replace(/'/g, "''");
        filter += ` and Location eq '${njLoc}'`;
      }
      // const items = await sp.web.lists
      //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
      const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
      const items = await getThresholdSafeListItems({
        web: empMasterWeb,
        listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
        select: "ID,FullName,EmailID,GlobalID,Location,Department",
        filter,
        maxItems: 100,
      });
      if (isMountedRef.current) {
        const mapped = (items || []).map((i: any) => ({
          id: i.ID, name: i.FullName || "", email: i.EmailID || "",
          globalId: i.GlobalID || "", location: i.Location || "",
          department: i.Department || "",
        }));
        setBuddySearchResults(activeRow ? mapped.filter((buddy) => !isSameBuddyAsCurrent(activeRow, buddy)) : mapped);
        setShowBuddyDropdown(true);
      }
    } catch (e) { console.error("buddySearch:", e); }
  };

  // ─── Reassign Buddy ────────────────────────────────────────────────────────

  const handleReassignSave = async () => {
    if (!activeRow || !selectedNewBuddy) return;
    if (isSameBuddyAsCurrent(activeRow, selectedNewBuddy)) {
      showSnackbar("Cannot reassign to the same buddy. Please select a different buddy.", "error", 4000);
      return;
    }
    if (isAllMeetingsCompleted(activeRow)) {
      showSnackbar("Cannot reassign because all meetings for this new joinee are completed.", "warning", 5000);
      return;
    }
    if (activeRow.employeeStatus === "Inactive") {
      showSnackbar("Cannot reassign an inactive new joinee.", "warning", 4000);
      return;
    }
    if ((isLndAdmin || isSuperAdmin) && activeRow.njLocation && selectedNewBuddy.location) {
      if (selectedNewBuddy.location.toLowerCase().trim() !== activeRow.njLocation.toLowerCase().trim()) {
        showSnackbar(
          `Cannot reassign: Selected buddy is in "${selectedNewBuddy.location}" but NJ is in "${activeRow.njLocation}". Location must match.`,
          "error", 5000,
        );
        return;
      }
    }

    // ─── Check buddy availability (max 4 non-completed active assignments) ─
    if (!isBuddyAvailableForReassignment(selectedNewBuddy.globalId, rows)) {
      showSnackbar(
        `Buddy "${selectedNewBuddy.name}" has reached the maximum allocation limit (4 active new joiners). Complete meetings to free up slots.`,
        "error",
        5000,
      );
      return;
    }

    setIsSavingReassign(true);
    try {
      const payload: Record<string, any> = {
        BuddyNameRefChoiceColumn: selectedNewBuddy.name,
        BuddyEmailID:             selectedNewBuddy.email,
        GlobalIDofBuddy:          selectedNewBuddy.globalId,
        LocationofBuddy:          selectedNewBuddy.location,
        BuddyDepartment:          selectedNewBuddy.department,
        RunWF:                    "Yes",
      };
      await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.BuddyAllocate).items.getById(activeRow.spId).update(payload);

      if (activeRow.njSpId) {
        const empPayload: Record<string, any> = { AssignedBuddyGlobalID: selectedNewBuddy.globalId, AssignedBuddyFullName: selectedNewBuddy.name, AssignedBuddyEmail: selectedNewBuddy.email, ReassignBuddyStatus: true, RunWF: "Yes" };
        // await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER).items.getById(activeRow.njSpId).update(empPayload);
        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        await empMasterWeb.lists.getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER).items.getById(activeRow.njSpId).update(empPayload);
      }

      if (!isMountedRef.current) return;
      setRows((prev) => prev.map((r) => r.id === activeRow.id ? {
        ...r, buddyName: selectedNewBuddy.name, buddyEmail: selectedNewBuddy.email,
        buddyGlobalId: selectedNewBuddy.globalId, buddyLocation: selectedNewBuddy.location,
        buddyDepartment: selectedNewBuddy.department,
      } : r));
      await loadRows();
      showSnackbar(`Buddy reassigned to ${selectedNewBuddy.name} successfully!`, "success", 3000);
      closeDrawer();
    } catch (e) {
      console.error("reassignSave:", e);
      showSnackbar("Failed to reassign buddy. Please try again.", "error", 4000);
    } finally {
      if (isMountedRef.current) setIsSavingReassign(false);
    }
  };

  // ─── Release NJ ───────────────────────────────────────────────────────────

  const handleReleaseSave = async () => {
    if (!activeRow) return;
    if (!activeRow.njSpId) { showSnackbar("Cannot find NJ in Employee Master. Operation aborted.", "error", 4000); return; }
    setIsSavingRel(true);
    const nextStatus = relActive ? "Active" : "Inactive";
    try {
      // await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER).items.getById(activeRow.njSpId).update({ EmployeeStatus: nextStatus, RunWF: "Yes" });
      const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
      await empMasterWeb.lists.getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER).items.getById(activeRow.njSpId).update({ EmployeeStatus: nextStatus, RunWF: "Yes" });
      await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.BuddyAllocate).items.getById(activeRow.spId).update({ Status: nextStatus, RunWF: "Yes" });
      if (!isMountedRef.current) return;
      setRows((prev) => prev.map((r) => r.id === activeRow.id ? { ...r, employeeStatus: nextStatus as EmpStatus, status: nextStatus } : r));
      await loadRows();
      showSnackbar(relActive ? `${activeRow.njName} re-activated successfully.` : `${activeRow.njName} released (marked Inactive).`, relActive ? "success" : "warning", 3000);
      closeDrawer();
    } catch (e) {
      console.error("releaseSave:", e);
      showSnackbar("Failed to update employee status. Please try again.", "error", 4000);
    } finally {
      if (isMountedRef.current) setIsSavingRel(false);
    }
  };

  // ─── Export single-row CSV ─────────────────────────────────────────────────

  const handleExport = async () => {
    if (!activeRow) return;

    try {
      const XLSX = await import("xlsx");

      const workbook = XLSX.utils.book_new();

      // ─── Sheet 1: Detail (Meeting Information) ────────────────────────
      const detailData: Record<string, string | number>[] = [];
      meetingSlots.forEach((m) => {
        detailData.push({
          "NJ Name": activeRow.njName,
          "NJ Global ID": activeRow.njGlobalId || "",
          "Buddy Name": activeRow.buddyName,
          "Buddy Global ID": activeRow.buddyGlobalId || "",
          "Meeting": m.label,
          "Status": m.status,
          "Scheduled Date": m.scheduledDate ? formatExportDate(m.scheduledDate) : "",
          "Actual Date": m.actualDate ? formatExportDate(m.actualDate) : "",
          "Location": activeRow.njLocation || "",
          "Department": activeRow.njDepartment || "",
          "DOJ": formatExportDate(activeRow.doj),
          "Allocated On": formatExportDate(activeRow.allocatedOn),
        });
      });
      const detailSheet = XLSX.utils.json_to_sheet(detailData);
      XLSX.utils.book_append_sheet(workbook, detailSheet, "Detail");

      // ─── Sheet 2: Buddy Questions & Answers (if user is admin or buddy) ─
      if (!isNJLogin) {
        const completedAnswers = savedAnswers.filter((a) => a.status === "Completed");
        const qaData: Record<string, string | number>[] = [];

        meetingSlots.forEach((m) => {
          const meetingAnswers = completedAnswers.filter((a) => a.meeting === m.label);
          if (meetingAnswers.length === 0) {
            qaData.push({
              "NJ Name": activeRow.njName,
              "NJ Global ID": activeRow.njGlobalId || "",
              "Buddy Name": activeRow.buddyName,
              "Buddy Global ID": activeRow.buddyGlobalId || "",
              "Meeting": m.label,
              "Question": "",
              "Answer": "",
              "Answer Status": "",
              "Answered By": "",
            });
          } else {
            meetingAnswers.forEach((qa) => {
              // const displayAnsweredBy = (qa.answeredBy && !/system|automate/i.test(qa.answeredBy))
              //   ? qa.answeredBy
              //   : activeRow.buddyName;
              qaData.push({
                "NJ Name": activeRow.njName,
                "NJ Global ID": activeRow.njGlobalId || "",
                "Buddy Name": activeRow.buddyName,
                "Buddy Global ID": activeRow.buddyGlobalId || "",
                "Meeting": m.label,
                "Question": qa.question,
                "Answer": qa.answer,
                "Answer Status": qa.status,
                "Answered By": activeRow.buddyName,
              });
            });
          }
        });

        const qaSheet = XLSX.utils.json_to_sheet(qaData);
        XLSX.utils.book_append_sheet(workbook, qaSheet, "Buddy Questions");
      }

      // ─── Sheet 3: Feedback (if user is admin or joinee) ────────────────
      if (isAdminLogin || isNJLogin) {
        const feedbackEntries = parseFeedbackHistory(activeRow.feedbackJsonHistory);
        const feedbackData: Record<string, string | number>[] = [];

        if (feedbackEntries.length === 0) {
          feedbackData.push({
            "NJ Name": activeRow.njName,
            "NJ Global ID": activeRow.njGlobalId || "",
            "Buddy Name": activeRow.buddyName,
            "Buddy Global ID": activeRow.buddyGlobalId || "",
            "Milestone": "",
            "Question": "",
            "Response": "",
            "Submitted At": "",
            "Submitted By": "",
          });
        } else {
          feedbackEntries.forEach((entry) => {
            const milestone = entry.milestone === "3rd" ? "3rd Milestone" : "6th Milestone";
            const relatedQuestions = getFeedbackQuestionsForEntry(entry, feedbackQuestions);

            if (relatedQuestions.length === 0) {
              feedbackData.push({
                "NJ Name": activeRow.njName,
                "NJ Global ID": activeRow.njGlobalId || "",
                "Buddy Name": activeRow.buddyName,
                "Buddy Global ID": activeRow.buddyGlobalId || "",
                "Milestone": milestone,
                "Question": "",
                "Response": "",
                "Submitted At": formatExportDate(entry.submittedAt),
                "Submitted By": activeRow.njName,
              });
            } else {
              relatedQuestions.forEach((q) => {
                const answer = entry.answers[q.id] || "";
                feedbackData.push({
                  "NJ Name": activeRow.njName,
                  "NJ Global ID": activeRow.njGlobalId || "",
                  "Buddy Name": activeRow.buddyName,
                  "Buddy Global ID": activeRow.buddyGlobalId || "",
                  "Milestone": milestone,
                  "Question": q.question,
                  "Response": q.type === "rating" ? formatStarRating(answer) : answer,
                  "Submitted At": formatExportDate(entry.submittedAt),
                  "Submitted By": activeRow.njName,
                });
              });
            }
          });
        }

        const feedbackSheet = XLSX.utils.json_to_sheet(feedbackData);
        XLSX.utils.book_append_sheet(workbook, feedbackSheet, "Feedback");
      }      // Write the workbook to a file
      XLSX.writeFile(workbook, `${activeRow.njName}_meetings.xlsx`);
    } catch (e) {
      console.error("handleExport:", e);
      showSnackbar("Failed to export data. Please try again.", "error", 4000);
    }
  };

  // ─── Export grid as XLSX ───────────────────────────────────────────────────

  const handleExportGrid = async () => {
    try {
      if (filtered.length === 0) { showSnackbar("No records available to export.", "warning", 3000); return; }
      // Admin export includes Q&A and Feedback; NJ export includes Feedback only
      const includeAnswers = !isNJLogin;
      const includeFeedback = isAdminLogin || isNJLogin; // Include feedback export for Super Admin, L&D Admin, and New Joinee
      let mapped: { njGlobalId: string; meeting: string; question: string; answer: string }[] = [];
      if (includeAnswers) {
        const allCompletedAnswers = await getThresholdSafeListItems({
          web: sp.web,
          listTitle: LIST_CONFIG.LISTS.MEETING_ANSWERS,
          select: "ID,NJID,GlobalIDofNJ,Meetings,Questions,Answers,Status",
          includeItem: (item: any) =>
            String(item.Status || "").trim().toLowerCase() === "completed",
        });
        mapped = (allCompletedAnswers || []).map((a: any) => ({
          njGlobalId: a.GlobalIDofNJ || a.NJID || "",
          meeting: a.Meetings || "", question: a.Questions || "", answer: a.Answers || "",
        }));
      }
      await exportToExcel("all_meetings_details_export", filtered, mapped, includeAnswers, includeFeedback, feedbackQuestions);
      showSnackbar(`Exported ${filtered.length} record(s) to Excel.`, "success", 3000);
    } catch (e) {
      console.error("MeetingDetails: handleExportGrid", e);
      showSnackbar("Failed to export meeting details.", "error", 4000);
    }
  };

  const drawerModalProps = { disablePortal: true, keepMounted: true, container: containerRef.current };

  // ─── DataGrid columns ──────────────────────────────────────────────────────

  const dateCell = (p: GridRenderCellParams) => (
    <span style={{ fontSize: "0.75rem" }}>{formatDate(p.value as string)}</span>
  );

  const columns: GridColDef[] = [
    {
      // field: "actions", headerName: "Actions", width: isNJLogin ? 300 : 300, sortable: false,
      field: "actions", headerName: "Actions",width:300, maxWidth: 300, sortable: false,
      renderCell: (p: GridRenderCellParams) => {
        const row = p.row as BuddyAllocationRow;

        // Admin: View + Reassign + Release
        if (isAdminLogin) {
          const cannotReassign = row.employeeStatus === "Inactive" || isAllMeetingsCompleted(row);
          const reassignTooltip = row.employeeStatus === "Inactive"
            ? "Cannot reassign inactive new joinee"
            : isAllMeetingsCompleted(row)
              ? "Cannot reassign after all meetings are completed"
              : "Reassign buddy";
          return (
            <div className={styles.actionBtns}>
              <Tooltip title="View meeting status">
                <button className={styles.btnView} onClick={() => openView(row)}>
                  <VisibilityOutlinedIcon sx={{ fontSize: "0.78rem" }} />View Status
                </button>
              </Tooltip>
              <Tooltip title={reassignTooltip}>
                <span>
                  <button className={styles.btnReassign} onClick={() => !cannotReassign && openReassign(row)} disabled={cannotReassign}>
                    <SyncAltIcon sx={{ fontSize: "0.78rem" }} />Reassign
                  </button>
                </span>
              </Tooltip>
              <Tooltip title={row.employeeStatus === "Inactive" ? "Already released" : isAllMeetingsCompleted(row) ? "Cannot release after all meetings are completed" : "Release new joiner"}>
                <span>
                  <button
                    className={`${styles.btnRelease} ${row.employeeStatus === "Inactive" ? styles.btnReleased : ""}`}
                    onClick={() => openRelease(row)}
                    disabled={row.employeeStatus === "Inactive" || isAllMeetingsCompleted(row)}
                  >
                    <PersonOffOutlinedIcon sx={{ fontSize: "0.78rem" }} />Release NJ
                  </button>
                </span>
              </Tooltip>
            </div>
          );
        }

        // NJ: View Status + conditional Feedback buttons
        if (isNJLogin) {
          const show3rdFeedback = canGiveFeedback(row, "3rd");
          const show6thFeedback = canGiveFeedback(row, "6th");
          const submitted3rd    = hasFeedbackSubmitted(row, "3rd");
          const submitted6th    = hasFeedbackSubmitted(row, "6th");
          return (
            <div className={styles.actionBtns} style={{  gap: "0.25rem" }}>
              <Tooltip title="View your meeting status">
                <button className={styles.btnView} onClick={() => openView(row)}>
                  <VisibilityOutlinedIcon sx={{ fontSize: "0.78rem" }} />View Status
                </button>
              </Tooltip>
              {show3rdFeedback && (
                <Tooltip title="Share your mid-program feedback">
                  <button className={styles.btnFeedback} onClick={() => openFeedback(row, "3rd")}>
                    <FeedbackOutlinedIcon sx={{ fontSize: "0.78rem" }} />Feedback (3rd)
                  </button>
                </Tooltip>
              )}
              {show6thFeedback && (
                <Tooltip title="Share your end-of-program feedback">
                  <button className={styles.btnFeedback} onClick={() => openFeedback(row, "6th")}>
                    <StarOutlineIcon sx={{ fontSize: "0.78rem" }} />Feedback (6th)
                  </button>
                </Tooltip>
              )}
              {submitted3rd && !show3rdFeedback && (
                <Tooltip title="3rd meeting feedback submitted">
                  <div className={styles.feedbackSubmittedBadge}>
                    <ThumbUpOutlinedIcon sx={{ fontSize: "0.72rem" }} /> 3rd ✓
                  </div>
                </Tooltip>
              )}
              {submitted6th && !show6thFeedback && (
                <Tooltip title="6th meeting feedback submitted">
                  <div className={styles.feedbackSubmittedBadge}>
                    <ThumbUpOutlinedIcon sx={{ fontSize: "0.72rem" }} /> 6th ✓
                  </div>
                </Tooltip>
              )}
            </div>
          );
        }

        // Buddy and other roles: View Status only
        return (
          <div className={styles.actionBtns}>
            <Tooltip title="View meeting status">
              <button className={styles.btnView} onClick={() => openView(row)}>
                <VisibilityOutlinedIcon sx={{ fontSize: "0.78rem" }} />View Status
              </button>
            </Tooltip>
          </div>
        );
      },
    },
    {
      field: "overallStatus", headerName: "Status", width: 135,
      type: "singleSelect", valueOptions: statusOpts,
      renderCell: (p: GridRenderCellParams) => <SChip s={p.value as string} kind="overall" />,
    },
    {
      field: "currentMeetingStatus", headerName: "Current Status", width: 175,
      renderCell: (p: GridRenderCellParams) => <SChip s={String(p.value || "")} kind="current" fontSize="0.65rem" />,
    },
    { field: "njName",          headerName: "NJ Name",          flex: 1,   minWidth: 130, renderCell: (p: GridRenderCellParams) => <span style={{ fontWeight: 600, color: "#047857" }}>{p.value as string}</span> },
    { field: "njEmail",         headerName: "NJ Email",         flex: 1,   minWidth: 150 },
    { field: "njGlobalId",      headerName: "Global ID (NJ)",   width: 125 },
    { field: "njDepartment",    headerName: "Department",       flex: 0.9, minWidth: 120 },
    { field: "njLocation",      headerName: "NJ Location",      width: 120 },
    { field: "doj",             headerName: "Date of Joining",  width: 120, renderCell: dateCell },
    { field: "buddyName",       headerName: "Buddy Name",       flex: 1,   minWidth: 130 },
    { field: "buddyEmail",      headerName: "Buddy Email",      flex: 1,   minWidth: 150 },
    { field: "buddyGlobalId",   headerName: "Global ID (Buddy)",width: 135 },
    { field: "buddyDepartment", headerName: "Buddy Dept",       flex: 0.9, minWidth: 110 },
    { field: "buddyLocation",   headerName: "Buddy Location",   width: 120 },
    { field: "allocatedOn",     headerName: "Allocated On",     width: 120, renderCell: dateCell },
    { field: "allocationCount", headerName: "Alloc Count",      width: 90 },
    { field: "firstInteractionDate",  headerName: "1st Interaction", width: 125, renderCell: dateCell },
    { field: "secondInteractionDate", headerName: "2nd Interaction", width: 125, renderCell: dateCell },
    { field: "thirdInteractionDate",  headerName: "3rd Interaction", width: 125, renderCell: dateCell },
    { field: "fourthInteractionDate", headerName: "4th Interaction", width: 125, renderCell: dateCell },
    { field: "fifthInteractionDate",  headerName: "5th Interaction", width: 125, renderCell: dateCell },
    { field: "sixthInteractionDate",  headerName: "6th Interaction", width: 125, renderCell: dateCell },
    { field: "firstMeetingStatus", type: "singleSelect", valueOptions: ["Pending", "Completed", "Skipped"],  headerName: "M1 Status", width: 105, sortable: false, renderCell: (p: GridRenderCellParams) => <SChip s={p.value as SPMeetingStatus} /> },
    { field: "secondMeetingStatus", type: "singleSelect", valueOptions: ["Pending", "Completed", "Skipped"], headerName: "M2 Status", width: 105, sortable: false, renderCell: (p: GridRenderCellParams) => <SChip s={p.value as SPMeetingStatus} /> },
    { field: "thirdMeetingStatus", type: "singleSelect", valueOptions: ["Pending", "Completed", "Skipped"],  headerName: "M3 Status", width: 105, sortable: false, renderCell: (p: GridRenderCellParams) => <SChip s={p.value as SPMeetingStatus} /> },
    { field: "fourthMeetingStatus", type: "singleSelect", valueOptions: ["Pending", "Completed", "Skipped"], headerName: "M4 Status", width: 105, sortable: false, renderCell: (p: GridRenderCellParams) => <SChip s={p.value as SPMeetingStatus} /> },
    { field: "fifthMeetingStatus", type: "singleSelect", valueOptions: ["Pending", "Completed", "Skipped"],  headerName: "M5 Status", width: 105, sortable: false, renderCell: (p: GridRenderCellParams) => <SChip s={p.value as SPMeetingStatus} /> },
    { field: "sixthMeetingStatus", type: "singleSelect", valueOptions: ["Pending", "Completed", "Skipped"],  headerName: "M6 Status", width: 105, sortable: false, renderCell: (p: GridRenderCellParams) => <SChip s={p.value as SPMeetingStatus} /> },
    { field: "firstActualInteractionDate",  headerName: "M1 Actual", width: 115, renderCell: dateCell },
    { field: "secondActualInteractionDate", headerName: "M2 Actual", width: 115, renderCell: dateCell },
    { field: "thirdActualInteractionDate",  headerName: "M3 Actual", width: 115, renderCell: dateCell },
    { field: "fourthActualInteractionDate", headerName: "M4 Actual", width: 115, renderCell: dateCell },
    { field: "fifthActualInteractionDate",  headerName: "M5 Actual", width: 115, renderCell: dateCell },
    { field: "sixthActualInteractionDate",  headerName: "M6 Actual", width: 115, renderCell: dateCell },
    {
      field: "employeeStatus", headerName: "Emp. Status", width: 105,
      type: "singleSelect", valueOptions: ["Active", "Inactive"],
      renderCell: (p: GridRenderCellParams) => <SChip s={p.value as string} kind="employee" />,
    },
  ];

  // ─── Render ────────────────────────────────────────────────────────────────

  return (
    <div className={styles.Container} ref={containerRef}>

      {/* ── Page Heading ─────────────────────────── */}
      <div className={styles.pageHeading}>
        <div className={styles.pageHeadingLeft}>
          <div className={styles.pageHeadingIcon}>
            <CalendarMonthOutlinedIcon sx={{ fontSize: "1.1rem" }} />
          </div>
          <div>
            <h2 className={styles.pageTitle}>All Meetings Details</h2>
            <p className={styles.pageSubtitle}>
              {isNJLogin ? "Track your buddy meeting progress" : "Track and manage buddy–NJ meeting progress"}
            </p>
          </div>
        </div>
        <div className={styles.pageHeadingRight}>
          {isLndAdmin && currentUserLocation && (
            <Tooltip title={`You are managing data for location: ${currentUserLocation}`}>
              <div className={styles.locationScopeBadge}>
                <LocationOnOutlinedIcon sx={{ fontSize: "0.85rem" }} />
                <span>{currentUserLocation}</span>
              </div>
            </Tooltip>
          )}

          {/* Hide "Show Inactive" toggle for NJ */}
          {!isNJLogin && (
            <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
              {/* <div style={{ display: "flex", alignItems: "center", gap: "6px", background: "rgba(255,255,255,0.12)", border: "1px solid rgba(255,255,255,0.18)", borderRadius: "7px", padding: "0.35rem 0.6rem", color: "#fff", fontSize: "0.72rem", fontWeight: 600 }}>
                <span>Show Inactive Joinees ({inactiveJoineeCount})</span>
                <Switch checked={showInactive} onChange={(e) => setShowInactive(e.target.checked)} size="small"
                  sx={{ "& .MuiSwitch-switchBase": { color: "#fff" }, "& .MuiSwitch-track": { backgroundColor: "rgba(255,255,255,0.35)" }, "& .MuiSwitch-switchBase.Mui-checked": { color: "#fff" }, "& .MuiSwitch-switchBase.Mui-checked + .MuiSwitch-track": { backgroundColor: "#00A859" } }} />
              </div> */}

              <div style={{ display: "flex", alignItems: "center", gap: "6px", background: "rgba(255,255,255,0.12)", border: "1px solid rgba(255,255,255,0.18)", borderRadius: "7px", padding: "0.35rem 0.6rem", color: "#fff", fontSize: "0.72rem", fontWeight: 600 }}>
                <span>Show Inactive Meetings ({inactiveMeetingCount})</span>
                <Switch checked={showInactiveMeetings} onChange={(e) => setShowInactiveMeetings(e.target.checked)} size="small"
                  sx={{ "& .MuiSwitch-switchBase": { color: "#fff" }, "& .MuiSwitch-track": { backgroundColor: "rgba(255,255,255,0.35)" }, "& .MuiSwitch-switchBase.Mui-checked": { color: "#fff" }, "& .MuiSwitch-switchBase.Mui-checked + .MuiSwitch-track": { backgroundColor: "#00A859" } }} />
              </div>
            </div>
          )}

          <button className={styles.exportBtn} onClick={() => void handleExportGrid()}>
            <FileDownloadOutlinedIcon sx={{ fontSize: "0.85rem" }} />
            Export Excel
          </button>
        </div>

          
      </div>
      <BuddyLoader visible={showStatsLoader} variant="default" message="Loading Detail..." subMessage="Preparing your meeting statistics." />
      {!showStatsLoader && !isNJLogin && (
      <div>
        <div className={styles.pageHeadingStats}>
            <Tooltip title="Total joinees in your current role/location scope, including inactive joinees.">
              <div className={styles.headingStat}><span className={styles.headingStatVal}>{roleScopedRows.length}</span><span className={styles.headingStatLbl}>Total Joinees</span></div>
            </Tooltip>
            <div className={styles.headingStatDivider} />
            <Tooltip title="Active joinees currently visible when Show Inactive is off.">
              <div className={styles.headingStat}><span className={styles.headingStatVal}>{activeJoineeCount}</span><span className={styles.headingStatLbl}>Active</span></div>
            </Tooltip>
            <div className={styles.headingStatDivider} />
            {/* <Tooltip title="Inactive / released joinees in your current scope.">
              <div className={styles.headingStat}><span className={styles.headingStatVal}>{inactiveJoineeCount}</span><span className={styles.headingStatLbl}>Inactive</span></div>
            </Tooltip> */}
            <div className={styles.headingStatDivider} />
            <Tooltip title="Joinees whose 1st through 6th meetings are completed.">
              <div className={styles.headingStat}><span className={styles.headingStatVal}>{completedJoineeCount}</span><span className={styles.headingStatLbl}>Completed</span></div>
            </Tooltip>
            <div className={styles.headingStatDivider} />
            <Tooltip title="Active joinees who have started meetings but have not completed all six.">
              <div className={styles.headingStat}><span className={styles.headingStatVal}>{inProgressJoineeCount}</span><span className={styles.headingStatLbl}>In Progress</span></div>
            </Tooltip>
            <div className={styles.headingStatDivider} />
            <Tooltip title="Active joinees with no completed meetings yet.">
              <div className={styles.headingStat}><span className={styles.headingStatVal}>{pendingJoineeCount}</span><span className={styles.headingStatLbl}>Pending</span></div>
            </Tooltip>
            <div className={styles.headingStatDivider} />
            <Tooltip title="Total completed individual meeting slots across all scoped joinees. Example: 2 completed joinees x 6 meetings = 12.">
              <div className={styles.headingStat}><span className={styles.headingStatVal}>{completedMeetingCount}</span><span className={styles.headingStatLbl}>Meetings Done</span></div>
            </Tooltip>
          </div>
      </div>
      )}

      {loadError && (
        <div className={styles.errorBanner}>
          <WarningAmberIcon sx={{ fontSize: "0.95rem" }} />
          {loadError}
        </div>
      )}

      {/* ── Filter bar — hide for NJ (they see only their own record) ───── */}
      {!isNJLogin && (
        <div className={styles.filterBar}>
          <div className={styles.filterGroup}>
            {/* Buddy Name filter */}
            <div className={styles.filterWrap}>
              {isBuddyLogin && buddySelfInfo ? (
                <Tooltip title="Showing only your assigned new joiners">
                  <div className={styles.filterDisabledChip}>
                    <span className={styles.filterDisabledDot} />
                    <span className={styles.filterDisabledLabel}>{buddySelfInfo.name || "My NJs"}</span>
                    <span className={styles.filterDisabledBadge}>Buddy</span>
                  </div>
                </Tooltip>
              ) : (
                <Select value={filterBuddy} onChange={(e: SelectChangeEvent) => setFilterBuddy(e.target.value)}
                  displayEmpty renderValue={(v) => v || "Buddy Name"} size="small"
                  IconComponent={KeyboardArrowDownIcon} sx={FILTER_SELECT_SX}>
                  <MenuItem value="" sx={{ fontSize: "0.775rem", color: "#9ca3af" }}><em>All Buddy Names</em></MenuItem>
                  {buddyOpts.map((o) => (<MenuItem key={o} value={o} sx={{ fontSize: "0.775rem" }}>{o}</MenuItem>))}
                </Select>
              )}
            </div>

            {/* Buddy Location filter */}
            <div className={styles.filterWrap}>
              {isBuddyLogin && buddySelfInfo ? (
                <Tooltip title="Filtered to your location">
                  <div className={styles.filterDisabledChip}>
                    <span className={styles.filterDisabledDot} />
                    <span className={styles.filterDisabledLabel}>{buddySelfInfo.location || "My Location"}</span>
                    <span className={styles.filterDisabledBadge}>Location</span>
                  </div>
                </Tooltip>
              ) : isLndAdmin && currentUserLocation ? (
                <Tooltip title={`Data scoped to your location: ${currentUserLocation}`}>
                  <div className={styles.filterDisabledChip}>
                    <span className={styles.filterDisabledDot} />
                    <span className={styles.filterDisabledLabel}>{currentUserLocation}</span>
                    <span className={styles.filterDisabledBadge}>L&D Scope</span>
                  </div>
                </Tooltip>
              ) : (
                <Select value={filterLoc} onChange={(e: SelectChangeEvent) => setFilterLoc(e.target.value)}
                  displayEmpty renderValue={(v) => v || "Buddy Location"} size="small"
                  IconComponent={KeyboardArrowDownIcon} sx={FILTER_SELECT_SX}>
                  <MenuItem value="" sx={{ fontSize: "0.775rem", color: "#9ca3af" }}><em>All Locations</em></MenuItem>
                  {locOpts.map((o) => (<MenuItem key={o} value={o} sx={{ fontSize: "0.775rem" }}>{o}</MenuItem>))}
                </Select>
              )}
            </div>

            {/* NJ Name filter */}
            <div className={styles.filterWrap}>
              <Select value={filterNJ} onChange={(e: SelectChangeEvent) => setFilterNJ(e.target.value)}
                displayEmpty renderValue={(v) => v || "NJ Name"} size="small"
                IconComponent={KeyboardArrowDownIcon} sx={FILTER_SELECT_SX}>
                <MenuItem value="" sx={{ fontSize: "0.775rem", color: "#9ca3af" }}><em>All NJ Names</em></MenuItem>
                {njOpts.map((o) => (<MenuItem key={o} value={o} sx={{ fontSize: "0.775rem" }}>{o}</MenuItem>))}
              </Select>
            </div>

            {/* Status filter */}
            <div className={styles.filterWrap}>
              <Select value={filterStatus} onChange={(e: SelectChangeEvent) => setFilterStatus(e.target.value)}
                displayEmpty renderValue={(v) => v || "Status"} size="small"
                IconComponent={KeyboardArrowDownIcon} sx={FILTER_SELECT_SX}>
                <MenuItem value="" sx={{ fontSize: "0.775rem", color: "#9ca3af" }}><em>Status</em></MenuItem>
                {statusOpts.map((o) => (<MenuItem key={o} value={o} sx={{ fontSize: "0.775rem" }}>{o}</MenuItem>))}
              </Select>
            </div>
          </div>

          <div className={styles.filterActions}>
            <TextField placeholder="Search records…" value={searchText}
              onChange={(e) => setSearchText(e.target.value)} size="small" variant="outlined"
              InputProps={{ startAdornment: (<InputAdornment position="start"><SearchIcon sx={{ fontSize: "0.95rem", color: "#047857" }} /></InputAdornment>) }}
              sx={{ width: "185px", "& .MuiOutlinedInput-root": { fontSize: "0.775rem", "& fieldset": { borderColor: "#e5e7eb" }, "&:hover fieldset": { borderColor: "#00A859" }, "&.Mui-focused fieldset": { borderColor: "#00A859" } } }} />

            {filterBuddy  && <div className={styles.activeFilterChip}><span>{filterBuddy}</span><button onClick={() => setFilterBuddy("")}  aria-label="Remove buddy filter">&times;</button></div>}
            {filterLoc    && <div className={styles.activeFilterChip}><span>{filterLoc}</span><button onClick={() => setFilterLoc("")}    aria-label="Remove location filter">&times;</button></div>}
            {filterNJ     && <div className={styles.activeFilterChip}><span>{filterNJ}</span><button onClick={() => setFilterNJ("")}      aria-label="Remove NJ filter">&times;</button></div>}
            {filterStatus && <div className={styles.activeFilterChip}><span>{filterStatus}</span><button onClick={() => setFilterStatus("")} aria-label="Remove status filter">&times;</button></div>}

            {(filterBuddy || filterLoc || filterNJ || filterStatus || searchText) && (
              <button className={styles.clearBtn} onClick={() => { setFilterBuddy(""); setFilterLoc(""); setFilterNJ(""); setFilterStatus(""); setSearchText(""); }}>
                Clear All
              </button>
            )}

            {/* <div className={styles.recordCount}>
              {filtered.length} / {baseVisibleRows.length} shown · {roleScopedRows.length} total · {inactiveJoineeCount} inactive
            </div> */}
          </div>
        </div>
      )}

      {/* ── DataGrid ─────────────────────────────── */}
      <div className={styles.gridCard}>
        <DataGrid
          key={`meeting-details-grid-${isUserBootstrapped ? "ready" : "booting"}-${rows.length}-${filtered.length}`}
          rows={filtered}
          columns={columns}
          loading={isLoading || !isUserBootstrapped}
          disableRowSelectionOnClick
          paginationModel={paginationModel}
          onPaginationModelChange={(model) => {
            console.info("[MeetingDetails/grid] pagination change", {
              from: paginationModel,
              to: model,
              filteredRows: filtered.length,
              first10FilteredIds: filtered.slice(0, 10).map((row) => row.id),
            });
            setPaginationModel(model);
          }}
          pageSizeOptions={[5, 10, 20, 50]}
          rowHeight={40}
          columnHeaderHeight={40}
          getRowClassName={(p) => p.row.status === "Inactive" ? "row-inactive" : ""}
          sx={GRID_SX}
        />
      </div>

      {/* ══════════════════════════════════════════════
          VIEW STATUS DRAWER
      ══════════════════════════════════════════════ */}
      <Drawer anchor="right" open={drawerKind === "view"} onClose={closeDrawer} ModalProps={drawerModalProps}
        // sx={{ "& .MuiDrawer-paper": { width: { xs: "100%", sm: "720px" }, boxShadow: "-4px 0 28px rgba(0,0,0,0.18)" } }}
        sx={{
          ...DRAWER_OVER_TOPBAR_SX,
          "& .MuiDrawer-paper": {
            ...DRAWER_OVER_TOPBAR_SX["& .MuiDrawer-paper"],
            width: { xs: "100%", sm: "720px" },
            boxShadow: "-4px 0 28px rgba(0,0,0,0.18)",
          },
        }}
        >
        {activeRow && (
          <div className={styles.drawer}>
            <div className={styles.drawerHeader}>
              <div className={styles.dHLeft}>
                <div>
                  <h2 className={styles.drawerTitle}>Meeting Status</h2>
                  <p className={styles.drawerSub}>NJ: <strong>{activeRow.njName}</strong>&nbsp;·&nbsp; Buddy: <strong>{activeRow.buddyName}</strong></p>
                  {/* <p style={{ fontSize: "0.72rem", color: "#6b7280", marginTop: "4px", lineHeight: "1.4" }}>
                    <span style={{ marginRight: "12px" }}>
                      <LocationOnOutlinedIcon sx={{ fontSize: "0.75rem", verticalAlign: "middle", marginRight: "3px" }} />
                      <strong>NJ Location:</strong> {activeRow.njLocation || "—"}
                    </span>
                    <span style={{ marginRight: "12px" }}>
                      <strong>Dept:</strong> {activeRow.njDepartment || "—"}
                    </span>
                  </p>
                  <p style={{ fontSize: "0.72rem", color: "#6b7280", marginTop: "2px", lineHeight: "1.4" }}>
                    <span style={{ marginRight: "12px" }}>
                      <LocationOnOutlinedIcon sx={{ fontSize: "0.75rem", verticalAlign: "middle", marginRight: "3px" }} />
                      <strong>Buddy Location:</strong> {activeRow.buddyLocation || "—"}
                    </span>
                    <span style={{ marginRight: "12px" }}>
                      <strong>Dept:</strong> {activeRow.buddyDepartment || "—"}
                    </span>
                  </p> */}
                </div>
              </div>
              <div className={styles.dHRight}>
                {/* ✅ NEW: Feedback button for NJ after 3rd or 6th meeting */}
                {isNJLogin && canNJSubmitFeedback("3rd", activeRow) && !hasNJSubmittedFeedback("3rd", currentUserEmail, activeRow) && (
                  <Tooltip title="Provide feedback after 3rd meeting">
                    <button
                      className={styles.exportBtn}
                      onClick={() => openFeedback(activeRow, "3rd")}
                      style={{ background: "#7c3aed" }}>
                      <DescriptionOutlinedIcon sx={{ fontSize: "0.85rem" }} />
                     3rd Feedback
                    </button>
                  </Tooltip>
                )}
                {isNJLogin && canNJSubmitFeedback("6th", activeRow) && !hasNJSubmittedFeedback("6th", currentUserEmail, activeRow) && (
                  <Tooltip title="Provide feedback after 6th meeting">
                    <button
                      className={styles.exportBtn}
                      onClick={() => openFeedback(activeRow, "6th")}
                      style={{ background: "#7c3aed" }}>
                      <DescriptionOutlinedIcon sx={{ fontSize: "0.85rem" }} />
                      6th Feedback
                    </button>
                  </Tooltip>
                )}
                {/* NJ sees Export CSV but WITHOUT Q&A */}
                <button className={styles.exportBtn} onClick={handleExport}><FileDownloadOutlinedIcon sx={{ fontSize: "0.85rem" }} />Export CSV</button>
                <IconButton onClick={closeDrawer} className={styles.closeIconBtn}><CloseIcon sx={{ fontSize: "1rem" }} /></IconButton>
              </div>
            </div>

            <div className={styles.meetingTabsBar}>
              {["Summary", ...meetingSlots.map((m) => m.label)].map((lbl) => {
                const locked = lbl !== "Summary" && isMeetingLocked(lbl);
                const slot   = meetingSlots.find((m) => m.label === lbl);
                return (
                  <button key={lbl}
                    className={`${styles.meetingTab} ${selMeeting === lbl ? styles.meetingTabActive : ""} ${locked ? styles.meetingTabLocked : ""}`}
                    onClick={() => handleMeetingTabSwitch(lbl)}>
                    {lbl !== "Summary" && locked ? <LockOutlinedIcon sx={{ fontSize: "0.8rem", opacity: 0.5 }} /> : lbl !== "Summary" ? <StatusIcon status={slot?.status ?? "Pending"} /> : null}
                    {lbl}
                  </button>
                );
              })}
            </div>

            <div className={styles.drawerBody}>
              <div className={styles.statsStrip}>
                <div className={`${styles.statBox} ${styles.statCompleted}`}><span className={styles.statNum}>{drawerSlotStats.completed}</span><span className={styles.statLbl}>Completed</span></div>
                <div className={`${styles.statBox} ${styles.statInProgress}`}><span className={styles.statNum}>{drawerSlotStats.pending}</span><span className={styles.statLbl}>Pending</span></div>
                <div className={`${styles.statBox} ${styles.statSkipped}`}><span className={styles.statNum}>{drawerSlotStats.total}</span><span className={styles.statLbl}>Total Meeting</span></div>
              </div>

              <div className={styles.infoStrip}>
                <div className={styles.infoGroup}><span className={styles.infoLabel}>NJ Global ID</span><span className={styles.infoValue}>{activeRow.njGlobalId || "—"}</span></div>
                <div className={styles.infoGroup}><span className={styles.infoLabel}>NJ Email</span><span className={styles.infoValue}>{activeRow.njEmail || "—"}</span></div>
                <div className={styles.infoGroup}><span className={styles.infoLabel}>Department</span><span className={styles.infoValue}>{activeRow.njDepartment || "—"}</span></div>
                <div className={styles.infoGroup}><span className={styles.infoLabel}>DOJ</span><span className={styles.infoValue}>{formatDate(activeRow.doj)}</span></div>
                <div className={styles.infoGroup}><span className={styles.infoLabel}>Allocated On</span><span className={styles.infoValue}>{formatDate(activeRow.allocatedOn)}</span></div>
                <div className={styles.infoGroup}><span className={styles.infoLabel}>Status</span><span className={styles.infoValue}><SChip s={activeRow.overallStatus} kind="overall" /></span></div>
              </div>

              {selMeeting === "Summary" ? (
                <div>
                  <div className={styles.meetingCardsGrid}>
                    {meetingSlots.map((m) => {
                      const locked      = isMeetingLocked(m.label);
                      // NJ should NOT see Q count / answered count
                      const qCount      = isNJLogin ? 0 : allQuestions.filter((q) => q.meeting === m.label).length;
                      const aCount      = isNJLogin ? 0 : savedAnswers.filter((a) => a.meeting === m.label && a.status === "Completed").length;
                      // ─── Only show draft count to the person editing (buddy) ─
                      const draftACount = (isNJLogin || !canCurrentUserEditSelectedMeeting) ? 0 : savedAnswers.filter((a) => a.meeting === m.label && a.status === "Draft").length;
                      const statusClass = m.status === "Completed" ? styles.meetingCompleted : m.status === "Pending" ? styles.meetingPending : m.status === "Skipped" ? styles.meetingSkipped : "";
                      return (
                        <div key={m.label} className={`${styles.meetingCard} ${statusClass} ${locked ? styles.meetingCardLocked : ""}`} onClick={() => handleMeetingTabSwitch(m.label)}>
                          <div className={styles.mcHeader}>
                            <span className={styles.mcLabel}>{m.label}</span>
                            {locked ? <LockOutlinedIcon sx={{ fontSize: "0.9rem", color: "#9ca3af" }} /> : <StatusIcon status={m.status} />}
                          </div>
                          <SChip s={m.status} />
                          {!isNJLogin && qCount > 0 && (
                            <div className={styles.mcQaBadgeRow}>
                              <div className={styles.mcQaBadge}><QuestionAnswerOutlinedIcon sx={{ fontSize: "0.7rem" }} />{aCount}/{qCount} answered</div>
                              {draftACount > 0 && m.status !== "Completed" && <div className={styles.mcDraftBadge}><BookmarkOutlinedIcon sx={{ fontSize: "0.7rem" }} />{draftACount} draft</div>}
                            </div>
                          )}
                          <div className={styles.mcDates}>
                            <div className={styles.mcDateRow}><span className={styles.mcDateLbl}>Scheduled</span><span className={styles.mcDateVal}>{formatDate(m.scheduledDate)}</span></div>
                            <div className={styles.mcDateRow}><span className={styles.mcDateLbl}>Actual</span><span className={styles.mcDateVal}>{formatDate(m.actualDate)}</span></div>
                          </div>
                        </div>
                      );
                    })}
                  </div>

                  {/* Meeting Update History — visible to all roles EXCEPT New Joinees */}
                  {!isNJLogin && activeRow?.meetingJsonHistory && parseMeetingHistory(activeRow.meetingJsonHistory).length > 0 && (
                    <div className={styles.historySection}>
                      <div className={styles.historySectionHeader} onClick={() => setShowHistory(!showHistory)}>
                        <div className={styles.historySectionTitle}>
                          <HistoryOutlinedIcon sx={{ fontSize: "0.9rem" }} />
                          Meeting Update History
                        </div>
                        {showHistory ? (
                          <ExpandLessIcon className={styles.historyToggleIcon} />
                        ) : (
                          <ExpandMoreIcon className={`${styles.historyToggleIcon} ${styles.collapsed}`} />
                        )}
                      </div>

                      {showHistory && (
                        <div className={styles.historyTimeline}>
                          {parseMeetingHistory(activeRow.meetingJsonHistory)
                            .sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime())
                            .map((entry, idx) => (
                              <div key={idx} className={styles.historyTimelineItem}>
                                <div className={styles.historyTimelineMarker}>
                                  <div className={styles.historyTimelineDot} />
                                </div>
                                <div className={styles.historyTimelineContent}>
                                  <div className={styles.historyItemHeader}>
                                    <div className={styles.historyActionBadge}>
                                      {entry.action === "date_updated" ? "📅 Date Updated" : entry.action === "completed" ? "✓ Completed" : "💾 Draft Saved"}
                                    </div>
                                    <div className={styles.historyItemMeta}>
                                      <span>{entry.meeting}</span>
                                      <span className={styles.historyMetaSeparator}>•</span>
                                      <span>{new Date(entry.updatedAt).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }).replace(",", " at")}</span>
                                    </div>
                                  </div>

                                  <div className={styles.historyItemDetails}>
                                    {entry.action === "date_updated" && (
                                      <>
                                        <div className={styles.historyDetailRow}>
                                          <span className={styles.historyDetailLabel}>Meeting:</span>
                                          <span className={styles.historyDetailValue}>{entry.meeting}</span>
                                        </div>
                                        <div className={styles.historyDetailRow}>
                                          <span className={styles.historyDetailLabel}>Previous Date:</span>
                                          <span className={`${styles.historyDetailValue} ${styles.historyDetailOld}`}>{entry.previousScheduledDate || "—"}</span>
                                        </div>
                                        <div className={styles.historyDetailRow}>
                                          <span className={styles.historyDetailLabel}>Updated Date:</span>
                                          <span className={`${styles.historyDetailValue} ${styles.historyDetailNew}`}>{entry.updatedScheduledDate || "—"}</span>
                                        </div>
                                        {entry.reason && (
                                          <div className={styles.historyDetailRow}>
                                            <span className={styles.historyDetailLabel}>Reason:</span>
                                            <span className={styles.historyDetailValue}>{entry.reason}</span>
                                          </div>
                                        )}
                                        {entry.locations && (
                                          <div className={styles.historyDetailRow}>
                                            <span className={styles.historyDetailLabel}>Locations:</span>
                                            <span className={styles.historyDetailValue}>{entry.locations}</span>
                                          </div>
                                        )}
                                        {entry.holidaysConsidered !== undefined && (
                                          <div className={styles.historyDetailRow}>
                                            <span className={styles.historyDetailLabel}>Holidays:</span>
                                            <span className={styles.historyDetailValue}>{entry.holidaysConsidered} considered in calculation</span>
                                          </div>
                                        )}
                                      </>
                                    )}
                                    {entry.previousStatus && entry.updatedStatus && (
                                      <div className={styles.historyDetailRow}>
                                        <span className={styles.historyDetailLabel}>Status:</span>
                                        <span className={`${styles.historyDetailValue} ${styles.historyDetailOld}`}>{entry.previousStatus}</span>
                                        <span className={styles.historyDetailValue}> → </span>
                                        <span className={`${styles.historyDetailValue} ${styles.historyDetailNew}`}>{entry.updatedStatus}</span>
                                      </div>
                                    )}
                                  </div>

                                  {entry.recalculatedDates && Object.keys(entry.recalculatedDates).length > 0 && (
                                    <div className={styles.historyRecalcDates}>
                                      <div className={styles.historyRecalcTitle}>
                                        <SyncIcon sx={{ fontSize: "0.75rem" }} />
                                        Recalculated Meeting Dates
                                      </div>
                                      <div className={styles.historyRecalcGrid}>
                                        {Object.entries(entry.recalculatedDates).map(([meeting, date]) => (
                                          <div key={meeting} className={styles.historyRecalcDate}>
                                            <div className={styles.historyRecalcDateMeeting}>{meeting}</div>
                                            <div className={styles.historyRecalcDateValue}>{date}</div>
                                          </div>
                                        ))}
                                      </div>
                                    </div>
                                  )}

                                  <div className={styles.historyDetailRow} style={{ marginTop: "0.5rem", paddingTop: "0.5rem", borderTop: "1px solid #e5e7eb", fontSize: "0.65rem" }}>
                                    <span className={styles.historyDetailLabel}>Updated By:</span>
                                    <span className={styles.historyDetailValue}>{entry.updatedBy} ({entry.updatedByEmail})</span>
                                  </div>
                                </div>
                              </div>
                            ))}
                        </div>
                      )}
                    </div>
                  )}


                  {/* ✅ Feedback View Section — visible only to L&D Site Admin, Super Admin, and New Joinee (NOT Buddy) */}
                  {activeRow?.feedbackJsonHistory && parseFeedbackHistory(activeRow.feedbackJsonHistory).length > 0 && (isNJLogin || isSuperAdmin || isLndAdmin) && !isBuddyLogin && (
                    <div className={styles.feedbackViewSection}>
                      <div className={styles.feedbackViewHeader} onClick={() => setShowFeedbackView(!showFeedbackView)}>
                        <div className={styles.feedbackViewTitle}>
                          <DescriptionOutlinedIcon sx={{ fontSize: "0.9rem" }} />
                          Submitted Feedback
                        </div>
                        {showFeedbackView ? (
                          <ExpandLessIcon className={styles.feedbackViewToggleIcon} />
                        ) : (
                          <ExpandMoreIcon className={`${styles.feedbackViewToggleIcon} ${styles.collapsed}`} />
                        )}
                      </div>

                      {showFeedbackView && (
                        <div className={styles.feedbackViewList}>
                          {parseFeedbackHistory(activeRow.feedbackJsonHistory)
                            .sort((a, b) => new Date(b.submittedAt).getTime() - new Date(a.submittedAt).getTime())
                            .map((entry, idx) => {
                              const questions = getFeedbackQuestionsForEntry(entry, feedbackQuestions);
                              return (
                                <div key={idx} className={styles.feedbackViewEntry}>
                                  <div className={styles.feedbackViewEntryHeader}>
                                    <div className={styles.feedbackMilestoneBadgeSmall}>
                                      {entry.milestone === "3rd" ? "🗓 3rd Milestone" : "🎓 6th Milestone"}
                                    </div>
                                    <div className={styles.feedbackViewEntryMeta}>
                                      <span>{new Date(entry.submittedAt).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" })}</span>
                                      <span className={styles.feedbackMetaSeparator}>•</span>
                                      <span>{new Date(entry.submittedAt).toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" })}</span>
                                    </div>
                                  </div>

                                  <div className={styles.feedbackViewAnswers}>
                                    {questions.map((q) => {
                                      const answer = entry.answers[q.id];
                                      if (!answer) return null;
                                      return (
                                        <div key={q.id} className={styles.feedbackViewAnswer}>
                                          <div className={styles.feedbackViewQuestionRow}>
                                            <span className={styles.feedbackViewQBadge}>Q</span>
                                            <span className={styles.feedbackViewQText}>{q.question}</span>
                                          </div>

                                          <div className={styles.feedbackViewAnswerRow}>
                                            <span className={styles.feedbackViewABadge}>Answer</span>
                                            <span className={styles.feedbackViewAText}>
                                              {q.type === "rating" ? renderFeedbackRatingStars(answer) : answer}
                                            </span>
                                          </div>
                                        </div>
                                      );
                                    })}
                                  </div>

                                  <div style={{ fontSize: "0.65rem", color: "#6b7280", marginTop: "0.5rem", paddingTop: "0.5rem", borderTop: "1px solid #e5e7eb" }}>
                                    Submitted by: <strong>{entry.submittedByEmail}</strong>
                                  </div>
                                </div>
                              );
                            })}
                        </div>
                      )}
                    </div>
                  )}


                  {/* Show when no feedback submitted — visible to all roles including Buddy */}
                  {/* Show when no feedback submitted — hide for Buddy */}
                  {(!isBuddyLogin && (!activeRow?.feedbackJsonHistory || parseFeedbackHistory(activeRow.feedbackJsonHistory).length === 0) && activeRow && ((feedbackMilestone === "3rd" && activeRow.thirdMeetingStatus === "Completed") || (feedbackMilestone === "6th" && activeRow.sixthMeetingStatus === "Completed"))) && (
                    <div className={styles.feedbackNoSubmissions}>
                      📋 No feedback submitted yet
                    </div>
                  )}

                  
                </div>
              ) : (
                activeMeetingSlot && (
                  <div className={styles.singleMeetingDetail}>
                    <div className={styles.smdRow}><span className={styles.smdLabel}>Status</span><SChip s={activeMeetingSlot.status} /></div>

                    {/* Date editing — only for Buddy, not NJ */}
                    {!isNJLogin && isEditingDates ? (
                      <div className={styles.dateEditBlock}>
                        <div className={styles.dateEditHeader}>
                          <CalendarMonthOutlinedIcon sx={{ fontSize: "0.9rem", color: "#047857" }} />
                          <span>Update Scheduled Date</span>
                        </div>
                        <div className={styles.dateEditFields}>
                          <div className={styles.dateEditField}>
                            <label className={styles.dateEditLabel}>Scheduled Date <span style={{ color: "#dc2626" }}>*</span></label>
                            <input type="date" className={`${styles.dateInput} ${dateError ? styles.dateInputError : ""}`}
                              value={editScheduledDate} min={minScheduledDateForSelectedMeeting}
                              onChange={(e) => { const v = e.target.value; setEditScheduledDate(v); setDateError(validateDateSelection(v)); }} />
                            {dateError && (
                              <div className={styles.dateErrorBanner}>
                                <EventBusyIcon sx={{ fontSize: "0.9rem", flexShrink: 0 }} />
                                <span>{dateError}</span>
                              </div>
                            )}
                          </div>
                          <div className={styles.dateEditField}>
                            <label className={styles.dateEditLabel}>Reason for Change <span style={{ color: "#dc2626" }}>*</span></label>
                            <textarea className={`${styles.reasonInput} ${!editReason.trim() && isSavingDates ? styles.reasonInputError : ""}`}
                              placeholder="Explain why the meeting date is being changed (mandatory)…"
                              value={editReason} rows={3}
                              onChange={(e) => setEditReason(e.target.value)} />
                            {!editReason.trim() && <span className={styles.reasonHint}>A reason is required to save the date change.</span>}
                          </div>
                        </div>
                        <div className={styles.dateEditActions}>
                          <button className={styles.cancelBtn} onClick={() => { setIsEditingDates(false); setEditReason(""); setDateError(""); }} disabled={isSavingDates}>Cancel</button>
                          <button className={styles.saveDatesBtn}
                            onClick={() => void handleSaveDates()}
                            disabled={isSavingDates || !editScheduledDate || !!dateError || !editReason.trim()}>
                            {isSavingDates ? <><SyncIcon sx={{ fontSize: "0.85rem", animation: "spin 1s linear infinite" }} /> Saving…</> : <><SaveIcon sx={{ fontSize: "0.85rem" }} /> Save Dates</>}
                          </button>
                        </div>
                      </div>
                    ) : (
                      <>
                        <div className={styles.smdRow}>
                          <span className={styles.smdLabel}>Scheduled Date</span>
                          <div className={styles.smdDateCell}>
                            <span className={styles.smdValue}>{formatDate(activeMeetingSlot.scheduledDate)}</span>
                            {/* Edit dates button only for Buddy */}
                            {!isNJLogin &&
                              !isMeetingLocked(selMeeting) &&
                              canCurrentUserEditSelectedMeeting &&
                              activeMeetingSlot.status !== "Completed" && (
                              <Tooltip title="Edit meeting dates">
                                <button className={styles.editDatesBtn} onClick={enterEditDates} aria-label="Edit dates">
                                  <EditOutlinedIcon sx={{ fontSize: "0.85rem" }} />
                                </button>
                              </Tooltip>
                            )}
                          </div>
                        </div>
                        <div className={styles.smdRow}><span className={styles.smdLabel}>Actual Date</span><span className={styles.smdValue}>{formatDate(activeMeetingSlot.actualDate)}</span></div>
                      </>
                    )}

                    {isMeetingLocked(selMeeting) ? (
                      <div className={styles.lockedBanner}>
                        <LockOutlinedIcon sx={{ fontSize: "1.1rem" }} />
                        <div><div className={styles.lockedBannerTitle}>Meeting Locked</div><div className={styles.lockedBannerDesc}>Complete the previous meeting first to unlock this one.</div></div>
                      </div>
                    ) : (
                      <>
                        {isLoadingQA ? (
                          <div className={styles.qaLoadingWrap}><CircularProgress size={18} sx={{ color: "#047857" }} /><span>Loading questions…</span></div>
                        ) : (
                          // NJ sees NO Q&A section at all
                          isNJLogin ? (
                            isCurrentMeetingCompleted ? (
                              <div className={styles.completedBanner}><CheckCircleIcon sx={{ fontSize: "1rem" }} />This meeting has been completed.</div>
                            ) : (
                              <div className={styles.noQaBanner}><CalendarMonthOutlinedIcon sx={{ fontSize: "1rem", opacity: 0.5 }} />Meeting is scheduled. Your buddy will update the status after the meeting.</div>
                            )
                          ) : currentMeetingQuestions.length > 0 ? (
                            <div className={styles.qaSection}>
                              <div className={styles.qaSectionHeader}>
                                <QuestionAnswerOutlinedIcon sx={{ fontSize: "0.9rem", color: "#047857" }} />
                                <span>Meeting Questions</span>
                                {canCurrentUserEditSelectedMeeting && savedAnswers.some((a) => a.meeting === selMeeting && a.status === "Draft") && !isCurrentMeetingCompleted && (
                                  <Chip label="Draft loaded" size="small" color="warning" sx={{ fontSize: "0.65rem", height: "18px", ml: "auto" }} />
                                )}
                                {isCurrentMeetingCompleted && <Chip label="Answered" size="small" color="success" sx={{ fontSize: "0.65rem", height: "18px", ml: "auto" }} />}
                              </div>
                              <div className={styles.qaList}>
                                {currentMeetingQuestions.map((q, idx) => {
                                  const savedAns    = savedAnswerMap[q.question] || "";
                                  const savedStatus = savedAnswerStatusMap[q.question] || "";
                                  const hasLocalDraft = Object.prototype.hasOwnProperty.call(draftAnswers, String(q.id));
                                  
                                  // ─── Only show draft answers to the person editing (buddy) ─
                                  // If viewing (not editing) and answer is draft, hide it
                                  const shouldHideDraft = !canCurrentUserEditSelectedMeeting && savedStatus === "Draft";
                                  const displayAns = shouldHideDraft ? "" : (hasLocalDraft ? draftAnswers[String(q.id)] : savedAns);
                                  const answerError = getMeetingAnswerValidationMessage(displayAns);
                                  const isAnswered  = savedAns.trim() !== "";
                                  const isDraftSaved = savedStatus === "Draft" && isAnswered && !hasLocalDraft && canCurrentUserEditSelectedMeeting;
                                  return (
                                    <div key={q.id} className={`${styles.qaItem} ${isAnswered && savedStatus === "Completed" && !draftAnswers[String(q.id)] ? styles.qaItemAnswered : ""} ${isDraftSaved ? styles.qaItemDraft : ""}`}>
                                      <div className={styles.qaQuestion}>
                                        <span className={styles.qaQNum}>Q{idx + 1}.</span>
                                        <span className={styles.qaQText}>{q.question}</span>
                                        {savedStatus === "Completed" && !draftAnswers[String(q.id)] && <CheckCircleIcon sx={{ fontSize: "0.85rem", color: "#10b981", ml: "auto", flexShrink: 0 }} />}
                                        {isDraftSaved && <span className={styles.draftBadge}><BookmarkOutlinedIcon sx={{ fontSize: "0.7rem" }} />Draft</span>}
                                      </div>
                                      <textarea
                                        className={`${styles.qaAnswerInput} ${isCurrentMeetingCompleted && !hasLocalDraft ? styles.qaAnswerReadonly : ""}`}
                                        placeholder={isCurrentMeetingCompleted ? "Answer recorded." : !canCurrentUserEditSelectedMeeting ? shouldHideDraft ? "Draft not available for viewing." : "Only assigned buddy can update answers." : savedStatus === "Draft" ? "Draft answer loaded." : "Type your answer here…"}
                                        value={displayAns} readOnly={(isCurrentMeetingCompleted && !hasLocalDraft) || !canCurrentUserEditSelectedMeeting}
                                        rows={3}
                                        onChange={(e) => setDraftAnswers((prev) => ({ ...prev, [String(q.id)]: e.target.value }))}
                                        onBlur={(e) => {
                                          if (!canCurrentUserEditSelectedMeeting || isCurrentMeetingCompleted) return;
                                          const val = normalizeSingleSpaces(e.target.value);
                                          if (val && val.length < MIN_MEETING_ANSWER_LENGTH) {
                                            showSnackbar(
                                              `Question ${idx + 1} validation: Minimum ${MIN_MEETING_ANSWER_LENGTH} characters required (${val.length}/${MIN_MEETING_ANSWER_LENGTH}).`,
                                              "error",
                                              4000
                                            );
                                          }
                                        }}
                                      />
                                      {answerError && <div className={styles.qaFieldError}>{answerError}</div>}
                                    </div>
                                  );
                                })}
                              </div>

                              {canCurrentUserEditSelectedMeeting && !isCurrentMeetingCompleted && (
                                <div className={styles.completeCTAWrap}>
                                  {!allAnswersFilled && (
                                    <div className={styles.completeCTAHint}><WarningAmberIcon sx={{ fontSize: "0.85rem", color: "#f59e0b" }} />Please answer all questions to mark this meeting as complete.</div>
                                  )}
                                  <div className={styles.ctaBtnGroup}>
                                    <Tooltip title={hasInvalidMeetingAnswers ? `Each answer must be at least ${MIN_MEETING_ANSWER_LENGTH} characters.` : "Save answers as draft — you can continue later"}>
                                      <button className={styles.saveDraftBtn} disabled={!hasSomeDraftAnswers || hasInvalidMeetingAnswers || isSavingDraft || isSavingComplete} onClick={() => void handleSaveDraft()}>
                                        {isSavingDraft ? <><SyncIcon sx={{ fontSize: "0.9rem", animation: "spin 1s linear infinite" }} /> Saving…</> : <><BookmarkBorderOutlinedIcon sx={{ fontSize: "0.9rem" }} /> Save Draft</>}
                                      </button>
                                    </Tooltip>
                                    <button className={styles.completeMeetingBtn} disabled={!allAnswersValid || isSavingComplete || isSavingDraft} onClick={() => void handleCompleteMeeting()}>
                                      {isSavingComplete ? <><SyncIcon sx={{ fontSize: "0.9rem", animation: "spin 1s linear infinite" }} /> Saving…</> : <><CheckCircleIcon sx={{ fontSize: "0.9rem" }} /> Mark as Completed</>}
                                    </button>
                                  </div>
                                </div>
                              )}
                              {isCurrentMeetingCompleted && <div className={styles.completedBanner}><CheckCircleIcon sx={{ fontSize: "1rem" }} />This meeting has been completed. Answers are locked.</div>}
                            </div>
                          ) : (
                            <>
                              <div className={styles.noQaBanner}><QuestionAnswerOutlinedIcon sx={{ fontSize: "1rem", opacity: 0.5 }} />No questions configured for {selMeeting}.</div>
                              {!isCurrentMeetingCompleted && canCurrentUserEditSelectedMeeting && (
                                <div className={styles.completeCTAWrap}>
                                  <button className={styles.completeMeetingBtn} disabled={isSavingComplete} onClick={() => void handleCompleteMeeting()}>
                                    {isSavingComplete ? <><SyncIcon sx={{ fontSize: "0.9rem", animation: "spin 1s linear infinite" }} /> Saving…</> : <><CheckCircleIcon sx={{ fontSize: "0.9rem" }} /> Mark as Completed</>}
                                  </button>
                                </div>
                              )}
                              {isCurrentMeetingCompleted && <div className={styles.completedBanner}><CheckCircleIcon sx={{ fontSize: "1rem" }} />This meeting has been completed.</div>}
                            </>
                          )
                        )}
                      </>
                    )}
                  </div>
                )
              )}
            </div>

            <div className={styles.drawerFooter}>
              <button className={styles.cancelBtn} onClick={closeDrawer}>Close</button>
            </div>
          </div>
        )}
      </Drawer>

      {/* ══════════════════════════════════════════════
          REASSIGN BUDDY DRAWER
      ══════════════════════════════════════════════ */}
      <Drawer
        anchor="right"
        open={drawerKind === "reassign"}
        onClose={closeDrawer}
        ModalProps={{ disablePortal: true, keepMounted: true, container: containerRef.current }}
        // sx={{ "& .MuiDrawer-paper": { width: { xs: "100%", sm: "750px" }, boxShadow: "-4px 0 20px rgba(0,0,0,0.15)" } }}
        sx={{
          ...DRAWER_OVER_TOPBAR_SX,
          "& .MuiDrawer-paper": {
            ...DRAWER_OVER_TOPBAR_SX["& .MuiDrawer-paper"],
            width: { xs: "100%", sm: "750px" },
            boxShadow: "-4px 0 20px rgba(0,0,0,0.15)",
          },
        }}>
        {activeRow && (
          <div className={styles.drawer}>
            <div className={styles.drawerHeader}>
              <div>
                <h2 className={styles.drawerTitle}>Reassign Buddy</h2>
                <p className={styles.drawerSubtitle}>{activeRow.njName}</p>
              </div>
              <IconButton onClick={closeDrawer} className={styles.closeBtn}><CloseIcon /></IconButton>
            </div>
            <div className={styles.drawerBody}>
              <div className={styles.editForm}>
                {isAllMeetingsCompleted(activeRow) && (
                  <div style={{ border: "1px solid #fde68a", background: "#fffbeb", color: "#92400e", borderRadius: "8px", padding: "10px 12px", fontSize: "12px", fontWeight: 600 }}>
                    All meetings are completed for this new joinee. Reassignment is not allowed.
                  </div>
                )}
                <div className={styles.editField}>
                  <label className={styles.editLabel}>Current Buddy</label>
                  <div className={styles.editValue}>{activeRow.buddyName}</div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "8px", marginTop: "8px" }}>
                    <div className={styles.editValue}><strong>Global ID:</strong> {activeRow.buddyGlobalId || "—"}</div>
                    <div className={styles.editValue}><strong>Department:</strong> {activeRow.buddyDepartment || "—"}</div>
                    <div className={styles.editValue}><strong>Location:</strong> {activeRow.buddyLocation || "—"}</div>
                  </div>
                </div>

                <div className={styles.editField}>
                  <label className={styles.editLabel}>New Buddy <span className={styles.required}>*</span></label>
                  <div style={{ position: "relative" }}>
                    <input type="text" className={styles.searchBuddyInput}
                      placeholder={isLndAdmin ? `Search buddies in ${activeRow.njLocation || "this location"}…` : "Search by name, email or Global ID…"}
                      value={buddySearchTerm}
                      onChange={(e) => { const v = e.target.value; setBuddySearchTerm(v); setSelectedNewBuddy(null); void handleBuddySearch(v); }}
                      onFocus={() => { if (buddySearchTerm.length >= 2) setShowBuddyDropdown(true); }}
                      onBlur={() => { buddyBlurRef.current = window.setTimeout(() => setShowBuddyDropdown(false), 200); }} />
                    {showBuddyDropdown && buddySearchResults.length > 0 && (
                      <div className={styles.buddyDropdown}>
                        {buddySearchResults.map((b) => (
                          <div key={b.id} className={styles.buddyOption} onMouseDown={(e) => e.preventDefault()}
                            onClick={() => { setSelectedNewBuddy(b); setBuddySearchTerm(b.name); setShowBuddyDropdown(false); setBuddySearchResults([]); }}>
                            <div className={styles.buddyOptionName}>{b.name}</div>
                            <div className={styles.buddyOptionMeta}>{b.department} · {b.location} · {b.email}</div>
                          </div>
                        ))}
                      </div>
                    )}
                    {showBuddyDropdown && buddySearchResults.length === 0 && buddySearchTerm.length >= 2 && (
                      <div className={styles.buddyDropdown}>
                        <div style={{ padding: "0.55rem 0.75rem", fontSize: "0.75rem", color: "#6b7280" }}>
                          {isLndAdmin
                            ? `No active buddies found in "${activeRow.njLocation}" for this search.`
                            : "No matching active buddies found."}
                        </div>
                      </div>
                    )}
                  </div>
                </div>

                {selectedNewBuddy && (
                  <div className={styles.reassignPreview}>
                    <div className={styles.reassignFrom}><span className={styles.reassignChipLabel}>From</span><span className={styles.reassignName}>{activeRow.buddyName}</span></div>
                    <SyncAltIcon sx={{ color: "#00A859", fontSize: "1.2rem" }} />
                    <div className={styles.reassignTo}><span className={styles.reassignChipLabel}>To</span><span className={styles.reassignName}>{selectedNewBuddy.name}</span></div>
                  </div>
                )}
              </div>
            </div>
            <div className={styles.drawerFooter}>
              <button className={styles.cancelBtn} onClick={closeDrawer}>Cancel</button>
              <button
                className={styles.saveBtn}
                onClick={handleReassignSave}
                disabled={!selectedNewBuddy || isSavingReassign || isSameBuddyAsCurrent(activeRow, selectedNewBuddy) || isAllMeetingsCompleted(activeRow) || activeRow.employeeStatus === "Inactive"}
              >
                <SaveIcon sx={{ fontSize: "0.95rem" }} />
                <span>{isSavingReassign ? "Saving…" : "Save Changes"}</span>
              </button>
            </div>
          </div>
        )}
      </Drawer>

      {/* ══════════════════════════════════════════════
          RELEASE NJ DRAWER
      ══════════════════════════════════════════════ */}
      <Drawer
        anchor="right"
        open={drawerKind === "release"}
        onClose={closeDrawer}
        ModalProps={{ disablePortal: true, keepMounted: true, container: containerRef.current }}
        // sx={{ "& .MuiDrawer-paper": { width: { xs: "100%", sm: "750px" }, boxShadow: "-4px 0 20px rgba(0,0,0,0.15)" } }}
        sx={{
          ...DRAWER_OVER_TOPBAR_SX,
          "& .MuiDrawer-paper": {
            ...DRAWER_OVER_TOPBAR_SX["& .MuiDrawer-paper"],
            width: { xs: "100%", sm: "750px" },
            boxShadow: "-4px 0 20px rgba(0,0,0,0.15)",
          },
        }}>
        {activeRow && (
          <div className={styles.drawer}>
            <div className={styles.drawerHeader}>
              <div>
                <h2 className={styles.drawerTitle}>Release New Joinee</h2>
                <p className={styles.drawerSubtitle}>{activeRow.njName}</p>
              </div>
              <IconButton onClick={closeDrawer} className={styles.closeBtn}><CloseIcon /></IconButton>
            </div>
            <div className={styles.drawerBody}>
              <div className={styles.editForm}>
                <div className={styles.formRow3}>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>Global ID</label>
                    <div className={styles.editValue}>{activeRow.njGlobalId || "—"}</div>
                  </div>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>Department</label>
                    <div className={styles.editValue}>{activeRow.njDepartment || "—"}</div>
                  </div>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>Location</label>
                    <div className={styles.editValue}>{activeRow.njLocation || "—"}</div>
                  </div>
                </div>
                <div className={styles.editField}>
                  <label className={styles.editLabel}>Current Status</label>
                  <div className={styles.editValue}>
                    <SChip s={activeRow.employeeStatus} kind="employee" fontSize="0.7rem" />
                  </div>
                </div>
                <div className={styles.editField}>
                  <label className={styles.editLabel}>New Employee Status</label>
                  <div className={styles.toggleRow}>
                    <span className={`${styles.toggleLabel} ${!relActive ? styles.toggleLabelOn : ""}`}>Inactive</span>
                    <Switch checked={relActive} onChange={() => { setRelActive((p) => !p); setRelConfirm(true); }}
                      sx={{ "& .MuiSwitch-switchBase.Mui-checked": { color: "#00A859" }, "& .MuiSwitch-switchBase.Mui-checked + .MuiSwitch-track": { backgroundColor: "#00A859" } }} />
                    <span className={`${styles.toggleLabel} ${relActive ? styles.toggleLabelOn : ""}`}>Active</span>
                  </div>
                </div>
                {relConfirm && (
                  <div style={{
                    border: "1px solid #fde68a", background: "#fffbeb", color: "#92400e",
                    borderRadius: "8px", padding: "10px 12px", fontSize: "12px", fontWeight: 600,
                  }}>
                    {relActive
                      ? `You are about to re-activate "${activeRow.njName}".`
                      : `You are about to release "${activeRow.njName}" and mark them inactive. The BuddyAllocate record will also be set to Inactive.`}
                  </div>
                )}
              </div>
            </div>
            <div className={styles.drawerFooter}>
              <button className={styles.cancelBtn} onClick={closeDrawer}>Cancel</button>
              <button className={styles.saveBtnDanger} onClick={handleReleaseSave} disabled={!relConfirm || isSavingRel}>
                <SaveIcon sx={{ fontSize: "0.95rem" }} />
                <span>{isSavingRel ? "Saving…" : "Confirm"}</span>
              </button>
            </div>
          </div>
        )}
      </Drawer>

      {/* ══════════════════════════════════════════════
          NJ FEEDBACK DRAWER  (3rd / 6th milestone)
      ══════════════════════════════════════════════ */}
      <Drawer
        anchor="right"
        open={drawerKind === "feedback"}
        onClose={closeDrawer}
        ModalProps={{ disablePortal: true, keepMounted: true, container: containerRef.current }}
        // sx={{ "& .MuiDrawer-paper": { width: { xs: "620px !important", sm: "620px !important" }, maxHeight: "100vh", display: "flex", flexDirection: "column", boxShadow: "-4px 0 28px rgba(0,0,0,0.18)" } }}
        sx={{
          ...DRAWER_OVER_TOPBAR_SX,
          "& .MuiDrawer-paper": {
            ...DRAWER_OVER_TOPBAR_SX["& .MuiDrawer-paper"],
            width: { xs: "620px !important", sm: "620px !important" }, maxHeight: "100vh", display: "flex", flexDirection: "column", boxShadow: "-4px 0 28px rgba(0,0,0,0.18)",
          },
        }}>
        {activeRow && (
          <div className={styles.drawer}>
            {/* Header */}
            <div className={`${styles.drawerHeader} ${styles.feedbackDrawerHeader}`}>
              <div className={styles.dHLeft}>
                <div className={styles.feedbackHeaderIcon}>
                  {feedbackMilestone === "3rd"
                    ? <FeedbackOutlinedIcon sx={{ fontSize: "1.2rem" }} />
                    : <StarOutlineIcon sx={{ fontSize: "1.2rem" }} />}
                </div>
                <div>
                  <h2 className={styles.drawerTitle}>
                    {feedbackMilestone === "3rd" ? "Mid-Program Feedback" : "Final Program Feedback"}
                  </h2>
                  <p className={styles.drawerSub}>
                    After <strong>{feedbackMilestone} Meeting</strong> · {activeRow.njName}
                  </p>
                </div>
              </div>
              <IconButton onClick={closeDrawer} className={styles.closeIconBtn}><CloseIcon sx={{ fontSize: "1rem" }} /></IconButton>
            </div>

            {/* Milestone badge */}
            <div className={styles.feedbackMilestoneBanner}>
              <div className={styles.feedbackMilestoneBadge}>
                {feedbackMilestone === "3rd" ? "🗓 Mid-Program — After 3rd Meeting" : "🎓 End of Program — After 6th Meeting"}
              </div>
              <p className={styles.feedbackMilestoneDesc}>
                {feedbackMilestone === "3rd"
                  ? "You've completed the first half of your buddy program. Please share how it has been going so far."
                  : "Congratulations on completing your buddy program! Your feedback helps us improve for future joiners."}
              </p>
            </div>

            {/* Questions */}
            <div className={styles.drawerBody}>
              {isLoadingFeedbackQuestions ? (
                <div className={styles.qaLoadingWrap}>
                  <CircularProgress size={18} sx={{ color: "#047857" }} />
                  <span>Loading feedback questions…</span>
                </div>
              ) : currentFeedbackQuestions.length === 0 ? (
                <div className={styles.noQaBanner}>
                  <FeedbackOutlinedIcon sx={{ fontSize: "1rem", opacity: 0.5 }} />
                  No feedback questions configured for {feedbackMilestone} meeting.
                </div>
              ) : (
                <>
                  <div className={styles.feedbackQuestionList}>
                    {currentFeedbackQuestions.map((q, idx) => (
                      <div key={q.id} className={styles.feedbackQuestionItem}>
                        <div className={styles.feedbackQHeader}>
                          <span className={styles.feedbackQNum}>Q{idx + 1}</span>
                          <span className={styles.feedbackQText}>
                            {normalizeSingleSpaces(q.question)}
                            {q.required && <span className={styles.feedbackQRequired}>*</span>}
                          </span>
                        </div>

                        {/* Yes / No */}
                        {q.type === "yes_no" && (
                          <div className={styles.feedbackRadioGroup}>
                            {(q.options && q.options.length > 0 ? q.options : ["Yes", "No"]).map((opt) => (
                              <label
                                key={normalizeSingleSpaces(opt)}
                                className={`${styles.feedbackRadioOption} ${feedbackAnswers[q.id] === normalizeSingleSpaces(opt) ? styles.feedbackRadioOptionActive : ""}`}>
                                <input
                                  type="radio"
                                  name={`feedback-${q.id}`}
                                  value={normalizeSingleSpaces(opt)}
                                  checked={feedbackAnswers[q.id] === normalizeSingleSpaces(opt)}
                                  onChange={() => setFeedbackAnswers((prev) => ({ ...prev, [q.id]: normalizeSingleSpaces(opt) }))}
                                  className={styles.feedbackRadioInput}
                                />
                                <span className={styles.feedbackRadioDot} />
                                <span>{normalizeSingleSpaces(opt)}</span>
                              </label>
                            ))}
                          </div>
                        )}

                        {/* Agree / Disagree */}
                        {q.type === "agree_disagree" && (
                          <div className={styles.feedbackRadioGroup}>
                            {(q.options && q.options.length > 0
                              ? q.options
                              : ["Strongly Agree", "Agree", "Neutral", "Disagree", "Strongly Disagree"]
                            ).map((opt) => (
                              <label
                                key={normalizeSingleSpaces(opt)}
                                className={`${styles.feedbackRadioOption} ${feedbackAnswers[q.id] === normalizeSingleSpaces(opt) ? styles.feedbackRadioOptionActive : ""}`}>
                                <input
                                  type="radio"
                                  name={`feedback-${q.id}`}
                                  value={normalizeSingleSpaces(opt)}
                                  checked={feedbackAnswers[q.id] === normalizeSingleSpaces(opt)}
                                  onChange={() => setFeedbackAnswers((prev) => ({ ...prev, [q.id]: normalizeSingleSpaces(opt) }))}
                                  className={styles.feedbackRadioInput}
                                />
                                <span className={styles.feedbackRadioDot} />
                                <span>{normalizeSingleSpaces(opt)}</span>
                              </label>
                            ))}
                          </div>
                        )}

                        {/* Rating (1–5 stars) */}
                        {q.type === "rating" && (
                          <div className={styles.feedbackRating}>
                            {(q.options && q.options.length > 0 ? q.options : ["1", "2", "3", "4", "5"]).map((star) => {
                              const selected = Number(feedbackAnswers[q.id] || 0);
                              const value = Number(star);
                              return (
                                <button
                                  key={star}
                                  type="button"
                                  className={`${styles.feedbackStar} ${selected >= value ? styles.feedbackStarActive : ""}`}
                                  onClick={() => setFeedbackAnswers((prev) => ({ ...prev, [q.id]: String(value) }))}
                                  aria-label={`${value} star${value > 1 ? "s" : ""}`}>
                                  ★
                                </button>
                              );
                            })}
                            <span className={styles.feedbackRatingLabel}>
                              {feedbackAnswers[q.id] ? `${feedbackAnswers[q.id]} / 5` : "Select a rating"}
                            </span>
                          </div>
                        )}


                        {/* Remark (free text) */}
                        {q.type === "remark" && (
                          <>
                            <textarea
                              className={styles.feedbackRemarkInput}
                              rows={3}
                              placeholder="Type your remarks here… (optional)"
                              value={feedbackAnswers[q.id] || ""}
                              onChange={(e) => setFeedbackAnswers((prev) => ({ ...prev, [q.id]: e.target.value }))} />
                            {feedbackValidationMessages[q.id] && (
                              <div style={{ color: "#dc2626", fontSize: "0.75rem", marginTop: "0.4rem" }}>
                                {feedbackValidationMessages[q.id]}
                              </div>
                            )}
                          </>
                        )}
                      </div>
                    ))}
                  </div>

                  {/* Required hint */}
                  {!feedbackCanSubmit && (
                    <div className={styles.feedbackRequiredHint}>
                      <WarningAmberIcon sx={{ fontSize: "0.85rem", color: "#f59e0b" }} />
                      Please answer all required (*) questions and meet the minimum character requirement for remarks before submitting.
                    </div>
                  )}
                </>
              )}
            </div>

            {/* Footer */}
            <div className={styles.drawerFooter}>
              <button className={styles.cancelBtn} onClick={closeDrawer}>Cancel</button>
              <button
                className={styles.feedbackSubmitBtn}
                onClick={() => void handleSubmitFeedback()}
                disabled={isLoadingFeedbackQuestions || currentFeedbackQuestions.length === 0 || !feedbackCanSubmit || isSavingFeedback}>
                {isSavingFeedback
                  ? <><SyncIcon sx={{ fontSize: "0.9rem", animation: "spin 1s linear infinite" }} /> Submitting…</>
                  : <><CheckCircleIcon sx={{ fontSize: "0.9rem" }} /> Submit Feedback</>}
              </button>
            </div>
          </div>
        )}
      </Drawer>

    </div>
  );
};

export default MeetingDetails;