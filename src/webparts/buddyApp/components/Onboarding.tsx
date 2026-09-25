/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-void */
/* eslint-disable @rushstack/no-new-null */
/* eslint-disable eqeqeq */
/* eslint-disable max-lines*/
/*eslint-disable prefer-const */
/* eslint-disable  @typescript-eslint/no-unused-vars */
import * as React from "react";
import { useRef, useState, useEffect, useCallback } from "react";
import { DataGrid, GridColDef, GridRenderCellParams } from "@mui/x-data-grid";
import {
  Chip,
  IconButton,
  Tooltip,
  TextField,
  InputAdornment,
  Drawer,
  Switch,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogContentText,
  DialogActions,
  Button,
} from "@mui/material";
import { useSnackbar } from "./Snackbar";
import { BuddyLoader } from "./Buddyloader";
import DeleteOutlineIcon from "@mui/icons-material/DeleteOutline";
import EditOutlinedIcon from "@mui/icons-material/EditOutlined";
import SaveIcon from "@mui/icons-material/Save";
import SearchIcon from "@mui/icons-material/Search";
import CloseIcon from "@mui/icons-material/Close";
import PersonAddAlt1Icon from "@mui/icons-material/PersonAddAlt1";
import CheckCircleOutlineIcon from "@mui/icons-material/CheckCircleOutline";
// import HowToRegIcon from "@mui/icons-material/HowToReg";
import RadioButtonUncheckedIcon from "@mui/icons-material/RadioButtonUnchecked";
import PeopleAltOutlinedIcon from "@mui/icons-material/PeopleAltOutlined";
// import RefreshIcon from "@mui/icons-material/Refresh";
import FileDownloadOutlinedIcon from "@mui/icons-material/FileDownloadOutlined";
import SyncAltIcon from "@mui/icons-material/SyncAlt";
import PersonOffOutlinedIcon from "@mui/icons-material/PersonOffOutlined";
import WarningAmberIcon from "@mui/icons-material/WarningAmber";
import LocationOnOutlinedIcon from "@mui/icons-material/LocationOnOutlined";
import { sp, Web } from "@pnp/sp/presets/all";
import LIST_CONFIG, { getThresholdSafeListItems } from "../../../config/spListConfig";
import styles from "./Onboarding.module.scss";
import type { IBuddyAppProps } from "./IBuddyAppProps";

type SharePointRole =
  | "Super Admin"
  | "Buddy"
  | "L&D Site Admin"
  | "New Joinee"
  | "No-Role"
  | "-Select One-";

type MeetingStatus = "Pending" | "Completed" | "Skipped" | string;
type AssignedDrawerKind = null | "reassign" | "release";

interface Buddy {
  id: string;
  name: string;
  email: string;
  department: string;
  globalId: string;
  location?: string;
}

interface User {
  id: string;
  name: string;
  email: string;
  department: string;
  role: SharePointRole;
  location: string;
  status: "Active" | "Inactive";
  buddyId?: string;
  globalId?: string;
  doj?: string;
  assignStatus?: "Assigned" | "Unassigned";
  isHelpDeskExecutive?: boolean;
  isSuccessStoryApprover?: boolean;
  grade?: string;
  designation?: string;
}

interface NewJoiner {
  id: string;
  name: string;
  email: string;
  department: string;
  location?: string;
  globalId?: string;
  doj?: string;
  status: "Active" | "Inactive";
  buddyId?: string;
  assignStatus?: "Assigned" | "Unassigned";
  grade?: string;
  designation?: string;
}

interface NJRef {
  id: string;
  department: string;
  location: string;
  status: "Active" | "Inactive";
  grade?: string;
  designation?: string;
}

interface BuddyAllocation {
  allocId: string;
  njEmployeeId?: string;
  employeeStatus: "Active" | "Inactive";
  njGlobalId: string;
  njName: string;
  njEmail: string;
  njDepartment: string;
  njLocation: string;
  njGrade?: string;
  njDesignation?: string;
  doj: string;
  buddyName: string;
  buddyEmail: string;
  buddyDepartment: string;
  buddyGlobalId: string;
  buddyLocation: string;
  buddyEmployeeId?: string;
  allocationCount: number;
  allocatedOn: string;
  status: string;
  overallStatus: string;
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
  firstMeetingStatus: MeetingStatus;
  secondMeetingStatus: MeetingStatus;
  thirdMeetingStatus: MeetingStatus;
  fourthMeetingStatus: MeetingStatus;
  fifthMeetingStatus: MeetingStatus;
  sixthMeetingStatus: MeetingStatus;
}

interface HolidayRecord {
  date: Date;
  locations: string[];
}

const meetingStatusColor = (
  s: string,
): "success" | "warning" | "error" | "default" => {
  if (s === "Completed") return "success";
  if (s === "Pending") return "warning";
  if (s === "Skipped") return "error";
  return "default";
};

const GRID_SX = {
  border: "none",
  width: "100%",
  height: "100%",
  "& .MuiDataGrid-columnHeaders": {
    backgroundColor: "#004632",
    color: "#fff",
    fontSize: "0.72rem",
    fontWeight: 700,
    minHeight: "38px !important",
    maxHeight: "38px !important",
  },
  "& .MuiDataGrid-columnHeaderTitle": {
    fontWeight: 700,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
  },
  "& .MuiDataGrid-columnHeaderCheckbox .MuiCheckbox-root": {
    color: "#fff",
    "&.Mui-checked": { color: "#fff" },
  },
  "& .MuiCheckbox-root.Mui-checked": { color: "#00D264" },
  "& .MuiCheckbox-root": { color: "#9ca3af" },
  "& .row-inactive": {
    backgroundColor: "rgba(209,213,219,0.5)",
    opacity: 0.75,
    "&:hover": { opacity: 1, backgroundColor: "rgba(209,213,219,0.7)" },
  },
  "& .MuiDataGrid-row:hover": { backgroundColor: "#e6fff0" },
  "& .MuiDataGrid-cell": {
    fontSize: "0.775rem",
    padding: "0 0.625rem",
    borderColor: "#f3f4f6",
  },
  "& .MuiDataGrid-footerContainer": {
    borderTop: "1px solid #e5e7eb",
    minHeight: "42px",
    maxHeight: "42px",
  },
  "& .MuiTablePagination-root": { fontSize: "0.725rem" },
  "& .MuiTablePagination-selectLabel, & .MuiTablePagination-displayedRows": {
    fontSize: "0.725rem",
  },
};

const inlineActionBtn: React.CSSProperties = {
  display: "inline-flex",
  alignItems: "center",
  gap: "4px",
  border: "1px solid #d1d5db",
  background: "#fff",
  color: "#374151",
  borderRadius: "6px",
  padding: "4px 8px",
  fontSize: "11px",
  fontWeight: 600,
  cursor: "pointer",
  whiteSpace: "nowrap",
};

const SWITCH_SX = {
  "& .MuiSwitch-switchBase.Mui-checked": { color: "#00D264" },
  "& .MuiSwitch-switchBase.Mui-checked + .MuiSwitch-track": {
    backgroundColor: "#00D264",
  },
};

const DRAWER_TOP_OFFSET = 48;

const DRAWER_SX = {
  "& .MuiDrawer-paper": {
    width: { xs: "100%", sm: "750px" },
    top: `${DRAWER_TOP_OFFSET}px`,
    height: `calc(100% - ${DRAWER_TOP_OFFSET}px)`,
    boxShadow: "-4px 0 20px rgba(0,0,0,0.15)",
  },
};

const toSPDate = (d: Date): string => d.toISOString();

const formatDisplayDate = (dateValue?: string): string => {
  if (!dateValue) return "—";
  const d = new Date(dateValue);
  return Number.isNaN(d.getTime())
    ? "—"
    : d.toLocaleDateString("en-IN", {
        day: "2-digit",
        month: "short",
        year: "numeric",
      });
};

const formatExportDate = (dateValue?: string): string => {
  if (!dateValue) return "";
  const d = new Date(dateValue);
  return Number.isNaN(d.getTime())
    ? ""
    : d.toLocaleDateString("en-IN", {
        day: "2-digit",
        month: "short",
        year: "numeric",
      });
};

const mapStatus = (s: string): "Active" | "Inactive" =>
  s === "Active" ? "Active" : "Inactive";

const mapAssignStatus = (s?: string): "Assigned" | "Unassigned" =>
  String(s || "").trim() === "Assigned" ? "Assigned" : "Unassigned";

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

type SimpleGridFilterModel = {
  items?: Array<{ field?: string; operator?: string; value?: any }>;
};

const activeFilterItems = (model?: SimpleGridFilterModel) =>
  (model?.items || []).filter((item) => {
    const op = String(item.operator || "").toLowerCase();
    return Boolean(item.field) && (op.includes("empty") || item.value !== undefined && item.value !== null && String(item.value).trim() !== "");
  });

const hasGridFilter = (model: SimpleGridFilterModel | undefined, fields: string | string[]): boolean => {
  const fieldSet = new Set(Array.isArray(fields) ? fields : [fields]);
  return activeFilterItems(model).some((item) => fieldSet.has(String(item.field)));
};

const rowMatchesGridFilterValue = (rawValue: any, operator: string | undefined, filterValue: any): boolean => {
  const cell = normalizeGridText(rawValue);
  const value = normalizeGridText(filterValue);
  const op = normalizeGridText(operator || "contains");

  if (op.includes("isempty")) return cell === "";
  if (op.includes("isnotempty")) return cell !== "";
  if (!value) return true;
  if (op.includes("not")) return !cell.includes(value);
  if (op.includes("equals") || op === "is" || op === "=") return cell === value;
  if (op.includes("starts")) return cell.startsWith(value);
  if (op.includes("ends")) return cell.endsWith(value);
  return cell.includes(value);
};

const applyClientGridFilters = <T,>(
  rows: T[],
  model: SimpleGridFilterModel | undefined,
  getValue: (row: T, field: string) => any,
): T[] => {
  const items = activeFilterItems(model);
  if (items.length === 0) return rows;

  return rows.filter((row) =>
    items.every((item) => rowMatchesGridFilterValue(getValue(row, String(item.field)), item.operator, item.value)),
  );
};

const ROLE_NEW_JOINER_FILTER = `(Role eq 'New Joinee')`;
const POSITION_TYPE_FILTER = `(PositionType eq 1)`;
const ASSIGN_STATUS_UNASSIGNED_FILTER = `(AssignStatus ne 'Assigned' or AssignStatus eq null or AssignStatus eq '')`;
const ASSIGN_STATUS_ASSIGNED_FILTER = `(AssignStatus eq 'Assigned')`;

const addWorkingDays = (
  startDate: Date,
  workingDays: number,
  holidays: HolidayRecord[],
  relevantLocations: string[],
): Date => {
  const toMidnight = (d: Date): Date => {
    const n = new Date(d);
    n.setHours(0, 0, 0, 0);
    return n;
  };

  const relevantLower = relevantLocations
    .map((l) => l.toLowerCase().trim())
    .filter(Boolean);

  const holidaySet = new Set<string>();
  holidays.forEach((h) => {
    const isGlobal = h.locations.length === 0;
    const matchesLocation =
      isGlobal ||
      h.locations.some((hl) => relevantLower.includes(hl.toLowerCase().trim()));
    if (matchesLocation) {
      holidaySet.add(toMidnight(h.date).toDateString());
    }
  });

  let current = toMidnight(new Date(startDate));
  let added = 0;

  while (added < workingDays) {
    current.setDate(current.getDate() + 1);
    const day = current.getDay();
    const isWeekend = day === 0 || day === 6;
    const isHoliday = holidaySet.has(current.toDateString());
    if (!isWeekend && !isHoliday) added++;
  }

  return current;
};

const exportRowsToCsv = (
  filename: string,
  columns: { key: string; label: string }[],
  rows: Record<string, any>[],
): void => {
  const escapeCell = (value: any): string => {
    const text = value == null ? "" : String(value);
    return `"${text.replace(/"/g, '""')}"`;
  };
  const header = columns.map((c) => escapeCell(c.label)).join(",");
  const body = rows
    .map((row) => columns.map((c) => escapeCell(row[c.key])).join(","))
    .join("\n");
  const blob = new Blob([header + "\n" + body], {
    type: "text/csv;charset=utf-8;",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
};

const isBuddyAvailableForReassignment = (
  buddyGlobalId: string,
  allocations: BuddyAllocation[],
): boolean => {
  const buddyAllocsActive = allocations.filter(
    (alloc) =>
      alloc.buddyGlobalId === buddyGlobalId &&
      alloc.employeeStatus === "Active" &&
      alloc.overallStatus !== "Completed",
  );

  return buddyAllocsActive.length < 4;
};


type BadgeColors = { bg: string; fg: string; border?: string };

const normalizeBadgeValue = (status?: string): string =>
  String(status || "")
    .trim()
    .toLowerCase();

const getOverallStatusColors = (status?: string): BadgeColors => {
  const v = normalizeBadgeValue(status);
  if (v.includes("completed")) {
    return { bg: "#dcfce7", fg: "#166534", border: "#86efac" };
  }
  if (v.includes("assigned")) {
    return {
      bg: "rgba(0, 210, 100, 0.14)",
      fg: "#004632",
      border: "rgba(0, 210, 100, 0.35)",
    };
  }
  if (v.includes("pending") || v.includes("progress")) {
    return { bg: "#fef3c7", fg: "#92400e", border: "#fde68a" };
  }
  if (
    v.includes("inactive") ||
    v.includes("released") ||
    v.includes("cancel")
  ) {
    return { bg: "#fee2e2", fg: "#991b1b", border: "#fecaca" };
  }
  return { bg: "#f3f4f6", fg: "#374151", border: "#e5e7eb" };
};

const getEmpStatusColors = (status?: string): BadgeColors => {
  const v = normalizeBadgeValue(status);
  if (v === "active") {
    return { bg: "#dcfce7", fg: "#166534", border: "#86efac" };
  }
  if (v === "inactive" || v === "no" || v === "disabled") {
    return { bg: "#fee2e2", fg: "#991b1b", border: "#fecaca" };
  }
  return { bg: "#f3f4f6", fg: "#374151", border: "#e5e7eb" };
};

const getCurrentStatusColors = (status?: string): BadgeColors => {
  const v = normalizeBadgeValue(status);
  if (v.includes("completed")) {
    return { bg: "#dcfce7", fg: "#166534", border: "#86efac" };
  }
  if (
    v.includes("pending") ||
    v.includes("upcoming") ||
    v.includes("scheduled")
  ) {
    return { bg: "#fef3c7", fg: "#92400e", border: "#fde68a" };
  }
  if (v.includes("skipped")) {
    return { bg: "#e0f2fe", fg: "#075985", border: "#bae6fd" };
  }
  if (
    v.includes("inactive") ||
    v.includes("released") ||
    v.includes("cancel")
  ) {
    return { bg: "#fee2e2", fg: "#991b1b", border: "#fecaca" };
  }
  return { bg: "#f3f4f6", fg: "#374151", border: "#e5e7eb" };
};

const getRoleColors = (role?: string): BadgeColors => {
  const v = normalizeBadgeValue(role);
  if (v === "super admin") {
    return { bg: "#ede9fe", fg: "#5b21b6", border: "#c4b5fd" };
  }
  if (v === "l&d site admin" || v === "l&d admin" || v === "ld site admin") {
    return { bg: "#fef3c7", fg: "#92400e", border: "#fde68a" };
  }
  if (v === "buddy") {
    return { bg: "#dcfce7", fg: "#166534", border: "#86efac" };
  }
  if (v === "new joinee") {
    return { bg: "#e0f2fe", fg: "#075985", border: "#bae6fd" };
  }
  if (
    !v ||
    v === "-select one-" ||
    v === "no-role" ||
    v === "no role assigned"
  ) {
    return { bg: "#f3f4f6", fg: "#6b7280", border: "#e5e7eb" };
  }
  return { bg: "#f3f4f6", fg: "#374151", border: "#e5e7eb" };
};

const getMeetingStatusColors = (status?: string): BadgeColors => {
  const v = normalizeBadgeValue(status);
  if (v === "completed") {
    return { bg: "#dcfce7", fg: "#166534", border: "#86efac" };
  }
  if (v === "pending") {
    return { bg: "#fef3c7", fg: "#92400e", border: "#fde68a" };
  }
  if (v === "skipped") {
    return { bg: "#e0f2fe", fg: "#075985", border: "#bae6fd" };
  }
  return { bg: "#f3f4f6", fg: "#374151", border: "#e5e7eb" };
};

const statusBadgeSx = (
  colors: BadgeColors,
  fontSize = "0.68rem",
  height = "20px",
) => ({
  fontSize,
  height,
  fontWeight: 700,
  backgroundColor: colors.bg,
  color: colors.fg,
  border: `1px solid ${colors.border || colors.bg}`,
  "& .MuiChip-label": {
    px: "8px",
  },
});

export const Onboarding: React.FC<IBuddyAppProps> = (props) => {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const isMountedRef = useRef(true);
  const buddyBlurTimeout = useRef<number | null>(null);
  const assignedBuddyBlurTimeout = useRef<number | null>(null);
  const { showSnackbar } = useSnackbar();

  const buddiesRef = useRef<Buddy[]>([]);
  const njRefMapRef = useRef<Map<string, NJRef>>(new Map());
  const holidaysRef = useRef<HolidayRecord[]>([]);

  const [activeTab, setActiveTab] = useState<"newJoiner" | "userMgmt">(
    "newJoiner",
  );
  const [searchText, setSearchText] = useState("");
  const [debouncedSearch, setDebouncedSearch] = useState("");
  const [njView, setNjView] = useState<"unassigned" | "assigned">("unassigned");
  const [showInactive, setShowInactive] = useState(false);
  const [njFilterModel, setNjFilterModel] = useState<SimpleGridFilterModel>({ items: [] });
  const [allocFilterModel, setAllocFilterModel] = useState<SimpleGridFilterModel>({ items: [] });
  const [userFilterModel, setUserFilterModel] = useState<SimpleGridFilterModel>({ items: [] });

  const [newJoiners, setNewJoiners] = useState<NewJoiner[]>([]);
  const [selectedNJIds, setSelectedNJIds] = useState<string[]>([]);

  const [njPaginationModel, setNjPaginationModel] = useState({
    page: 0,
    pageSize: 10,
  });
  const [njRowCount, setNjRowCount] = useState(0);
  const [njLoading, setNjLoading] = useState(false);

  const [allocPaginationModel, setAllocPaginationModel] = useState({
    page: 0,
    pageSize: 10,
  });
  const [allocRowCount, setAllocRowCount] = useState(0);

  const [buddyAllocations, setBuddyAllocations] = useState<BuddyAllocation[]>(
    [],
  );
  const [allocLoading, setAllocLoading] = useState(false);

  const [userPaginationModel, setUserPaginationModel] = useState({
    page: 0,
    pageSize: 10,
  });
  const [userRowCount, setUserRowCount] = useState(0);
  const [userLoading, setUserLoading] = useState(false);

  const [unassignedActiveCount, setUnassignedActiveCount] = useState(0);
  const [assignedCount, setAssignedCount] = useState(0);
  const [totalNJCount, setTotalNJCount] = useState(0);
  const [inactiveUnassignedCount, setInactiveUnassignedCount] = useState(0);
  const [inactiveAssignedCount, setInactiveAssignedCount] = useState(0);
  const [activeUserCount, setActiveUserCount] = useState(0);
  const [inactiveUserCount, setInactiveUserCount] = useState(0);
  const [buddyCount, setBuddyCount] = useState(0);
  const [superAdminCount, setSuperAdminCount] = useState(0);
  const [ldAdminCount, setLdAdminCount] = useState(0);
  const [totalUserCount, setTotalUserCount] = useState(0);

  const [buddies, setBuddies] = useState<Buddy[]>([]);
  const [selectedBuddy, setSelectedBuddy] = useState("");
  const [selectedBuddyData, setSelectedBuddyData] = useState<Buddy | null>(
    null,
  );
  const [buddySearchText, setBuddySearchText] = useState("");
  const [buddySearchResults, setBuddySearchResults] = useState<Buddy[]>([]);
  const [showBuddySearch, setShowBuddySearch] = useState(false);
  const [isAssigning, setIsAssigning] = useState(false);

  const [users, setUsers] = useState<User[]>([]);
  const [editDrawerOpen, setEditDrawerOpen] = useState(false);
  const [editingUser, setEditingUser] = useState<User | null>(null);
  const [editRole, setEditRole] = useState<SharePointRole>("-Select One-");
  const [editIsHelpDeskExecutive, setEditIsHelpDeskExecutive] = useState(false);
  const [editIsSuccessStoryApprover, setEditIsSuccessStoryApprover] =
    useState(false);
  const [isSavingRole, setIsSavingRole] = useState(false);
  const [deleteConfirmDialog, setDeleteConfirmDialog] = useState<{
    open: boolean;
    user: User | null;
  }>({ open: false, user: null });
  const [isDeletingUser, setIsDeletingUser] = useState(false);

  const [assignedDrawerKind, setAssignedDrawerKind] =
    useState<AssignedDrawerKind>(null);
  const [activeAssignedRow, setActiveAssignedRow] =
    useState<BuddyAllocation | null>(null);
  const [assignedBuddySearchText, setAssignedBuddySearchText] = useState("");
  const [assignedBuddySearchResults, setAssignedBuddySearchResults] = useState<
    Buddy[]
  >([]);
  const [showAssignedBuddySearch, setShowAssignedBuddySearch] = useState(false);
  const [selectedAssignedBuddy, setSelectedAssignedBuddy] =
    useState<Buddy | null>(null);
  const [isSavingAssignedAction, setIsSavingAssignedAction] = useState(false);
  const [releaseActive, setReleaseActive] = useState(true);
  const [releaseConfirm, setReleaseConfirm] = useState(false);

  const [currentUserRole, setCurrentUserRole] =
    useState<SharePointRole>("-Select One-");
  const [currentUserLocation, setCurrentUserLocation] = useState("");
  const [currentUserResolved, setCurrentUserResolved] = useState(false);
  const [initialDataLoading, setInitialDataLoading] = useState<boolean>(true);

  buddiesRef.current = buddies;


  useEffect(() => {
    const t = setTimeout(() => setDebouncedSearch(searchText), 300);
    return () => clearTimeout(t);
  }, [searchText]);


  useEffect(() => {
    setNjPaginationModel((p) => ({ ...p, page: 0 }));
    setAllocPaginationModel((p) => ({ ...p, page: 0 }));
    setUserPaginationModel((p) => ({ ...p, page: 0 }));
  }, [debouncedSearch, njView, showInactive, njFilterModel, allocFilterModel, userFilterModel]);


  const isSuperAdmin = currentUserRole === "Super Admin";
  const isLdAdmin = currentUserRole === "L&D Site Admin";
  const adminLocLower = (currentUserLocation || "").toLowerCase().trim();


  useEffect(() => {
    if (!props.context) return;
    const bootstrap = async () => {
      try {
        const email =
          props.context?.pageContext?.user?.email ||
          props.context?.pageContext?.legacyPageContext?.userEmail ||
          "";
        if (!email) {
          setCurrentUserResolved(true);
          return;
        }
        const safeEmail = String(email).replace(/'/g, "''");

        // const items = await sp.web.lists
        //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        const items = await getThresholdSafeListItems({
          web: empMasterWeb,
          listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
          select: "ID,Role,Location,FullName",
          filter: `EmailID eq '${safeEmail}' and PositionType eq 1`,
          maxItems: 1,
        });

        if (!isMountedRef.current) return;
        if (items?.[0]) {
          setCurrentUserRole(
            (items[0].Role || "-Select One-") as SharePointRole,
          );
          setCurrentUserLocation(items[0].Location || "");
        }
      } catch (e) {
        console.error("Onboarding: bootstrapCurrentUser", e);
      } finally {
        if (isMountedRef.current) setCurrentUserResolved(true);
      }
    };
    void bootstrap();
  }, [props.context]);


  const escSP = (v: string) => v.replace(/'/g, "''").trim();

  const appendSPGridFilters = (
    parts: string[],
    model: SimpleGridFilterModel,
    fieldMap: Record<string, string>,
  ): void => {
    activeFilterItems(model).forEach((item) => {
      const spField = fieldMap[String(item.field)];
      if (!spField) return;

      const op = normalizeGridText(item.operator || "contains");
      const value = escSP(String(item.value ?? ""));

      if (op.includes("isempty")) {
        parts.push(`(${spField} eq null or ${spField} eq '')`);
        return;
      }

      if (op.includes("isnotempty")) {
        parts.push(`(${spField} ne null and ${spField} ne '')`);
        return;
      }

      if (!value) return;

      if (op.includes("equals") || op === "is" || op === "=") {
        parts.push(`${spField} eq '${value}'`);
      } else if (op.includes("starts")) {
        parts.push(`startswith(${spField}, '${value}')`);
      } else if (op.includes("not")) {
        parts.push(`not substringof('${value}',${spField})`);
      } else {
        parts.push(`substringof('${value}',${spField})`);
      }
    });
  };

  const logGridLoad = (
    label: string,
    details: {
      filter?: string;
      page: number;
      pageSize: number;
      skip: number;
      total: number;
      items: any[];
    },
  ): void => {
    console.groupCollapsed(
      `[${label}] page=${details.page}, pageSize=${details.pageSize}, skip=${details.skip}, rowCount=${details.total}, fetched=${details.items?.length || 0}`,
    );
    console.debug("filter:", details.filter || "<none>");
    console.debug("skipApplied:", details.skip > 0);
    console.debug(
      "ids:",
      (details.items || []).map((x: any) => x.ID).slice(0, 20),
    );
    console.debug("first item:", details.items?.[0]);
    console.debug("raw items:", details.items);
    console.groupEnd();
  };

  /*
  const locClause = useCallback(
    (field: string): string => {
      if (!isLdAdmin || !adminLocLower) return "";
      return ` and ${field} eq '${escSP(adminLocLower)}'`;
    },
    [isLdAdmin, adminLocLower],
  );
*/

  useEffect(() => {
    if (!props.context) return;
    const load = async () => {
      try {
        // const items = await sp.web.lists
        //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        const items = await getThresholdSafeListItems({
          web: empMasterWeb,
          listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
          select: "ID,GlobalID,Department,Location,EmployeeStatus,Grade,Designation",
          filter: `${ROLE_NEW_JOINER_FILTER} and ${POSITION_TYPE_FILTER}`,
        });

        const map = new Map<string, NJRef>();
        (items || []).forEach((i: any) => {
          const gid: string = i.GlobalID || "";
          if (gid) {
            map.set(gid, {
              id: i.ID.toString(),
              department: i.Department || "",
              location: i.Location || "",
              status: mapStatus(i.EmployeeStatus),
              grade: i.Grade || "",
              designation: i.Designation || "",
            });
          }
        });
        njRefMapRef.current = map;
      } catch (e) {
        console.error("loadNJRefMap:", e);
      }
    };
    void load();
  }, [props.context]);


  useEffect(() => {
    if (!props.context) return;
    const load = async () => {
      try {
        // const items = await sp.web.lists
        //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        const items = await getThresholdSafeListItems({
          web: empMasterWeb,
          listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
          select: "ID,Title,FullName,EmailID,Department,GlobalID,Location",
          filter: "EmployeeStatus eq 'Active' and Role eq 'Buddy' and PositionType eq 1",
        });

        if (isMountedRef.current && items) {
          setBuddies(
            items.map((i: any) => ({
              id: i.ID.toString(),
              name: i.FullName || i.Title || "",
              email: i.EmailID || "",
              department: i.Department || "",
              globalId: i.GlobalID || "",
              location: i.Location || "",
            })),
          );
        }
      } catch (e) {
        console.error("loadBuddies:", e);
      }
    };
    void load();
  }, [props.context]);


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
            const raw: string = String(i.Date || "").trim();
            if (!raw) return acc;
            const isoStr = raw.includes(";") ? raw.split(";")[1].trim() : raw;
            const d = new Date(isoStr);
            if (Number.isNaN(d.getTime())) {
              console.warn(
                `NeulandHolidaysList: unparseable Date value "${raw}" — skipped`,
              );
              return acc;
            }
            d.setHours(0, 0, 0, 0);
            let locs: string[] = [];
            if (Array.isArray(i.Location)) {
              locs = (i.Location as any[])
                .map((l) => String(l || "").trim())
                .filter(Boolean);
            } else if (i.Location) {
              const locStr = String(i.Location).trim();
              if (locStr) locs = [locStr];
            }
            acc.push({ date: d, locations: locs });
            return acc;
          },
          [],
        );
        holidaysRef.current = records;
      } catch (e) {
        console.warn("NeulandHolidaysList could not be loaded:", e);
        holidaysRef.current = [];
      }
    };
    void load();
  }, [props.context]);

  const [allowedGrades, setAllowedGrades] = useState<string[]>([]);
  const [userManagementGrades, setUserManagementGrades] = useState<string[]>([]);
  const [allowedGradesLoaded, setAllowedGradesLoaded] = useState<boolean>(false);

  useEffect(() => {
    if (!props.context) return;
    const loadGrades = async () => {
      try {
        const criteriaItems = await getThresholdSafeListItems({
          web: sp.web,
          listTitle: LIST_CONFIG.LISTS.BuddyNewJoineeCriteria,
          select: "ID,Code,Status,ShowinUserManagement",
        });

        const grades = (criteriaItems || [])
          .filter((item: any) => {
            const status = String(item.Status || "").trim().toLowerCase();
            const code = String(item.Code || "").trim().toLowerCase();
            return status === "yes" && /^jm[1-5]$/.test(code);
          })
          .map((item: any) => String(item.Code || "").trim().toUpperCase())
          .filter(Boolean);

        const userMgmtGrades = Array.from(
          new Set(
            (criteriaItems || [])
              .filter((item: any) => {
                const showInUserManagement =
                  item.ShowinUserManagement === true ||
                  String(item.ShowinUserManagement || "")
                    .trim()
                    .toLowerCase() === "yes";
                const code = String(item.Code || "").trim();
                return showInUserManagement && Boolean(code);
              })
              .map((item: any) =>
                String(item.Code || "").trim().toUpperCase(),
              )
              .filter(Boolean),
          ),
        );

        setAllowedGrades(grades);
        setUserManagementGrades(userMgmtGrades);
        setAllowedGradesLoaded(true);
      } catch (err) {
        console.error("Onboarding: Failed to load BuddyNewJoineeCriteria:", err);
        setAllowedGrades([]);
        setUserManagementGrades([]);
        setAllowedGradesLoaded(true);
      }
    };
    void loadGrades();
  }, [props.context]);





  const loadSummaryCounts = useCallback(async () => {
    if (!props.context || !allowedGradesLoaded) return;
    try {
      const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
      const [employeeRows, allocationRows] = await Promise.all([
        getThresholdSafeListItems({
          web: empMasterWeb,
          listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
          select: "ID,Role,EmployeeStatus,AssignStatus,Location,Grade",
          filter: POSITION_TYPE_FILTER,
        }),
        getThresholdSafeListItems({
          web: sp.web,
          listTitle: LIST_CONFIG.LISTS.BuddyAllocate,
          select: "ID,Status,LocationofBuddy",
        }),
      ]);

      const inAdminLocation = (value: any): boolean =>
        !isLdAdmin || !adminLocLower || normalizeGridText(value) === adminLocLower;
      const employeeScope = employeeRows.filter((row: any) => inAdminLocation(row.Location));
      const allocationScope = allocationRows.filter((row: any) => inAdminLocation(row.LocationofBuddy));
      const isAllowedNewJoiner = (row: any): boolean =>
        normalizeGridText(row.Role) === "new joinee" &&
        allowedGrades.includes(String(row.Grade || "").trim().toUpperCase());
      const isAssigned = (row: any): boolean => normalizeGridText(row.AssignStatus) === "assigned";
      const isActive = (row: any): boolean => normalizeGridText(row.EmployeeStatus) === "active";

      const unassignedActive = employeeScope.filter(
        (row: any) => isAllowedNewJoiner(row) && !isAssigned(row) && isActive(row),
      ).length;
      const unassignedInactive = employeeScope.filter(
        (row: any) => isAllowedNewJoiner(row) && !isAssigned(row) && !isActive(row),
      ).length;
      const assigned = allocationScope.filter(
        (row: any) => normalizeGridText(row.Status) === "active",
      ).length;
      const assignedInactive = allocationScope.filter(
        (row: any) => normalizeGridText(row.Status) === "inactive",
      ).length;
      const userManagementScope = employeeScope.filter((row: any) =>
        userManagementGrades.includes(
          String(row.Grade || "").trim().toUpperCase(),
        ),
      );
      const activeUsers = userManagementScope.filter(isActive).length;
      const inactiveUsers = userManagementScope.length - activeUsers;
      const buddyCnt = userManagementScope.filter((row: any) => normalizeGridText(row.Role) === "buddy").length;
      const superAdminCnt = userManagementScope.filter((row: any) => normalizeGridText(row.Role) === "super admin").length;
      const ldAdminCnt = userManagementScope.filter((row: any) => normalizeGridText(row.Role) === "l&d site admin").length;
      const totalUsers = userManagementScope.length;

      if (!isMountedRef.current) return;
      setUnassignedActiveCount(unassignedActive);
      setInactiveUnassignedCount(unassignedInactive);
      setAssignedCount(assigned);
      setInactiveAssignedCount(assignedInactive);
      setTotalNJCount(
        unassignedActive + unassignedInactive + assigned + assignedInactive,
      );
      setActiveUserCount(activeUsers);
      setInactiveUserCount(inactiveUsers);
      setBuddyCount(buddyCnt);
      setSuperAdminCount(superAdminCnt);
      setLdAdminCount(ldAdminCnt);
      setTotalUserCount(totalUsers);
    } catch (e) {
      console.error("loadSummaryCounts:", e);
    }
  }, [
    props.context,
    isLdAdmin,
    adminLocLower,
    allowedGrades,
    userManagementGrades,
    allowedGradesLoaded,
  ]);

  useEffect(() => {
    if (currentUserResolved) {
      void loadSummaryCounts();
    }
  }, [loadSummaryCounts, currentUserResolved]);


  useEffect(() => {
    if (!props.context || !currentUserResolved || !allowedGradesLoaded) return;
    const load = async () => {
      setNjLoading(true);
      try {
        const { page, pageSize } = njPaginationModel;
        const skip = page * pageSize;

        const parts: string[] = [ROLE_NEW_JOINER_FILTER, POSITION_TYPE_FILTER];

        if (njView === "unassigned") {
          parts.push(ASSIGN_STATUS_UNASSIGNED_FILTER);
          // Only codes from BuddyNewJoineeCriteria where Status=Yes AND Code in JM1-JM5
          if (allowedGrades.length > 0) {
            parts.push("(" + allowedGrades.map((g) => `Grade eq '${escSP(g)}'`).join(" or ") + ")");
          } else {
            parts.push("Grade eq '____none____'");
          }
        } else {
          parts.push(ASSIGN_STATUS_ASSIGNED_FILTER);
        }

        const inactiveRequested = isInactiveSearchText(debouncedSearch) || hasGridFilter(njFilterModel, "status");
        if (!showInactive && !inactiveRequested) {
          parts.push(`EmployeeStatus eq 'Active'`);
        }

        if (isLdAdmin && adminLocLower) {
          parts.push(`Location eq '${escSP(adminLocLower)}'`);
        }

        if (debouncedSearch.trim()) {
          const esc = escSP(debouncedSearch);
          parts.push(
            `(substringof('${esc}',FullName) or substringof('${esc}',EmailID) or substringof('${esc}',Department) or substringof('${esc}',Location) or substringof('${esc}',GlobalID) or substringof('${esc}',EmployeeStatus))`,
          );
        }

        appendSPGridFilters(parts, njFilterModel, {
          name: "FullName",
          email: "EmailID",
          department: "Department",
          location: "Location",
          globalId: "GlobalID",
          status: "EmployeeStatus",
        });

        const filter = parts.join(" and ");

        // const njItemsQuery = sp.web.lists
        //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        const njSelect =
          "ID,Title,FullName,EmailID,Department,Location,GlobalID,DateofJoining,EmployeeStatus,Role,AssignStatus,AssignedBuddyGlobalID,Grade,Designation";

        const allItems = await getThresholdSafeListItems({
          web: empMasterWeb,
          listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
          select: njSelect,
          filter,
        });
        const total = allItems.length;
        const items = allItems.slice(skip, skip + pageSize);

        const data: NewJoiner[] = (items || []).map((i: any) => ({
          id: i.ID.toString(),
          name: i.FullName || i.Title || "",
          email: i.EmailID || "",
          department: i.Department || "",
          location: i.Location || "",
          globalId: i.GlobalID || "",
          doj: i.DateofJoining || "",
          status: mapStatus(i.EmployeeStatus),
          buddyId: i.AssignedBuddyGlobalID ? String(i.AssignedBuddyGlobalID) : undefined,
          assignStatus: mapAssignStatus(i.AssignStatus),
          grade: i.Grade || "",
          designation: i.Designation || "",
        }));

        logGridLoad("loadNJ", {
          filter,
          page,
          pageSize,
          skip,
          total,
          items: items || [],
        });

        if (isMountedRef.current) {
          setNewJoiners(data);
          setNjRowCount(total);
        }
      } catch (e) {
        console.error("loadNewJoiners:", e);
        showSnackbar("Failed to load new joiners.", "error", 3000);
      } finally {
        if (isMountedRef.current) setNjLoading(false);
      }
    };
    void load();
  }, [
    props.context,
    currentUserResolved,
    njPaginationModel,
    debouncedSearch,
    njView,
    showInactive,
    njFilterModel,
    isLdAdmin,
    adminLocLower,
    allowedGrades,
    allowedGradesLoaded,
  ]);


  const loadBuddyAllocations = useCallback(async () => {
    if (!props.context || !currentUserResolved) return;
    setAllocLoading(true);
    try {
      const { page, pageSize } = allocPaginationModel;
      const skip = page * pageSize;

      const parts: string[] = [];
      const inactiveRequested = isInactiveSearchText(debouncedSearch) || hasGridFilter(allocFilterModel, ["status", "employeeStatus"]);
      const shouldApplyMappedSearch = isInactiveSearchText(debouncedSearch);

      if (!showInactive && !inactiveRequested) {
        parts.push(`Status eq 'Active'`);
      }

      if (isLdAdmin && adminLocLower) {
        parts.push(`LocationofBuddy eq '${escSP(adminLocLower)}'`);
      }

      if (debouncedSearch.trim() && !shouldApplyMappedSearch) {
        const esc = escSP(debouncedSearch);
        parts.push(
          `(substringof('${esc}',NameOfTheNewJoinerRefColumn) or substringof('${esc}',NJEmailID) or substringof('${esc}',GlobalIDofNJ) or substringof('${esc}',BuddyNameRefChoiceColumn) or substringof('${esc}',BuddyEmailID) or substringof('${esc}',GlobalIDofBuddy) or substringof('${esc}',NJDepartment) or substringof('${esc}',Status) or substringof('${esc}',CurrentMeetingStatus) or substringof('${esc}',OverallStatus))`,
        );
      }

      appendSPGridFilters(parts, allocFilterModel, {
        njName: "NameOfTheNewJoinerRefColumn",
        njEmail: "NJEmailID",
        njGlobalId: "GlobalIDofNJ",
        njDepartment: "NJDepartment",
        buddyName: "BuddyNameRefChoiceColumn",
        buddyEmail: "BuddyEmailID",
        buddyGlobalId: "GlobalIDofBuddy",
        buddyLocation: "LocationofBuddy",
        status: "Status",
        overallStatus: "OverallStatus",
        currentMeetingStatus: "CurrentMeetingStatus",
      });

      const filter = parts.length ? parts.join(" and ") : undefined;

      const allocFields =
        "ID,Title,Status,OverallStatus,CurrentMeetingStatus,AllocationCount," +
        "GlobalIDofBuddy,LocationofBuddy," +
        "BuddyEmailID,BuddyDepartment,BuddyAllocatedOn," +
        "BuddyNameRefChoiceColumn," +
        "GlobalIDofNJ,NJEmailID," +
        "NJDepartment," +
        "NameOfTheNewJoinerRefColumn," +
        "DateofJoining," +
        "FirstInteractionDate,SecondInteractionDate,ThirdInteractionDate," +
        "FourthInteractionDate,FifthInteractionDate,SixthInteractionDate," +
        "FirstActualInteractionDate,SecondActualInteractionDate,ThirdActualInteractionDate," +
        "FourthActualInteractionDate,FifthActualInteractionDate,SixthActualInteractionDate," +
        "FirstMeetingStatus,SecondMeetingStatus,ThirdMeetingStatus," +
        "FourthMeetingStatus,FifthMeetingStatus,SixthMeetingStatus";

      const items = await getThresholdSafeListItems({
        web: sp.web,
        listTitle: LIST_CONFIG.LISTS.BuddyAllocate,
        select: allocFields,
        filter,
        sort: (a: any, b: any) => {
          const byAllocatedOn =
            new Date(b.BuddyAllocatedOn || 0).getTime() -
            new Date(a.BuddyAllocatedOn || 0).getTime();
          return byAllocatedOn || Number(b.ID || 0) - Number(a.ID || 0);
        },
      });

      if (!isMountedRef.current) return;

      const byNjGlobal = new Map<string, NJRef>();
      njRefMapRef.current.forEach((v, k) => byNjGlobal.set(k, v));

      const byBuddyGlobal = new Map<string, Buddy>();
      buddiesRef.current.forEach((b) => {
        if (b.globalId) byBuddyGlobal.set(b.globalId, b);
      });

      let data: BuddyAllocation[] = (items || []).map((i: any) => {
        const njGlobalId = i.GlobalIDofNJ || "";
        const buddyGlobalId = i.GlobalIDofBuddy || "";
        const njRef = byNjGlobal.get(njGlobalId);
        const buddy = byBuddyGlobal.get(buddyGlobalId);

        return {
          allocId: i.ID.toString(),
          njEmployeeId: njRef?.id,
          employeeStatus: njRef?.status === "Active" ? "Active" : "Inactive",
          njGlobalId,
          njName: i.NameOfTheNewJoinerRefColumn || "",
          njEmail: i.NJEmailID || "",
          njDepartment: i.NJDepartment || njRef?.department || "",
          njLocation: njRef?.location || "",
          njGrade: njRef?.grade || "",
          njDesignation: njRef?.designation || "",
          doj: i.DateofJoining || "",
          buddyName: i.BuddyNameRefChoiceColumn || "",
          buddyEmail: i.BuddyEmailID || "",
          buddyDepartment: i.BuddyDepartment || "",
          buddyGlobalId,
          buddyLocation: i.LocationofBuddy || "",
          buddyEmployeeId: buddy?.id,
          allocationCount: i.AllocationCount ?? 1,
          allocatedOn: i.BuddyAllocatedOn || "",
          status: normalizeListStatus(i.Status),
          overallStatus: i.OverallStatus || "—",
          currentMeetingStatus: i.CurrentMeetingStatus || "—",
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
        };
      });

      if (shouldApplyMappedSearch) {
        const q = normalizeGridText(debouncedSearch);
        data = data.filter((row) =>
          [
            row.status,
            row.employeeStatus,
            row.overallStatus,
            row.currentMeetingStatus,
            row.njName,
            row.njEmail,
            row.njGlobalId,
            row.njDepartment,
            row.njLocation,
            row.buddyName,
            row.buddyEmail,
            row.buddyGlobalId,
            row.buddyDepartment,
            row.buddyLocation,
          ].some((value) => normalizeGridText(value).includes(q)),
        );
      }

      data = applyClientGridFilters(data, allocFilterModel, (row, field) => (row as any)[field]);
      const total = data.length;
      const pagedData = data.slice(skip, skip + pageSize);

      logGridLoad("loadBuddyAllocations", {
        filter,
        page,
        pageSize,
        skip,
        total,
        items: pagedData || [],
      });

      setBuddyAllocations(pagedData);
      setAllocRowCount(total);
    } catch (e) {
      console.error("loadBuddyAllocations:", e);
      showSnackbar("Failed to load buddy allocations.", "error", 3000);
    } finally {
      if (isMountedRef.current) setAllocLoading(false);
    }
  }, [
    props.context,
    currentUserResolved,
    allocPaginationModel,
    debouncedSearch,
    showInactive,
    allocFilterModel,
    isLdAdmin,
    adminLocLower,
  ]);

  useEffect(() => {
    void loadBuddyAllocations();
  }, [loadBuddyAllocations]);


  useEffect(() => {
    if (!props.context || !currentUserResolved || !allowedGradesLoaded) return;
    const load = async () => {
      setUserLoading(true);
      try {
        const { page, pageSize } = userPaginationModel;
        const skip = page * pageSize;

        // User Management only shows Employee Master users whose PositionType is 1
        // and whose Grade is enabled in BuddyNewJoineeCriteria.ShowinUserManagement.
        const parts: string[] = [POSITION_TYPE_FILTER];
        if (userManagementGrades.length > 0) {
          parts.push(
            "(" +
              userManagementGrades
                .map((grade) => `Grade eq '${escSP(grade)}'`)
                .join(" or ") +
              ")",
          );
        } else {
          // No criteria row is enabled for User Management, so return no users.
          parts.push("Grade eq '____none____'");
        }

        const inactiveRequested = isInactiveSearchText(debouncedSearch) || hasGridFilter(userFilterModel, "status");
        if (!showInactive && !inactiveRequested) {
          parts.push(`EmployeeStatus eq 'Active'`);
        }

        if (isLdAdmin && adminLocLower) {
          parts.push(`Location eq '${escSP(adminLocLower)}'`);
        }

        if (debouncedSearch.trim()) {
          const esc = escSP(debouncedSearch);
          parts.push(
            `(substringof('${esc}',FullName) or substringof('${esc}',EmailID) or substringof('${esc}',Department) or substringof('${esc}',Role) or substringof('${esc}',Location) or substringof('${esc}',GlobalID) or substringof('${esc}',EmployeeStatus))`,
          );
        }

        appendSPGridFilters(parts, userFilterModel, {
          name: "FullName",
          email: "EmailID",
          department: "Department",
          role: "Role",
          location: "Location",
          globalId: "GlobalID",
          status: "EmployeeStatus",
        });

        const filter = parts.length ? parts.join(" and ") : undefined;

        // const userItemsQuery = sp.web.lists
        //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        const userSelect =
          "ID,Title,FullName,EmailID,Department,Location,GlobalID,DateofJoining,EmployeeStatus,Role,AssignStatus,AssignedBuddyGlobalID,IsHelpDeskExecutive,IsSuccessStoryApprover,Grade,Designation";

        const allItems = await getThresholdSafeListItems({
          web: empMasterWeb,
          listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
          select: userSelect,
          filter,
        });
        const total = allItems.length;
        const items = allItems.slice(skip, skip + pageSize);

        logGridLoad("loadUsers", {
          filter,
          page,
          pageSize,
          skip,
          total,
          items: items || [],
        });

        const data: User[] = (items || []).map((i: any) => ({
          id: i.ID.toString(),
          name: i.FullName || i.Title || "",
          email: i.EmailID || "",
          department: i.Department || "",
          role: (i.Role || "-Select One-") as SharePointRole,
          location: i.Location || "",
          status: mapStatus(i.EmployeeStatus || i.EmployeeStaus || ""),
          buddyId: i.AssignedBuddyGlobalID ? String(i.AssignedBuddyGlobalID) : undefined,
          globalId: i.GlobalID || "",
          doj: i.DateofJoining || "",
          assignStatus: mapAssignStatus(i.AssignStatus),
          isHelpDeskExecutive: i.IsHelpDeskExecutive === "Yes" || i.IsHelpDeskExecutive === true,
          isSuccessStoryApprover:
            i.IsSuccessStoryApprover === "Yes" ||
            i.IsSuccessStoryApprover === true,
          grade: i.Grade || "",
          designation: i.Designation || "",
        }));

        if (isMountedRef.current) {
          setUsers(data);
          setUserRowCount(total);
        }
      } catch (e) {
        console.error("loadUsers:", e);
        showSnackbar("Failed to load users.", "error", 3000);
      } finally {
        if (isMountedRef.current) setUserLoading(false);
      }
    };
    void load();
  }, [
    props.context,
    currentUserResolved,
    userPaginationModel,
    debouncedSearch,
    showInactive,
    userFilterModel,
    isLdAdmin,
    adminLocLower,
    userManagementGrades,
    allowedGradesLoaded,
  ]);


  useEffect(() => {
    const isLoading = allocLoading || njLoading || userLoading;
    if (!isLoading && initialDataLoading) {
      const timer = setTimeout(() => setInitialDataLoading(false), 600);
      return () => clearTimeout(timer);
    }
  }, [allocLoading, njLoading, userLoading, initialDataLoading]);


  useEffect(() => {
    return () => {
      isMountedRef.current = false;
      if (buddyBlurTimeout.current)
        window.clearTimeout(buddyBlurTimeout.current);
      if (assignedBuddyBlurTimeout.current)
        window.clearTimeout(assignedBuddyBlurTimeout.current);
    };
  }, []);

  // Reset IsHelpDeskExecutive to false when role changes
  useEffect(() => {
    if (editDrawerOpen && editRole !== editingUser?.role) {
      setEditIsHelpDeskExecutive(false);
    }
    if (editDrawerOpen && editRole !== "Super Admin") {
      setEditIsSuccessStoryApprover(false);
    }
  }, [editRole, editDrawerOpen, editingUser?.role]);

  const resetBuddySelection = useCallback(() => {
    setSelectedBuddy("");
    setSelectedBuddyData(null);
    setBuddySearchText("");
    setBuddySearchResults([]);
    setShowBuddySearch(false);
  }, []);

  const resetAssignedDrawer = useCallback(() => {
    setAssignedDrawerKind(null);
    setActiveAssignedRow(null);
    setAssignedBuddySearchText("");
    setAssignedBuddySearchResults([]);
    setShowAssignedBuddySearch(false);
    setSelectedAssignedBuddy(null);
    setReleaseActive(true);
    setReleaseConfirm(false);
  }, []);


  const handleTabChange = (tab: "newJoiner" | "userMgmt") => {
    setActiveTab(tab);
    setSearchText("");
    setSelectedNJIds([]);
    resetBuddySelection();
    resetAssignedDrawer();
    setNjView("unassigned");
    setShowInactive(false);
  };

  const handleNjViewChange = (view: "unassigned" | "assigned") => {
    setNjView(view);
    setSelectedNJIds([]);
    resetBuddySelection();
    resetAssignedDrawer();
  };


  const handleBuddySearch = async (
    term: string,
    target: "unassigned" | "assigned" = "unassigned",
    locationScope?: string,
  ) => {
    if (!term || term.trim().length < 2) {
      if (target === "unassigned") {
        setBuddySearchResults([]);
        setShowBuddySearch(true);
      } else {
        setAssignedBuddySearchResults([]);
        setShowAssignedBuddySearch(true);
      }
      return;
    }

    try {
      const esc = escSP(term);
      let locationClause = "";
      if (locationScope && locationScope.trim()) {
        locationClause = ` and Location eq '${escSP(locationScope)}'`;
      }

      const filter =
        `EmployeeStatus eq 'Active' and Role eq 'Buddy' and PositionType eq 1` +
        locationClause +
        ` and (substringof('${esc}',FullName) or substringof('${esc}',EmailID) or substringof('${esc}',GlobalID))`;

      // const items = await sp.web.lists
      //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
      const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
      const items = await getThresholdSafeListItems({
        web: empMasterWeb,
        listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
        select: "ID,Title,FullName,EmailID,Department,GlobalID,Location",
        filter,
        maxItems: 100,
      });

      const mapped: Buddy[] = (items || []).map((i: any) => ({
        id: i.ID.toString(),
        name: i.FullName || i.Title || "",
        email: i.EmailID || "",
        department: i.Department || "",
        globalId: i.GlobalID || "",
        location: i.Location || "",
      }));

      if (!isMountedRef.current) return;

      if (target === "unassigned") {
        setBuddySearchResults(mapped);
        setShowBuddySearch(true);
      } else {
        setAssignedBuddySearchResults(mapped);
        setShowAssignedBuddySearch(true);
      }
    } catch (e) {
      console.error("buddySearch:", e);
      if (!isMountedRef.current) return;
      if (target === "unassigned") setBuddySearchResults([]);
      else setAssignedBuddySearchResults([]);
    }
  };

  const handleBuddySelect = (buddy: Buddy) => {
    setSelectedBuddy(buddy.id);
    setSelectedBuddyData(buddy);
    setBuddySearchText(buddy.name);
    setBuddySearchResults([]);
    setShowBuddySearch(false);
    setBuddies((prev) =>
      prev.some((b) => b.id === buddy.id) ? prev : [...prev, buddy],
    );
  };

  const handleAssignedBuddySelect = (buddy: Buddy) => {
    setSelectedAssignedBuddy(buddy);
    setAssignedBuddySearchText(buddy.name);
    setAssignedBuddySearchResults([]);
    setShowAssignedBuddySearch(false);
    setBuddies((prev) =>
      prev.some((b) => b.id === buddy.id) ? prev : [...prev, buddy],
    );
  };


  const selectedNJLocations: string[] = React.useMemo(() => {
    const locs = selectedNJIds
      .map((id) => newJoiners.find((nj) => nj.id === id)?.location || "")
      .filter(Boolean);
    return [...new Set(locs)];
  }, [selectedNJIds, newJoiners]);

  const hasLocationMismatch = selectedNJLocations.length > 1;
  const selectedNJLocation =
    selectedNJLocations.length === 1 ? selectedNJLocations[0] : "";


  const handleAssignBuddy = async () => {
    if (!selectedBuddy || selectedNJIds.length === 0 || !selectedBuddyData)
      return;

    if (hasLocationMismatch) {
      showSnackbar(
        "Selected New Joiners are from different locations. Please select NJs from the same location to assign a buddy.",
        "error",
        5000,
      );
      return;
    }

    const buddyAllocsForSelectedBuddy = buddyAllocations.filter(
      (alloc) =>
        alloc.buddyGlobalId === selectedBuddyData.globalId &&
        alloc.employeeStatus === "Active" &&
        alloc.overallStatus !== "Completed",
    );

    const availableSlots = 4 - buddyAllocsForSelectedBuddy.length;
    if (availableSlots <= 0) {
      showSnackbar(
        `Buddy "${selectedBuddyData.name}" has reached the maximum allocation limit (4 active new joiners). Complete meetings to free up slots.`,
        "error",
        5000,
      );
      return;
    }

    if (selectedNJIds.length > availableSlots) {
      showSnackbar(
        `Buddy "${selectedBuddyData.name}" has only ${availableSlots} available slot(s). You selected ${selectedNJIds.length} NJs. Please reduce selection or complete meetings to free slots.`,
        "error",
        5000,
      );
      return;
    }

    setIsAssigning(true);

    try {
      const toAssign = newJoiners.filter((nj) => selectedNJIds.includes(nj.id));

      for (const nj of toAssign) {
        if (!nj.doj) {
          console.warn(`Skip ${nj.name}: no DOJ`);
          continue;
        }
        const dojDate = new Date(nj.doj);
        if (Number.isNaN(dojDate.getTime())) {
          console.warn(`Skip ${nj.name}: bad DOJ`);
          continue;
        }

        const holidays = holidaysRef.current;
        const relevantLocations = [
          nj.location || "",
          selectedBuddyData.location || "",
        ].filter(Boolean);

        const d1 = new Date(dojDate);
        d1.setHours(0, 0, 0, 0);

        const d2 = addWorkingDays(dojDate, 15, holidays, relevantLocations);
        const d3 = addWorkingDays(d2, 30, holidays, relevantLocations);
        const d4 = addWorkingDays(d3, 30, holidays, relevantLocations);
        const d5 = addWorkingDays(d4, 30, holidays, relevantLocations);
        const d6 = addWorkingDays(d5, 30, holidays, relevantLocations);


        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        await empMasterWeb.lists
          .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
          .items.getById(Number(nj.id))
          .update({
            AssignStatus: "Assigned",
            AssignedBuddyGlobalID: selectedBuddyData.globalId,
            AssignedBuddyEmail: selectedBuddyData.email || "",
            AssignedBuddyFullName: selectedBuddyData.name || "",
            RunWF: "Yes",
          });

        await sp.web.lists
          .getByTitle(LIST_CONFIG.LISTS.BuddyAllocate)
          .items.add({
            Title: `${selectedBuddyData.name} - ${nj.name}`,
            GlobalIDofBuddy: selectedBuddyData.globalId || "",
            LocationofBuddy: selectedBuddyData.location || "",
            GlobalIDofNJ: nj.globalId || "",
            NJDepartment: nj.department || "",
            DateofJoining: toSPDate(dojDate),
            FirstInteractionDate: toSPDate(d1),
            SecondInteractionDate: toSPDate(d2),
            ThirdInteractionDate: toSPDate(d3),
            FourthInteractionDate: toSPDate(d4),
            FifthInteractionDate: toSPDate(d5),
            SixthInteractionDate: toSPDate(d6),
            AllocationCount: 1,
            BuddyAllocatedOn: toSPDate(new Date()),
            Status: "Active",
            BuddyEmailID: selectedBuddyData.email || "",
            BuddyNameRefChoiceColumn: selectedBuddyData.name || "",
            NameOfTheNewJoinerRefColumn: nj.name || "",
            BuddyDepartment: selectedBuddyData.department || "",
            NJEmailID: nj.email || "",
            FirstMeetingStatus: "Pending",
            SecondMeetingStatus: "Pending",
            ThirdMeetingStatus: "Pending",
            FourthMeetingStatus: "Pending",
            FifthMeetingStatus: "Pending",
            SixthMeetingStatus: "Pending",
            OverallStatus: "Assigned",
            CurrentMeetingStatus: "1st Meeting Pending",
            UpcomingMeetingDate: toSPDate(d1),
            RunWF: "Yes",
          });
      }

      if (!isMountedRef.current) return;

      setSelectedNJIds([]);
      resetBuddySelection();
      await Promise.all([loadBuddyAllocations(), loadSummaryCounts()]);
      setNjPaginationModel((p) => ({ ...p, page: 0 }));
      showSnackbar(
        `Buddy assigned to ${toAssign.length} new joiner(s)!`,
        "success",
        4000,
      );
    } catch (error) {
      console.error("Error assigning buddy:", error);
      showSnackbar("Failed to assign buddy. Please try again.", "error", 4000);
    } finally {
      if (isMountedRef.current) setIsAssigning(false);
    }
  };

  const handleClearSelection = () => {
    setSelectedNJIds([]);
    resetBuddySelection();
  };


  const openAssignedReassign = (row: BuddyAllocation) => {
    setActiveAssignedRow(row);
    setAssignedDrawerKind("reassign");
    setAssignedBuddySearchText("");
    setAssignedBuddySearchResults([]);
    setShowAssignedBuddySearch(false);
    setSelectedAssignedBuddy(null);
  };

  const openAssignedRelease = (row: BuddyAllocation) => {
    setActiveAssignedRow(row);
    setAssignedDrawerKind("release");
    setReleaseActive(row.employeeStatus === "Active");
    setReleaseConfirm(false);
  };

  const handleAssignedReassignSave = async () => {
    if (!activeAssignedRow || !selectedAssignedBuddy) return;

    if (activeAssignedRow.buddyGlobalId === selectedAssignedBuddy.globalId) {
      showSnackbar(
        `Cannot reassign to the same buddy. Please select a different buddy.`,
        "error",
        4000,
      );
      return;
    }

    if (
      (isLdAdmin || isSuperAdmin) &&
      activeAssignedRow.njLocation &&
      selectedAssignedBuddy.location
    ) {
      if (
        selectedAssignedBuddy.location.toLowerCase().trim() !==
        activeAssignedRow.njLocation.toLowerCase().trim()
      ) {
        showSnackbar(
          `Cannot reassign: Selected buddy is in "${selectedAssignedBuddy.location}" but NJ is in "${activeAssignedRow.njLocation}". Location must match.`,
          "error",
          5000,
        );
        return;
      }
    }

    if (
      !isBuddyAvailableForReassignment(
        selectedAssignedBuddy.globalId,
        buddyAllocations,
      )
    ) {
      showSnackbar(
        `Buddy "${selectedAssignedBuddy.name}" has reached the maximum allocation limit (4 active new joiners). To reassign this buddy, at least one of their existing assignments must have ALL 6 meetings completed.`,
        "error",
        5000,
      );
      return;
    }

    setIsSavingAssignedAction(true);
    try {
      await sp.web.lists
        .getByTitle(LIST_CONFIG.LISTS.BuddyAllocate)
        .items.getById(Number(activeAssignedRow.allocId))
        .update({
          BuddyNameRefChoiceColumn: selectedAssignedBuddy.name,
          BuddyEmailID: selectedAssignedBuddy.email,
          GlobalIDofBuddy: selectedAssignedBuddy.globalId,
          LocationofBuddy: selectedAssignedBuddy.location || "",
          BuddyDepartment: selectedAssignedBuddy.department || "",
          RunWF: "Yes",
        });

      if (activeAssignedRow.njEmployeeId) {
        // await sp.web.lists
        //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        await empMasterWeb.lists
          .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
          .items.getById(Number(activeAssignedRow.njEmployeeId))
          .update({
            AssignedBuddyGlobalID: selectedAssignedBuddy.globalId,
            AssignedBuddyFullName: selectedAssignedBuddy.name,
            AssignedBuddyEmail: selectedAssignedBuddy.email,
            ReassignBuddyStatus: true,
            RunWF: "Yes",
          });
      }

      await loadBuddyAllocations();
      setUsers((prev) =>
        prev.map((u) =>
          u.id === activeAssignedRow?.njEmployeeId
            ? {
                ...u,
                buddyId: selectedAssignedBuddy.id,
                assignStatus: "Assigned",
              }
            : u,
        ),
      );

      showSnackbar(
        `Buddy reassigned to ${selectedAssignedBuddy.name}.`,
        "success",
        3000,
      );
      resetAssignedDrawer();
    } catch (e) {
      console.error("assignedReassignSave:", e);
      showSnackbar("Failed to reassign buddy.", "error", 4000);
    } finally {
      if (isMountedRef.current) setIsSavingAssignedAction(false);
    }
  };

  const handleAssignedReleaseSave = async () => {
    if (!activeAssignedRow) return;

    const employeeId =
      activeAssignedRow.njEmployeeId ??
      njRefMapRef.current.get(activeAssignedRow.njGlobalId)?.id;

    if (!employeeId) {
      showSnackbar(
        "Employee record not found for this NJ. Please refresh and try again.",
        "error",
        4000,
      );
      return;
    }

    setIsSavingAssignedAction(true);
    const nextEmployeeStatus = releaseActive ? "Active" : "Inactive";
    const isRelease = !releaseActive;

    try {
      // await sp.web.lists
      //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
      const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
      await empMasterWeb.lists
        .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        .items.getById(Number(employeeId))
        .update({ EmployeeStatus: nextEmployeeStatus, RunWF: "Yes" });

      await sp.web.lists
        .getByTitle(LIST_CONFIG.LISTS.BuddyAllocate)
        .items.getById(Number(activeAssignedRow.allocId))
        .update({ Status: isRelease ? "Inactive" : "Active", RunWF: "Yes" });

      setUsers((prev) =>
        prev.map((u) =>
          u.id === employeeId
            ? {
                ...u,
                status: nextEmployeeStatus === "Active" ? "Active" : "Inactive",
              }
            : u,
        ),
      );
      setNewJoiners((prev) =>
        prev.map((nj) =>
          nj.id === employeeId
            ? {
                ...nj,
                status: nextEmployeeStatus === "Active" ? "Active" : "Inactive",
              }
            : nj,
        ),
      );

      await Promise.all([loadBuddyAllocations(), loadSummaryCounts()]);

      showSnackbar(
        releaseActive
          ? `${activeAssignedRow.njName} re-activated successfully.`
          : `${activeAssignedRow.njName} released and marked inactive.`,
        releaseActive ? "success" : "warning",
        3000,
      );
      resetAssignedDrawer();
    } catch (e) {
      console.error("assignedReleaseSave:", e);
      showSnackbar("Failed to update employee status.", "error", 4000);
    } finally {
      if (isMountedRef.current) setIsSavingAssignedAction(false);
    }
  };


  const openEditDrawer = (user: User) => {
    if (isLdAdmin && user.role === "Super Admin") {
      showSnackbar(
        "L&D Site Admin cannot modify Super Admin.",
        "error",
        3000,
      );
      return;
    }
    setEditingUser(user);
    setEditRole(user.role);
    // Default to false if IsHelpDeskExecutive is not explicitly set
    setEditIsHelpDeskExecutive(user.isHelpDeskExecutive === true ? true : false);
    setEditIsSuccessStoryApprover(
      user.isSuccessStoryApprover === true ? true : false,
    );
    setEditDrawerOpen(true);
  };

  const handleSaveRole = async () => {
    if (!editingUser || editRole === "-Select One-") {
      showSnackbar("Please select a valid role.", "error", 3000);
      return;
    }

    const nextRole: SharePointRole =
      editRole === "No-Role" ? "-Select One-" : editRole;
    const roleChanged = nextRole !== editingUser.role;

    if (isLdAdmin && editingUser.role === "Super Admin") {
      showSnackbar(
        "L&D Site Admin cannot modify Super Admin role.",
        "error",
        3000,
      );
      return;
    }

    if (
      !roleChanged &&
      editIsHelpDeskExecutive ===
        (editingUser.isHelpDeskExecutive || false) &&
      editIsSuccessStoryApprover ===
        (editingUser.isSuccessStoryApprover || false)
    ) {
      showSnackbar(
        `"${editingUser.name}" already has the same role and access settings. Please make a change.`,
        "warning",
        4000,
      );
      return;
    }

    setIsSavingRole(true);
    try {
      const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
      
      // Prepare update object
      const updateObj: any = {};
      
      // Update role if it changed
      if (roleChanged) {
        updateObj.Role = editRole === "No-Role" ? null : editRole;
        updateObj.RunWF = "Yes";
      }
      
      // Update IsHelpDeskExecutive if it changed
      if (editIsHelpDeskExecutive !== (editingUser.isHelpDeskExecutive || false)) {
        updateObj.IsHelpDeskExecutive = editIsHelpDeskExecutive;
      }

      if (
        editIsSuccessStoryApprover !==
        (editingUser.isSuccessStoryApprover || false)
      ) {
        updateObj.IsSuccessStoryApprover = editIsSuccessStoryApprover;
      }

      await empMasterWeb.lists
        .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        .items.getById(Number(editingUser.id))
        .update(updateObj);

      if (!isMountedRef.current) return;
      setUsers((prev) =>
        prev.map((u) =>
          u.id === editingUser.id
            ? {
                ...u,
                role: nextRole,
                isHelpDeskExecutive: editIsHelpDeskExecutive,
                isSuccessStoryApprover: editIsSuccessStoryApprover,
              }
            : u,
        ),
      );

      if (
        editingUser.email?.toLowerCase() ===
        props.userloginDetails?.email?.toLowerCase()
      ) {
        (window as any).refreshBuddyApp?.();
      }

      void loadSummaryCounts();
      setEditDrawerOpen(false);
      setEditingUser(null);
      showSnackbar(
        editRole === "No-Role"
          ? `The role assigned to "${editingUser.name}" has been removed successfully.`
          : `The role for "${editingUser.name}" has been updated to "${editRole}" successfully.`,
        "success",
        3000,
      );
    } catch (e) {
      console.error("saveRole:", e);
      showSnackbar("Failed to update role. Please try again.", "error", 4000);
    } finally {
      if (isMountedRef.current) setIsSavingRole(false);
    }
  };
/*
  const handleActivateUser = async (user: User) => {
    if (isLdAdmin && user.role === "Super Admin") {
      showSnackbar(
        "L&D Site Admin cannot activate Super Admin.",
        "error",
        3000,
      );
      return;
    }

    try {
      const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
      await empMasterWeb.lists
        .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        .items.getById(Number(user.id))
        .update({ EmployeeStatus: "Active" });

      if (!isMountedRef.current) return;
      setUsers((prev) =>
        prev.map((item) =>
          item.id === user.id ? { ...item, status: "Active" } : item,
        ),
      );
      void loadSummaryCounts();
      showSnackbar(
        `User "${user.name}" has been activated successfully.`,
        "success",
        3000,
      );
    } catch (e) {
      console.error("handleActivateUser:", e);
      showSnackbar(
        "Failed to activate user. Please try again.",
        "error",
        4000,
      );
    }
  };*/

  const handleDeleteUser = (user: User) => {
    if (isLdAdmin && user.role === "Super Admin") {
      showSnackbar(
        "L&D Site Admin cannot deactivate Super Admin.",
        "error",
        3000,
      );
      return;
    }
    setDeleteConfirmDialog({ open: true, user });
  };

  const handleConfirmDelete = async () => {
    if (!deleteConfirmDialog.user) return;
    const userToDelete = deleteConfirmDialog.user;

    if (isLdAdmin && userToDelete.role === "Super Admin") {
      showSnackbar(
        "L&D Site Admin cannot deactivate Super Admin.",
        "error",
        3000,
      );
      setDeleteConfirmDialog({ open: false, user: null });
      return;
    }

    setIsDeletingUser(true);
    try {
      // await sp.web.lists
      //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
      const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
      await empMasterWeb.lists
        .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        .items.getById(Number(userToDelete.id))
        .update({ EmployeeStatus: "Inactive" });

      if (!isMountedRef.current) return;
      setUsers((prev) =>
        prev.map((u) =>
          u.id === userToDelete.id ? { ...u, status: "Inactive" } : u,
        ),
      );
      void loadSummaryCounts();
      setDeleteConfirmDialog({ open: false, user: null });
      showSnackbar(
        `User "${userToDelete.name}" has been deactivated successfully!`,
        "success",
        3000,
      );
    } catch (e) {
      console.error("handleConfirmDelete:", e);
      showSnackbar(
        "Failed to deactivate user. Please try again.",
        "error",
        4000,
      );
    } finally {
      if (isMountedRef.current) setIsDeletingUser(false);
    }
  };

  const handleCancelDelete = () => {
    setDeleteConfirmDialog({ open: false, user: null });
  };


  const inactiveCount =
    activeTab === "newJoiner"
      ? njView === "assigned"
        ? inactiveAssignedCount
        : inactiveUnassignedCount
      : inactiveUserCount;


  const handleExportNewJoiner = async () => {
    try {
      showSnackbar("Preparing export…", "info", 2000);

      if (njView === "unassigned") {
        const parts: string[] = [
          ROLE_NEW_JOINER_FILTER,
          POSITION_TYPE_FILTER,
          ASSIGN_STATUS_UNASSIGNED_FILTER,
        ];
        // Only codes from BuddyNewJoineeCriteria where Status=Yes AND Code in JM1-JM5
        if (allowedGrades.length > 0) {
          parts.push("(" + allowedGrades.map((g) => `Grade eq '${escSP(g)}'`).join(" or ") + ")");
        } else {
          parts.push("Grade eq '____none____'");
        }
        if (!showInactive) parts.push(`EmployeeStatus eq 'Active'`);
        if (isLdAdmin && adminLocLower)
          parts.push(`Location eq '${escSP(adminLocLower)}'`);
        if (debouncedSearch.trim()) {
          const esc = escSP(debouncedSearch);
          parts.push(
            `(substringof('${esc}',FullName) or substringof('${esc}',EmailID) or substringof('${esc}',Department) or substringof('${esc}',Location) or substringof('${esc}',GlobalID))`,
          );
        }

        // const items = await sp.web.lists
        //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        const items = await getThresholdSafeListItems({
          web: empMasterWeb,
          listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
          select: "ID,Title,FullName,EmailID,Department,Location,GlobalID,DateofJoining,EmployeeStatus,AssignStatus,Grade,Designation",
          filter: parts.join(" and "),
        });

        exportRowsToCsv(
          "new_joiner_unassigned.csv",
          [
            { key: "name", label: "Name" },
            { key: "email", label: "Email" },
            { key: "department", label: "Department" },
            { key: "location", label: "Location" },
            { key: "globalId", label: "Global ID" },
            { key: "grade", label: "Grade" },
            { key: "designation", label: "Designation" },
            { key: "doj", label: "Date of Joining" },
            { key: "status", label: "Status" },
            { key: "assignStatus", label: "Assign Status" },
          ],
          (items || []).map((i: any) => ({
            name: i.FullName || i.Title || "",
            email: i.EmailID || "",
            department: i.Department || "",
            location: i.Location || "",
            globalId: i.GlobalID || "",
            grade: i.Grade || "",
            designation: i.Designation || "",
            doj: formatExportDate(i.DateofJoining),
            status: mapStatus(i.EmployeeStatus),
            assignStatus: mapAssignStatus(i.AssignStatus),
          })),
        );
      } else {
        const parts: string[] = [];
        if (!showInactive) parts.push(`Status eq 'Active'`);
        if (isLdAdmin && adminLocLower)
          parts.push(`LocationofBuddy eq '${escSP(adminLocLower)}'`);
        if (debouncedSearch.trim()) {
          const esc = escSP(debouncedSearch);
          parts.push(
            `(substringof('${esc}',NameOfTheNewJoinerRefColumn) or substringof('${esc}',NJEmailID) or substringof('${esc}',GlobalIDofNJ) or substringof('${esc}',BuddyNameRefChoiceColumn) or substringof('${esc}',BuddyEmailID) or substringof('${esc}',GlobalIDofBuddy))`,
          );
        }

        const items = await getThresholdSafeListItems({
          web: sp.web,
          listTitle: LIST_CONFIG.LISTS.BuddyAllocate,
          select:
            "ID,Title,Status,OverallStatus,CurrentMeetingStatus,AllocationCount," +
            "GlobalIDofBuddy,LocationofBuddy,BuddyEmailID,BuddyDepartment,BuddyAllocatedOn," +
            "BuddyNameRefChoiceColumn,GlobalIDofNJ,NJEmailID,NJDepartment," +
            "NameOfTheNewJoinerRefColumn,DateofJoining," +
            "FirstInteractionDate,SecondInteractionDate,ThirdInteractionDate," +
            "FourthInteractionDate,FifthInteractionDate,SixthInteractionDate," +
            "FirstActualInteractionDate,SecondActualInteractionDate,ThirdActualInteractionDate," +
            "FourthActualInteractionDate,FifthActualInteractionDate,SixthActualInteractionDate," +
            "FirstMeetingStatus,SecondMeetingStatus,ThirdMeetingStatus," +
            "FourthMeetingStatus,FifthMeetingStatus,SixthMeetingStatus",
          filter: parts.length ? parts.join(" and ") : undefined,
          sort: (a: any, b: any) => {
            const byAllocatedOn =
              new Date(b.BuddyAllocatedOn || 0).getTime() -
              new Date(a.BuddyAllocatedOn || 0).getTime();
            return byAllocatedOn || Number(b.ID || 0) - Number(a.ID || 0);
          },
        });

        exportRowsToCsv(
          "new_joiner_assigned.csv",
          [
            { key: "njName", label: "NJ Name" },
            { key: "njEmail", label: "NJ Email" },
            { key: "njGlobalId", label: "Global ID (NJ)" },
            { key: "njDepartment", label: "NJ Department" },
            { key: "njGrade", label: "Grade" },
            { key: "njDesignation", label: "Designation" },
            { key: "doj", label: "Date of Joining" },
            { key: "buddyName", label: "Buddy Name" },
            { key: "buddyEmail", label: "Buddy Email" },
            { key: "buddyGlobalId", label: "Global ID (Buddy)" },
            { key: "buddyDepartment", label: "Buddy Department" },
            { key: "buddyLocation", label: "Buddy Location" },
            { key: "allocatedOn", label: "Allocated On" },
            { key: "allocationCount", label: "Allocation Count" },
            { key: "overallStatus", label: "Overall Status" },
            { key: "currentMeetingStatus", label: "Current Meeting Status" },
            { key: "firstInteractionDate", label: "1st Interaction Date" },
            { key: "secondInteractionDate", label: "2nd Interaction Date" },
            { key: "thirdInteractionDate", label: "3rd Interaction Date" },
            { key: "fourthInteractionDate", label: "4th Interaction Date" },
            { key: "fifthInteractionDate", label: "5th Interaction Date" },
            { key: "sixthInteractionDate", label: "6th Interaction Date" },
            { key: "firstActualInteractionDate", label: "M1 Actual Date" },
            { key: "secondActualInteractionDate", label: "M2 Actual Date" },
            { key: "thirdActualInteractionDate", label: "M3 Actual Date" },
            { key: "fourthActualInteractionDate", label: "M4 Actual Date" },
            { key: "fifthActualInteractionDate", label: "M5 Actual Date" },
            { key: "sixthActualInteractionDate", label: "M6 Actual Date" },
            { key: "firstMeetingStatus", label: "M1 Status" },
            { key: "secondMeetingStatus", label: "M2 Status" },
            { key: "thirdMeetingStatus", label: "M3 Status" },
            { key: "fourthMeetingStatus", label: "M4 Status" },
            { key: "fifthMeetingStatus", label: "M5 Status" },
            { key: "sixthMeetingStatus", label: "M6 Status" },
          ],
          (items || []).map((i: any) => ({
            njName: i.NameOfTheNewJoinerRefColumn || "",
            njEmail: i.NJEmailID || "",
            njGlobalId: i.GlobalIDofNJ || "",
            njDepartment: i.NJDepartment || "",
            njGrade: njRefMapRef.current.get(i.GlobalIDofNJ || "")?.grade || "",
            njDesignation: njRefMapRef.current.get(i.GlobalIDofNJ || "")?.designation || "",
            doj: formatExportDate(i.DateofJoining),
            buddyName: i.BuddyNameRefChoiceColumn || "",
            buddyEmail: i.BuddyEmailID || "",
            buddyGlobalId: i.GlobalIDofBuddy || "",
            buddyDepartment: i.BuddyDepartment || "",
            buddyLocation: i.LocationofBuddy || "",
            allocatedOn: formatExportDate(i.BuddyAllocatedOn),
            allocationCount: i.AllocationCount ?? 1,
            overallStatus: i.OverallStatus || "",
            currentMeetingStatus: i.CurrentMeetingStatus || "",
            firstInteractionDate: formatExportDate(i.FirstInteractionDate),
            secondInteractionDate: formatExportDate(i.SecondInteractionDate),
            thirdInteractionDate: formatExportDate(i.ThirdInteractionDate),
            fourthInteractionDate: formatExportDate(i.FourthInteractionDate),
            fifthInteractionDate: formatExportDate(i.FifthInteractionDate),
            sixthInteractionDate: formatExportDate(i.SixthInteractionDate),
            firstActualInteractionDate: formatExportDate(
              i.FirstActualInteractionDate,
            ),
            secondActualInteractionDate: formatExportDate(
              i.SecondActualInteractionDate,
            ),
            thirdActualInteractionDate: formatExportDate(
              i.ThirdActualInteractionDate,
            ),
            fourthActualInteractionDate: formatExportDate(
              i.FourthActualInteractionDate,
            ),
            fifthActualInteractionDate: formatExportDate(
              i.FifthActualInteractionDate,
            ),
            sixthActualInteractionDate: formatExportDate(
              i.SixthActualInteractionDate,
            ),
            firstMeetingStatus: i.FirstMeetingStatus || "",
            secondMeetingStatus: i.SecondMeetingStatus || "",
            thirdMeetingStatus: i.ThirdMeetingStatus || "",
            fourthMeetingStatus: i.FourthMeetingStatus || "",
            fifthMeetingStatus: i.FifthMeetingStatus || "",
            sixthMeetingStatus: i.SixthMeetingStatus || "",
          })),
        );
      }
    } catch (e) {
      console.error("exportNewJoiner:", e);
      showSnackbar("Export failed. Please try again.", "error", 3000);
    }
  };

  const handleExportUsers = async () => {
    try {
      showSnackbar("Preparing export…", "info", 2000);

      // Export follows the same rules as the User Management grid.
      const parts: string[] = [POSITION_TYPE_FILTER];
      if (userManagementGrades.length > 0) {
        parts.push(
          "(" +
            userManagementGrades
              .map((grade) => `Grade eq '${escSP(grade)}'`)
              .join(" or ") +
            ")",
        );
      } else {
        parts.push("Grade eq '____none____'");
      }
      if (!showInactive) parts.push(`EmployeeStatus eq 'Active'`);
      if (isLdAdmin && adminLocLower)
        parts.push(`Location eq '${escSP(adminLocLower)}'`);
      if (debouncedSearch.trim()) {
        const esc = escSP(debouncedSearch);
        parts.push(
          `(substringof('${esc}',FullName) or substringof('${esc}',EmailID) or substringof('${esc}',Department) or substringof('${esc}',Role) or substringof('${esc}',Location) or substringof('${esc}',GlobalID))`,
        );
      }

      // const items = await sp.web.lists
      //   .getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER)
      const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
      const items = await getThresholdSafeListItems({
        web: empMasterWeb,
        listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
        select: "ID,Title,FullName,EmailID,Department,Location,GlobalID,DateofJoining,EmployeeStatus,Role,AssignStatus,Grade,Designation",
        filter: parts.length ? parts.join(" and ") : undefined,
      });

      exportRowsToCsv(
        "user_management.csv",
        [
          { key: "name", label: "Name" },
          { key: "email", label: "Email" },
          { key: "globalId", label: "Global ID" },
          { key: "department", label: "Department" },
          { key: "location", label: "Location" },
          { key: "grade", label: "Grade" },
          { key: "designation", label: "Designation" },
          { key: "doj", label: "Date of Joining" },
          { key: "role", label: "Role" },
          { key: "assignStatus", label: "Assign Status" },
          { key: "status", label: "Status" },
        ],
        (items || []).map((i: any) => ({
          name: i.FullName || i.Title || "",
          email: i.EmailID || "",
          globalId: i.GlobalID || "",
          department: i.Department || "",
          location: i.Location || "",
          grade: i.Grade || "",
          designation: i.Designation || "",
          doj: formatExportDate(i.DateofJoining),
          role: i.Role || "",
          assignStatus: mapAssignStatus(i.AssignStatus),
          status: mapStatus(i.EmployeeStatus),
        })),
      );
    } catch (e) {
      console.error("exportUsers:", e);
      showSnackbar("Export failed. Please try again.", "error", 3000);
    }
  };


  const dateCell = (p: GridRenderCellParams) => (
    <span style={{ fontSize: "0.75rem" }}>
      {formatDisplayDate(p.value as string)}
    </span>
  );

  const chipCell = (
    _color: "success" | "warning" | "error" | "default",
    p: GridRenderCellParams,
  ) => {
    const colors = getMeetingStatusColors(String(p.value || ""));
    return (
      <Chip
        label={p.value || "—"}
        size="small"
        sx={statusBadgeSx(colors, "0.62rem", "19px")}
      />
    );
  };


  const njUnassignedColumns: GridColDef[] = [
    { field: "name", headerName: "Name", flex: 1, minWidth: 130 },
    { field: "email", headerName: "Email", flex: 1.2, minWidth: 160 },
    { field: "department", headerName: "Department", flex: 0.9, minWidth: 120 },
    { field: "grade", headerName: "Grade", flex: 0.7, minWidth: 90 },
    { field: "designation", headerName: "Designation", flex: 1, minWidth: 140 },
    { field: "location", headerName: "Location", flex: 0.8, minWidth: 110 },
    { field: "globalId", headerName: "Global ID", flex: 0.7, minWidth: 100 },
    {
      field: "doj",
      headerName: "DOJ",
      flex: 0.8,
      minWidth: 110,
      renderCell: dateCell,
    },
    {
      field: "status",
      headerName: "Status",
      width: 90,
      type: "singleSelect",
      valueOptions: ["Active", "Inactive"],
      renderCell: (p) => {
        const colors = getEmpStatusColors(String(p.value || ""));
        return (
          <Chip
            label={p.value || "—"}
            size="small"
            sx={statusBadgeSx(colors)}
          />
        );
      },
    },
  ];

  const assignedActionColumn: GridColDef = {
    field: "actions",
    headerName: "Actions",
    width: 220,
    sortable: false,
    renderCell: (p: GridRenderCellParams) => {
      const row = p.row as BuddyAllocation;
      const isInactive = row.employeeStatus === "Inactive";
      // const isInactive = row.employeeStatus === "Inactive";
      return (
        <div style={{ display: "flex", gap: "6px", alignItems: "center" }}>
          <button
            style={{
              ...inlineActionBtn,
              color:
                isInactive || row.overallStatus === "Completed"
                  ? "#9ca3af"
                  : "#00D264",
              borderColor:
                isInactive || row.overallStatus === "Completed"
                  ? "#e5e7eb"
                  : "rgba(0, 210, 100, 0.4)",
              background:
                isInactive || row.overallStatus === "Completed"
                  ? "#f9fafb"
                  : "rgba(0, 210, 100, 0.08)",
              cursor:
                isInactive || row.overallStatus === "Completed"
                  ? "not-allowed"
                  : "pointer",
            }}
            onClick={() =>
              !(isInactive || row.overallStatus === "Completed") &&
              openAssignedReassign(row)
            }
            disabled={isInactive || row.overallStatus === "Completed"}
            title={
              row.overallStatus === "Completed"
                ? "Cannot reassign: All meetings completed"
                : isInactive
                  ? "Cannot reassign: Active status required"
                  : undefined
            }
          >
            <SyncAltIcon sx={{ fontSize: "0.8rem" }} />
            Reassign
          </button>
          <button
            style={{
              ...inlineActionBtn,
              color: isInactive ? "#9ca3af" : "#e04848",
              borderColor: isInactive ? "#e5e7eb" : "rgba(224, 72, 72, 0.4)",
              background: isInactive ? "#f9fafb" : "rgba(224, 72, 72, 0.08)",
              cursor: isInactive ? "not-allowed" : "pointer",
            }}
            onClick={() => !isInactive && openAssignedRelease(row)}
            disabled={isInactive}
          >
            <PersonOffOutlinedIcon sx={{ fontSize: "0.8rem" }} />
            Release NJ
          </button>
        </div>
      );
    },
  };

  const njAssignedColumns: GridColDef[] = [
    assignedActionColumn,
    { field: "njName", headerName: "NJ Name", flex: 1, minWidth: 130 },
    { field: "njEmail", headerName: "NJ Email", flex: 1.1, minWidth: 150 },
    {
      field: "njGlobalId",
      headerName: "Global ID (NJ)",
      flex: 0.7,
      minWidth: 110,
    },
    {
      field: "njDepartment",
      headerName: "Department",
      flex: 0.9,
      minWidth: 110,
    },
    { field: "njGrade", headerName: "Grade", flex: 0.7, minWidth: 90 },
    { field: "njDesignation", headerName: "Designation", flex: 1, minWidth: 140 },
    { field: "njLocation", headerName: "Location", flex: 0.7, minWidth: 100 },
    {
      field: "doj",
      headerName: "Date of Joining",
      width: 120,
      renderCell: dateCell,
    },
    { field: "buddyName", headerName: "Buddy Name", flex: 1, minWidth: 130 },
    { field: "buddyEmail", headerName: "Buddy Email", flex: 1, minWidth: 140 },
    {
      field: "buddyGlobalId",
      headerName: "Global ID (Buddy)",
      flex: 0.7,
      minWidth: 120,
    },
    {
      field: "buddyDepartment",
      headerName: "Buddy Dept",
      flex: 0.8,
      minWidth: 110,
    },
    {
      field: "buddyLocation",
      headerName: "Buddy Location",
      flex: 0.7,
      minWidth: 110,
    },
    {
      field: "allocatedOn",
      headerName: "Allocated On",
      width: 120,
      renderCell: dateCell,
    },
    { field: "allocationCount", headerName: "Alloc Count", width: 90 },
    {
      field: "overallStatus",
      headerName: "Overall Status",
      width: 120,
      type: "singleSelect",
      valueOptions: ["Assigned", "Completed", "Pending", "Inactive", "Released"],
      renderCell: (p) => {
        const { bg, fg } = getOverallStatusColors(p.value);
        return (
          <Chip
            label={p.value || "—"}
            size="small"
            sx={statusBadgeSx({ bg, fg })}
          />
        );
      },
    },
    {
      field: "employeeStatus",
      headerName: "Emp Status",
      width: 105,
      type: "singleSelect",
      valueOptions: ["Active", "Inactive"],
      renderCell: (p) => {
        const { bg, fg } = getEmpStatusColors(p.value);
        return (
          <Chip label={p.value} size="small" sx={statusBadgeSx({ bg, fg })} />
        );
      },
    },
    {
      field: "currentMeetingStatus",
      headerName: "Current Status",
      width: 175,
      renderCell: (p) => {
        const v = String(p.value || "");
        const { bg, fg } = getCurrentStatusColors(v);
        return (
          <Chip
            label={v || "—"}
            size="small"
            sx={statusBadgeSx({ bg, fg }, "0.65rem", "20px")}
          />
        );
      },
    },
    {
      field: "firstInteractionDate",
      headerName: "1st Interaction",
      width: 125,
      renderCell: dateCell,
    },
    {
      field: "secondInteractionDate",
      headerName: "2nd Interaction",
      width: 125,
      renderCell: dateCell,
    },
    {
      field: "thirdInteractionDate",
      headerName: "3rd Interaction",
      width: 125,
      renderCell: dateCell,
    },
    {
      field: "fourthInteractionDate",
      headerName: "4th Interaction",
      width: 125,
      renderCell: dateCell,
    },
    {
      field: "fifthInteractionDate",
      headerName: "5th Interaction",
      width: 125,
      renderCell: dateCell,
    },
    {
      field: "sixthInteractionDate",
      headerName: "6th Interaction",
      width: 125,
      renderCell: dateCell,
    },
    {
      field: "firstMeetingStatus",
      headerName: "M1 Status",
      width: 105,
      renderCell: (p) => chipCell(meetingStatusColor(p.value as string), p),
    },
    {
      field: "secondMeetingStatus",
      headerName: "M2 Status",
      width: 105,
      renderCell: (p) => chipCell(meetingStatusColor(p.value as string), p),
    },
    {
      field: "thirdMeetingStatus",
      headerName: "M3 Status",
      width: 105,
      renderCell: (p) => chipCell(meetingStatusColor(p.value as string), p),
    },
    {
      field: "fourthMeetingStatus",
      headerName: "M4 Status",
      width: 105,
      renderCell: (p) => chipCell(meetingStatusColor(p.value as string), p),
    },
    {
      field: "fifthMeetingStatus",
      headerName: "M5 Status",
      width: 105,
      renderCell: (p) => chipCell(meetingStatusColor(p.value as string), p),
    },
    {
      field: "sixthMeetingStatus",
      headerName: "M6 Status",
      width: 105,
      renderCell: (p) => chipCell(meetingStatusColor(p.value as string), p),
    },
    {
      field: "firstActualInteractionDate",
      headerName: "M1 Actual",
      width: 115,
      renderCell: dateCell,
    },
    {
      field: "secondActualInteractionDate",
      headerName: "M2 Actual",
      width: 115,
      renderCell: dateCell,
    },
    {
      field: "thirdActualInteractionDate",
      headerName: "M3 Actual",
      width: 115,
      renderCell: dateCell,
    },
    {
      field: "fourthActualInteractionDate",
      headerName: "M4 Actual",
      width: 115,
      renderCell: dateCell,
    },
    {
      field: "fifthActualInteractionDate",
      headerName: "M5 Actual",
      width: 115,
      renderCell: dateCell,
    },
    {
      field: "sixthActualInteractionDate",
      headerName: "M6 Actual",
      width: 115,
      renderCell: dateCell,
    },
  ];

  const userColumns: GridColDef[] = [
    {
      field: "actions",
      headerName: "Actions",
      width: 85,
      sortable: false,
      renderCell: (p) => {
        const isInactive = p.row.status === "Inactive";
        const isSuperAdmin = p.row.role === "Super Admin";
        const canModify = !(isLdAdmin && isSuperAdmin);
        return (
          <div className={styles.actionButtons}>
            <Tooltip
              title={
                isInactive
                  ? "Cannot edit inactive user"
                  : isLdAdmin && isSuperAdmin
                    ? "L&D Site Admin cannot modify Super Admin role"
                    : "Edit Role"
              }
            >
              <span>
                <IconButton
                  size="small"
                  onClick={() => openEditDrawer(p.row)}
                  className={styles.iconBtn}
                  disabled={isInactive || !canModify}
                  sx={{
                    opacity: isInactive || !canModify ? 0.5 : 1,
                    cursor: isInactive || !canModify ? "not-allowed" : "pointer",
                  }}
                >
                  <EditOutlinedIcon sx={{ fontSize: "1rem" }} />
                </IconButton>
              </span>
            </Tooltip>
            {isInactive ? (
              // <Tooltip
              //   title={
              //     isLdAdmin && isSuperAdmin
              //       ? "L&D Site Admin cannot activate Super Admin"
              //       : "Activate user"
              //   }
              // >
              //   <span>
              //     <IconButton
              //       size="small"
              //       onClick={() => void handleActivateUser(p.row)}
              //       className={styles.iconBtn}
              //       disabled={!canModify}
              //       sx={{
              //         opacity: !canModify ? 0.5 : 1,
              //         cursor: !canModify ? "not-allowed" : "pointer",
              //         color: !canModify ? "#9ca3af !important" : "#15803d !important",
              //         "&:hover": {
              //           background: "rgba(21,128,61,0.1) !important",
              //           color: !canModify ? "#9ca3af !important" : "#166534 !important",
              //         },
              //       }}
              //     >
              //       <HowToRegIcon sx={{ fontSize: "1.05rem" }} />
              //     </IconButton>
              //   </span>
              // </Tooltip>
              <></>
            ) : (
              <Tooltip
                title={
                  isLdAdmin && isSuperAdmin
                    ? "L&D Site Admin cannot deactivate Super Admin"
                    : "Deactivate user"
                }
              >
                <span>
                  <IconButton
                    size="small"
                    onClick={() => handleDeleteUser(p.row)}
                    className={styles.delIconBtn}
                    disabled={!canModify}
                    sx={{
                      opacity: !canModify ? 0.5 : 1,
                      cursor: !canModify ? "not-allowed" : "pointer",
                      color: !canModify
                        ? "#9ca3af !important"
                        : "#78736e !important",
                      "&:hover": {
                        background: "rgba(120,115,110,0.1) !important",
                        color: !canModify ? "#9ca3af !important" : "#004632 !important",
                      },
                    }}
                  >
                    <DeleteOutlineIcon sx={{ fontSize: "1rem" }} />
                  </IconButton>
                </span>
              </Tooltip>
            )}
          </div>
        );
      },
    },
    { field: "name", headerName: "Name", flex: 1, minWidth: 140 },
    { field: "email", headerName: "Email", flex: 1.2, minWidth: 160 },
    { field: "globalId", headerName: "Global ID", flex: 0.8, minWidth: 110 },
    { field: "department", headerName: "Department", flex: 0.9, minWidth: 120 },
    { field: "grade", headerName: "Grade", width: 90 },
    { field: "designation", headerName: "Designation", flex: 1, minWidth: 140 },
    { field: "location", headerName: "Location", width: 120 },
    { field: "doj", headerName: "DOJ", width: 115, renderCell: dateCell },
    {
      field: "role",
      headerName: "Role",
      width: 145,
      renderCell: (p) => {
        const roleValue =
          p.value && p.value !== "-Select One-" ? p.value : "No Role Assigned";
        const colors = getRoleColors(roleValue);
        return (
          <Chip label={roleValue} size="small" sx={statusBadgeSx(colors)} />
        );
      },
    },
    {
      field: "status",
      headerName: "Status",
      width: 85,
      type: "singleSelect",
      valueOptions: ["Active", "Inactive"],
      renderCell: (p) => {
        const colors = getEmpStatusColors(String(p.value || ""));
        return (
          <Chip
            label={p.value || "—"}
            size="small"
            sx={statusBadgeSx(colors)}
          />
        );
      },
    },
  ];

  return (
    <div
      className={styles.OnboardingContainer}
      ref={containerRef}
      style={{ width: "100%", height: "100%", overflow: "hidden" }}
    >
      <BuddyLoader
        visible={initialDataLoading}
        variant="default"
        message="Loading onboarding…"
        subMessage="Preparing your new joiner and buddy data."
      />

      <main className={styles.main}>
        <div className={styles.pageHeading}>
          <div className={styles.pageHeadingLeft}>
            <div className={styles.pageHeadingIcon}>
              <PeopleAltOutlinedIcon sx={{ fontSize: "1.1rem" }} />
            </div>
            <div>
              <h2 className={styles.pageTitle}>
                {activeTab === "newJoiner"
                  ? "New Joiner Onboarding"
                  : "User Management"}
              </h2>
              <p className={styles.pageSubtitle}>
                {isLdAdmin && currentUserLocation ? (
                  <>
                    Scoped to location:{" "}
                    <strong style={{ color: "#8CFF8C" }}>
                      {currentUserLocation}
                    </strong>
                  </>
                ) : activeTab === "newJoiner" ? (
                  "Assign buddies and track onboarding progress"
                ) : (
                  "Manage roles and access for all employees"
                )}
              </p>
            </div>
          </div>
          <div className={styles.pageHeadingStats}>
            {activeTab === "newJoiner" ? (
              <>
                <div className={styles.headingStat}>
                  <span>{unassignedActiveCount}</span>
                  <span>Unassigned</span>
                </div>
                <div className={styles.headingStat}>
                  <span>{assignedCount}</span>
                  <span>Assigned</span>
                </div>
                <div className={styles.headingStat}>
                  <span>{totalNJCount}</span>
                  <span>Total NJs</span>
                </div>
              </>
            ) : (
              <>
                <div className={styles.headingStat}>
                  <span>{activeUserCount}</span>
                  <span>Active</span>
                </div>
                <div className={styles.headingStat}>
                  <span>{inactiveUserCount}</span>
                  <span>Inactive</span>
                </div>
                <div className={styles.headingStat}>
                  <span>{buddyCount}</span>
                  <span>Buddies</span>
                </div>
                <div className={styles.headingStat}>
                  <span>{superAdminCount}</span>
                  <span>Super Admin</span>
                </div>
                <div className={styles.headingStat}>
                  <span>{ldAdminCount}</span>
                  <span>L&amp;D Admin</span>
                </div>
                <div
                  className={`${styles.headingStat} ${styles.headingStatTotal}`}
                >
                  <span>{totalUserCount}</span>
                  <span>Total</span>
                </div>
              </>
            )}
          </div>
        </div>

        <div className={styles.toolbar}>
          <div className={styles.tabs}>
            <button
              className={`${styles.tab} ${activeTab === "newJoiner" ? styles.tabActive : ""}`}
              onClick={() => handleTabChange("newJoiner")}
            >
              New Joiner Onboarding
            </button>
            <button
              className={`${styles.tab} ${activeTab === "userMgmt" ? styles.tabActive : ""}`}
              onClick={() => handleTabChange("userMgmt")}
            >
              User Management
            </button>
          </div>
          <div className={styles.toolbarRight}>
            <div className={styles.inactiveToggleWrap}>
              <span className={styles.inactiveToggleLabel}>
                Show Inactive
                <span
                  className={`${styles.inactiveBadge} ${inactiveCount === 0 ? styles.inactiveBadgeZero : ""}`}
                >
                  {inactiveCount}
                </span>
              </span>
              <Switch
                checked={showInactive}
                onChange={(e) => setShowInactive(e.target.checked)}
                size="small"
                sx={SWITCH_SX}
              />
            </div>
            <TextField
              placeholder={
                activeTab === "newJoiner"
                  ? "Search new joiners…"
                  : "Search users…"
              }
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
              size="small"
              variant="outlined"
              InputProps={{
                startAdornment: (
                  <InputAdornment position="start">
                    <SearchIcon sx={{ fontSize: "1rem", color: "#004632" }} />
                  </InputAdornment>
                ),
              }}
              sx={{
                minWidth: 200,
                "& .MuiOutlinedInput-root": {
                  fontSize: "0.775rem",
                  "& fieldset": { borderColor: "#e5e7eb" },
                  "&:hover fieldset": { borderColor: "#00D264" },
                  "&.Mui-focused fieldset": { borderColor: "#00D264" },
                },
              }}
            />
            {activeTab === "userMgmt" && (
              <button
                className={styles.assignBtn}
                onClick={() => void handleExportUsers()}
                style={{ padding: "6px 10px", flexShrink: 0, cursor: "pointer" }}
              >
                <FileDownloadOutlinedIcon sx={{ fontSize: "0.85rem", pointerEvents: "none" }} />
                <span>Export</span>
              </button>
            )}
          </div>
        </div>

        {activeTab === "newJoiner" && (
          <>
            <div className={styles.subViewBar}>
              <div className={styles.subViewTabs}>
                <button
                  className={`${styles.subTab} ${njView === "unassigned" ? styles.subTabActive : ""}`}
                  onClick={() => handleNjViewChange("unassigned")}
                >
                  <RadioButtonUncheckedIcon sx={{ fontSize: "0.825rem" }} />
                  Unassigned
                  <span
                    className={`${styles.subTabBadge} ${njView === "unassigned" ? styles.subTabBadgeActive : ""}`}
                  >
                    {unassignedActiveCount}
                  </span>
                </button>
                <button
                  className={`${styles.subTab} ${njView === "assigned" ? styles.subTabActive : ""}`}
                  onClick={() => handleNjViewChange("assigned")}
                >
                  <CheckCircleOutlineIcon sx={{ fontSize: "0.825rem" }} />
                  Assigned
                  <span
                    className={`${styles.subTabBadge} ${njView === "assigned" ? styles.subTabBadgeActive : ""}`}
                  >
                    {assignedCount}
                  </span>
                </button>
              </div>
              <div
                style={{ display: "flex", alignItems: "center", gap: "8px" }}
              >
                {njView === "unassigned" &&
                  selectedNJIds.length > 0 &&
                  !hasLocationMismatch &&
                  selectedNJLocation && (
                    <div className={styles.locationPill}>
                      <LocationOnOutlinedIcon sx={{ fontSize: "0.85rem" }} />
                      Buddy search scoped to:{" "}
                      <strong>{selectedNJLocation}</strong>
                    </div>
                  )}

                {/* {njView === "assigned" && (
                  <Tooltip title="Refresh allocations">
                    <IconButton
                      size="small"
                      onClick={() => void loadBuddyAllocations()}
                      disabled={allocLoading}
                      sx={{ color: "#004632", padding: "0.2rem" }}
                    >
                      <RefreshIcon sx={{ fontSize: "1rem" }} />
                    </IconButton>
                  </Tooltip>
                )} */}
                <button
                  className={styles.assignBtn}
                  onClick={() => void handleExportNewJoiner()}
                  style={{ padding: "6px 10px", cursor: "pointer" }}
                >
                  <FileDownloadOutlinedIcon sx={{ fontSize: "0.85rem", pointerEvents: "none" }} />
                  <span>Export</span>
                </button>
              </div>
            </div>

            {njView === "unassigned" &&
              selectedNJIds.length > 0 &&
              hasLocationMismatch && (
                <div className={styles.locationMismatchAlert}>
                  <WarningAmberIcon sx={{ fontSize: "1rem" }} />
                  <span>
                    Selected New Joiners are from{" "}
                    <strong>different locations</strong> (
                    {selectedNJLocations.join(", ")}). Please select NJs from
                    the <strong>same location</strong> to assign a buddy.
                  </span>
                </div>
              )}

            {njView === "unassigned" &&
              selectedNJIds.length > 0 &&
              !hasLocationMismatch && (
                <div className={styles.assignBar}>
                  <span className={styles.assignBarCount}>
                    <PersonAddAlt1Icon sx={{ fontSize: "0.9rem" }} />
                    {selectedNJIds.length} selected
                  </span>
                  <span className={styles.assignBarLabel}>Assign buddy:</span>

                  <div className={styles.buddySearchWrap}>
                    <input
                      type="text"
                      className={styles.buddySelect}
                      placeholder={
                        selectedNJLocation
                          ? `Search buddies in ${selectedNJLocation}…`
                          : "Search by name, email or Global ID…"
                      }
                      value={buddySearchText}
                      onChange={(e) => {
                        const v = e.target.value;
                        setBuddySearchText(v);
                        setSelectedBuddy("");
                        setSelectedBuddyData(null);
                        void handleBuddySearch(
                          v,
                          "unassigned",
                          selectedNJLocation,
                        );
                      }}
                      onFocus={() => {
                        setShowBuddySearch(true);
                        if (buddySearchText.trim().length >= 2)
                          void handleBuddySearch(
                            buddySearchText,
                            "unassigned",
                            selectedNJLocation,
                          );
                      }}
                      onBlur={() => {
                        buddyBlurTimeout.current = window.setTimeout(
                          () => setShowBuddySearch(false),
                          200,
                        );
                      }}
                    />
                    {showBuddySearch && (
                      <div className={styles.buddyDropdown}>
                        {buddySearchText.trim().length < 2 ? (
                          <div className={styles.buddyDropdownHint}>
                            Type at least 2 characters to search…
                          </div>
                        ) : buddySearchResults.length > 0 ? (
                          buddySearchResults.map((b) => (
                            <div
                              key={b.id}
                              className={styles.buddyOption}
                              onMouseDown={(e) => e.preventDefault()}
                              onClick={() => handleBuddySelect(b)}
                            >
                              <div className={styles.buddyOptionName}>
                                {b.name}
                              </div>
                              <div className={styles.buddyOptionDept}>
                                {b.department}
                              </div>
                              <div className={styles.buddyOptionMeta}>
                                {b.email}&nbsp;|&nbsp;{b.globalId}
                                {b.location && (
                                  <>&nbsp;|&nbsp;📍 {b.location}</>
                                )}
                              </div>
                            </div>
                          ))
                        ) : (
                          <div className={styles.buddyDropdownHint}>
                            {selectedNJLocation
                              ? `No buddy found in location "${selectedNJLocation}".`
                              : "No buddy found."}
                          </div>
                        )}
                      </div>
                    )}
                  </div>

                  {selectedBuddyData && (
                    <div className={styles.buddyInfo}>
                      <span className={styles.buddyInfoItem}>
                        <strong>Email:</strong> {selectedBuddyData.email}
                      </span>
                      <span className={styles.buddyInfoItem}>
                        <strong>Global ID:</strong> {selectedBuddyData.globalId}
                      </span>
                      <span className={styles.buddyInfoItem}>
                        <strong>Dept:</strong> {selectedBuddyData.department}
                      </span>
                      {selectedBuddyData.location && (
                        <span className={styles.buddyInfoItem}>
                          <strong>Location:</strong>{" "}
                          {selectedBuddyData.location}
                        </span>
                      )}
                    </div>
                  )}

                  <button
                    className={styles.assignBtn}
                    onClick={handleAssignBuddy}
                    disabled={!selectedBuddy || isAssigning}
                  >
                    <SaveIcon sx={{ fontSize: "0.85rem" }} />
                    <span>{isAssigning ? "Assigning…" : "Assign"}</span>
                  </button>
                  <button
                    className={styles.clearBtn}
                    onClick={handleClearSelection}
                  >
                    Clear
                  </button>
                </div>
              )}

            <div className={styles.gridCard}>
              {njView === "unassigned" ? (
                <DataGrid
                  rows={newJoiners}
                  columns={njUnassignedColumns}
                  checkboxSelection
                  isRowSelectable={(p) =>
                    p.row.assignStatus !== "Assigned" &&
                    p.row.status === "Active"
                  }
                  rowSelectionModel={selectedNJIds}
                  onRowSelectionModelChange={(ids) =>
                    setSelectedNJIds(ids as string[])
                  }
                  disableRowSelectionOnClick
                  paginationMode="server"
                  filterMode="server"
                  filterModel={njFilterModel as any}
                  onFilterModelChange={(model) => setNjFilterModel(model as SimpleGridFilterModel)}
                  rowCount={njRowCount}
                  paginationModel={njPaginationModel}
                  onPaginationModelChange={setNjPaginationModel}
                  loading={njLoading}
                  pageSizeOptions={[5, 10, 20, 50]}
                  rowHeight={38}
                  columnHeaderHeight={38}
                  getRowClassName={(p) =>
                    p.row.status === "Inactive" ? "row-inactive" : ""
                  }
                  sx={GRID_SX}
                />
              ) : (
                <DataGrid
                  rows={buddyAllocations}
                  getRowId={(row: BuddyAllocation) => row.allocId}
                  columns={njAssignedColumns}
                  paginationMode="server"
                  filterMode="server"
                  filterModel={allocFilterModel as any}
                  onFilterModelChange={(model) => setAllocFilterModel(model as SimpleGridFilterModel)}
                  rowCount={allocRowCount}
                  paginationModel={allocPaginationModel}
                  onPaginationModelChange={setAllocPaginationModel}
                  loading={allocLoading}
                  disableRowSelectionOnClick
                  pageSizeOptions={[5, 10, 20, 50]}
                  rowHeight={38}
                  columnHeaderHeight={38}
                  getRowClassName={(p) =>
                    p.row.employeeStatus === "Inactive" ? "row-inactive" : ""
                  }
                  sx={GRID_SX}
                />
              )}
            </div>
          </>
        )}

        {activeTab === "userMgmt" && (
          <div className={styles.gridCard}>
            <DataGrid
              rows={users}
              columns={userColumns}
              disableRowSelectionOnClick
              paginationMode="server"
              filterMode="server"
              filterModel={userFilterModel as any}
              onFilterModelChange={(model) => setUserFilterModel(model as SimpleGridFilterModel)}
              rowCount={userRowCount}
              paginationModel={userPaginationModel}
              onPaginationModelChange={setUserPaginationModel}
              loading={userLoading}
              pageSizeOptions={[5, 10, 20, 50]}
              rowHeight={38}
              columnHeaderHeight={38}
              getRowClassName={(p) =>
                p.row.status === "Inactive" ? "row-inactive" : ""
              }
              sx={GRID_SX}
            />
          </div>
        )}
      </main>

      <Drawer
        anchor="right"
        open={editDrawerOpen}
        onClose={() => setEditDrawerOpen(false)}
        ModalProps={{
          disablePortal: true,
          keepMounted: true,
          container: containerRef.current,
        }}
        sx={DRAWER_SX}
      >
        <div className={styles.drawer}>
          <div className={styles.drawerHeader}>
            <div>
              <h2 className={styles.drawerTitle}>Edit User Role</h2>
              <p className={styles.drawerSubtitle}>{editingUser?.name ?? ""}</p>
            </div>
            <IconButton
              onClick={() => setEditDrawerOpen(false)}
              className={styles.closeBtn}
            >
              <CloseIcon />
            </IconButton>
          </div>
          <div className={styles.drawerBody}>
            {editingUser && (
              <div className={styles.editForm}>
                <div className={styles.formRow3}>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>Name</label>
                    <div className={styles.editValue}>{editingUser.name}</div>
                  </div>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>Email</label>
                    <div className={styles.editValue}>{editingUser.email}</div>
                  </div>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>Department</label>
                    <div className={styles.editValue}>
                      {editingUser.department || "—"}
                    </div>
                  </div>
                </div>
                <div className={styles.formRow3}>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>Location</label>
                    <div className={styles.editValue}>
                      {editingUser.location || "—"}
                    </div>
                  </div>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>Global ID</label>
                    <div className={styles.editValue}>
                      {editingUser.globalId || "—"}
                    </div>
                  </div>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>DOJ</label>
                    <div className={styles.editValue}>
                      {formatDisplayDate(editingUser.doj)}
                    </div>
                  </div>
                </div>
                <div className={styles.formRow2}>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>Current Role</label>
                    <div className={styles.editValue}>
                      {(() => {
                        const roleValue =
                          editingUser.role &&
                          editingUser.role !== "-Select One-"
                            ? editingUser.role
                            : "No Role Assigned";
                        return (
                          <Chip
                            label={roleValue}
                            size="small"
                            sx={statusBadgeSx(
                              getRoleColors(roleValue),
                              "0.7rem",
                            )}
                          />
                        );
                      })()}
                    </div>
                  </div>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>
                      New Role <span className={styles.required}>*</span>
                    </label>
                    <select
                      className={styles.select}
                      value={editRole}
                      onChange={(e) =>
                        setEditRole(e.target.value as SharePointRole)
                      }
                    >
                      <option value="-Select One-">— Select Role —</option>
                      {isSuperAdmin && (
                        <option value="Super Admin">Super Admin</option>
                      )}
                      <option value="Buddy">Buddy</option>
                      <option value="L&D Site Admin">L&amp;D Site Admin</option>
                      <option value="New Joinee">New Joinee</option>
                      <option value="No-Role">No-Role</option>
                    </select>
                  </div>
                  {(isSuperAdmin || isLdAdmin) && (editRole === "Super Admin" || editRole === "L&D Site Admin") && (
                    <div className={styles.editField}>
                      <label className={styles.editLabel}>
                        Is Help Desk Executive
                      </label>
                      <div className={styles.editValue}>
                        <Switch
                          checked={editIsHelpDeskExecutive}
                          onChange={(e) =>
                            setEditIsHelpDeskExecutive(e.target.checked)
                          }
                          sx={SWITCH_SX}
                        />
                        <span style={{ marginLeft: "8px" }}>
                          {editIsHelpDeskExecutive ? "Yes" : "No"}
                        </span>
                      </div>
                    </div>
                  )}
                  {isSuperAdmin && editRole === "Super Admin" && (
                    <div className={styles.editField}>
                      <label className={styles.editLabel}>
                        Is Success Story Approver
                      </label>
                      <div className={styles.editValue}>
                        <Switch
                          checked={editIsSuccessStoryApprover}
                          onChange={(e) =>
                            setEditIsSuccessStoryApprover(e.target.checked)
                          }
                          sx={SWITCH_SX}
                        />
                        <span style={{ marginLeft: "8px" }}>
                          {editIsSuccessStoryApprover ? "Yes" : "No"}
                        </span>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
          <div className={styles.drawerFooter}>
            <button
              className={styles.cancelBtn}
              onClick={() => setEditDrawerOpen(false)}
            >
              Cancel
            </button>
            <button
              className={styles.saveBtn}
              onClick={handleSaveRole}
              disabled={
                isSavingRole ||
                editRole === "-Select One-" ||
                (!!editingUser && 
                 (editRole === "No-Role" ? "-Select One-" : editRole) === editingUser.role && 
                 editIsHelpDeskExecutive === (editingUser.isHelpDeskExecutive || false) &&
                 editIsSuccessStoryApprover ===
                   (editingUser.isSuccessStoryApprover || false))
              }
              title={
                editingUser && 
                (editRole === "No-Role" ? "-Select One-" : editRole) === editingUser.role && 
                editIsHelpDeskExecutive === (editingUser.isHelpDeskExecutive || false) &&
                editIsSuccessStoryApprover ===
                  (editingUser.isSuccessStoryApprover || false)
                  ? "No changes detected"
                  : undefined
              }
            >
              <SaveIcon sx={{ fontSize: "0.95rem" }} />
              <span>{isSavingRole ? "Saving…" : "Save Role"}</span>
            </button>
          </div>
        </div>
      </Drawer>

      <Drawer
        anchor="right"
        open={assignedDrawerKind === "reassign"}
        onClose={resetAssignedDrawer}
        ModalProps={{
          disablePortal: true,
          keepMounted: true,
          container: containerRef.current,
        }}
        sx={DRAWER_SX}
      >
        {activeAssignedRow && (
          <div className={styles.drawer}>
            <div className={styles.drawerHeader}>
              <div>
                <h2 className={styles.drawerTitle}>Reassign Buddy</h2>
                <p className={styles.drawerSubtitle}>
                  {activeAssignedRow.njName}
                </p>
              </div>
              <IconButton
                onClick={resetAssignedDrawer}
                className={styles.closeBtn}
              >
                <CloseIcon />
              </IconButton>
            </div>
            <div className={styles.drawerBody}>
              <div className={styles.editForm}>
                <div className={styles.editField}>
                  <label className={styles.editLabel}>Current Buddy</label>
                  <div className={styles.editValue}>
                    {activeAssignedRow.buddyName}
                  </div>
                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: "1fr 1fr 1fr",
                      gap: "8px",
                      marginTop: "8px",
                    }}
                  >
                    <div className={styles.editValue}>
                      <strong>Global ID:</strong>{" "}
                      {activeAssignedRow.buddyGlobalId || "—"}
                    </div>
                    <div className={styles.editValue}>
                      <strong>Department:</strong>{" "}
                      {activeAssignedRow.buddyDepartment || "—"}
                    </div>
                    <div className={styles.editValue}>
                      <strong>Location:</strong>{" "}
                      {activeAssignedRow.buddyLocation || "—"}
                    </div>
                  </div>
                </div>
                <div className={styles.editField}>
                  <label className={styles.editLabel}>
                    New Buddy <span className={styles.required}>*</span>
                    {activeAssignedRow.njLocation && (
                      <span
                        style={{
                          color: "#004632",
                          fontWeight: 600,
                          marginLeft: "6px",
                          fontSize: "0.65rem",
                        }}
                      >
                        📍 {activeAssignedRow.njLocation}
                      </span>
                    )}
                  </label>
                  <div
                    className={styles.buddySearchWrap}
                    style={{ width: "100%" }}
                  >
                    <input
                      type="text"
                      className={styles.buddySelect}
                      placeholder={
                        activeAssignedRow.njLocation
                          ? `Search buddies in ${activeAssignedRow.njLocation}…`
                          : "Search by name, email or Global ID…"
                      }
                      value={assignedBuddySearchText}
                      onChange={(e) => {
                        const v = e.target.value;
                        setAssignedBuddySearchText(v);
                        setSelectedAssignedBuddy(null);
                        void handleBuddySearch(
                          v,
                          "assigned",
                          activeAssignedRow.njLocation,
                        );
                      }}
                      onFocus={() => {
                        setShowAssignedBuddySearch(true);
                        if (assignedBuddySearchText.trim().length >= 2)
                          void handleBuddySearch(
                            assignedBuddySearchText,
                            "assigned",
                            activeAssignedRow.njLocation,
                          );
                      }}
                      onBlur={() => {
                        assignedBuddyBlurTimeout.current = window.setTimeout(
                          () => setShowAssignedBuddySearch(false),
                          200,
                        );
                      }}
                    />
                    {showAssignedBuddySearch && (
                      <div className={styles.buddyDropdown}>
                        {assignedBuddySearchText.trim().length < 2 ? (
                          <div className={styles.buddyDropdownHint}>
                            Type at least 2 characters to search…
                          </div>
                        ) : assignedBuddySearchResults.length > 0 ? (
                          assignedBuddySearchResults.map((b) => (
                            <div
                              key={b.id}
                              className={styles.buddyOption}
                              onMouseDown={(e) => e.preventDefault()}
                              onClick={() => handleAssignedBuddySelect(b)}
                            >
                              <div className={styles.buddyOptionName}>
                                {b.name}
                              </div>
                              <div className={styles.buddyOptionDept}>
                                {b.department}
                              </div>
                              <div className={styles.buddyOptionMeta}>
                                {b.email}&nbsp;|&nbsp;{b.globalId}
                                {b.location && (
                                  <>&nbsp;|&nbsp;📍 {b.location}</>
                                )}
                              </div>
                            </div>
                          ))
                        ) : (
                          <div className={styles.buddyDropdownHint}>
                            {activeAssignedRow.njLocation
                              ? `No buddy found in location "${activeAssignedRow.njLocation}".`
                              : "No buddy found."}
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                </div>
                {selectedAssignedBuddy && (
                  <div className={styles.buddyInfo}>
                    <span className={styles.buddyInfoItem}>
                      <strong>Email:</strong> {selectedAssignedBuddy.email}
                    </span>
                    <div
                      style={{
                        display: "grid",
                        gridTemplateColumns: "1fr 1fr 1fr",
                        gap: "8px",
                        width: "100%",
                        marginTop: "8px",
                      }}
                    >
                      <span className={styles.buddyInfoItem}>
                        <strong>Global ID:</strong>{" "}
                        {selectedAssignedBuddy.globalId || "—"}
                      </span>
                      <span className={styles.buddyInfoItem}>
                        <strong>Department:</strong>{" "}
                        {selectedAssignedBuddy.department || "—"}
                      </span>
                      <span className={styles.buddyInfoItem}>
                        <strong>Location:</strong>{" "}
                        {selectedAssignedBuddy.location || "—"}
                      </span>
                    </div>
                  </div>
                )}
              </div>
            </div>
            <div className={styles.drawerFooter}>
              <button
                className={styles.cancelBtn}
                onClick={resetAssignedDrawer}
              >
                Cancel
              </button>
              <button
                className={styles.saveBtn}
                onClick={handleAssignedReassignSave}
                disabled={
                  !selectedAssignedBuddy ||
                  isSavingAssignedAction ||
                  (activeAssignedRow &&
                    activeAssignedRow.buddyGlobalId ===
                      selectedAssignedBuddy?.globalId) ||
                  (activeAssignedRow &&
                    activeAssignedRow.overallStatus === "Completed")
                }
                title={
                  selectedAssignedBuddy &&
                  activeAssignedRow?.buddyGlobalId ===
                    selectedAssignedBuddy.globalId
                    ? "Cannot reassign to the same buddy"
                    : activeAssignedRow?.overallStatus === "Completed"
                      ? "Cannot reassign: All meetings completed"
                      : undefined
                }
              >
                <SaveIcon sx={{ fontSize: "0.95rem" }} />
                <span>
                  {isSavingAssignedAction ? "Saving…" : "Save Changes"}
                </span>
              </button>
            </div>
          </div>
        )}
      </Drawer>

      <Drawer
        anchor="right"
        open={assignedDrawerKind === "release"}
        onClose={resetAssignedDrawer}
        ModalProps={{
          disablePortal: true,
          keepMounted: true,
          container: containerRef.current,
        }}
        sx={DRAWER_SX}
      >
        {activeAssignedRow && (
          <div className={styles.drawer}>
            <div className={styles.drawerHeader}>
              <div>
                <h2 className={styles.drawerTitle}>Release New Joinee</h2>
                <p className={styles.drawerSubtitle}>
                  {activeAssignedRow.njName}
                </p>
              </div>
              <IconButton
                onClick={resetAssignedDrawer}
                className={styles.closeBtn}
              >
                <CloseIcon />
              </IconButton>
            </div>
            <div className={styles.drawerBody}>
              <div className={styles.editForm}>
                <div className={styles.formRow3}>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>Global ID</label>
                    <div className={styles.editValue}>
                      {activeAssignedRow.njGlobalId || "—"}
                    </div>
                  </div>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>Department</label>
                    <div className={styles.editValue}>
                      {activeAssignedRow.njDepartment || "—"}
                    </div>
                  </div>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>Location</label>
                    <div className={styles.editValue}>
                      {activeAssignedRow.njLocation || "—"}
                    </div>
                  </div>
                </div>
                <div className={styles.formRow2}>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>
                      Current Employee Status
                    </label>
                    <div className={styles.editValue}>
                      {activeAssignedRow.employeeStatus}
                    </div>
                  </div>
                  <div className={styles.editField}>
                    <label className={styles.editLabel}>Change Status</label>
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: "8px",
                      }}
                    >
                      <span style={{ fontSize: "12px", fontWeight: 600 }}>
                        Inactive
                      </span>
                      <Switch
                        checked={releaseActive}
                        onChange={() => {
                          setReleaseActive((p) => !p);
                          setReleaseConfirm(true);
                        }}
                        sx={SWITCH_SX}
                      />
                      <span style={{ fontSize: "12px", fontWeight: 600 }}>
                        Active
                      </span>
                    </div>
                  </div>
                </div>
                {releaseConfirm && (
                  <div
                    style={{
                      border: "1px solid rgba(245,158,11,0.4)",
                      background: "#fffbeb",
                      color: "#92400e",
                      borderRadius: "8px",
                      padding: "10px 12px",
                      fontSize: "12px",
                      fontWeight: 600,
                    }}
                  >
                    {releaseActive
                      ? `You are about to re-activate "${activeAssignedRow.njName}".`
                      : `You are about to release "${activeAssignedRow.njName}" and mark them inactive. The BuddyAllocate record will also be set to Inactive.`}
                  </div>
                )}
              </div>
            </div>
            <div className={styles.drawerFooter}>
              <button
                className={styles.cancelBtn}
                onClick={resetAssignedDrawer}
              >
                Cancel
              </button>
              <button
                className={styles.saveBtn}
                onClick={handleAssignedReleaseSave}
                disabled={!releaseConfirm || isSavingAssignedAction}
              >
                <SaveIcon sx={{ fontSize: "0.95rem" }} />
                <span>{isSavingAssignedAction ? "Saving…" : "Confirm"}</span>
              </button>
            </div>
          </div>
        )}
      </Drawer>

      <Dialog
        open={deleteConfirmDialog.open}
        onClose={handleCancelDelete}
        aria-labelledby="delete-dialog-title"
        aria-describedby="delete-dialog-description"
        PaperProps={{
          sx: {
            borderRadius: "12px",
            boxShadow: "0 10px 40px rgba(0, 0, 0, 0.15)",
          },
        }}
      >
        <DialogTitle
          id="delete-dialog-title"
          sx={{
            backgroundColor: "#004632",
            color: "#fff",
            fontWeight: 700,
            fontSize: "1.1rem",
            padding: "1rem 1.5rem",
          }}
        >
          Confirm Deactivation
        </DialogTitle>
        <DialogContent
          sx={{ padding: "1.5rem", fontSize: "0.95rem", color: "#374151" }}
        >
          <DialogContentText
            id="delete-dialog-description"
            sx={{ color: "#374151", fontSize: "0.95rem", marginTop: "0.5rem" }}
          >
            Are you sure you want to deactivate{" "}
            <strong>{deleteConfirmDialog.user?.name}</strong>? This will change
            their status from Active to InActive.
          </DialogContentText>
        </DialogContent>
        <DialogActions sx={{ padding: "1rem 1.5rem", gap: "0.75rem" }}>
          <Button
            onClick={handleCancelDelete}
            sx={{
              color: "#004632",
              fontWeight: 600,
              "&:hover": { backgroundColor: "rgba(0, 70, 50, 0.06)" },
            }}
          >
            Cancel
          </Button>
          <Button
            onClick={handleConfirmDelete}
            variant="contained"
            disabled={isDeletingUser}
            sx={{
              backgroundColor: "#dc2626",
              color: "#fff",
              fontWeight: 600,
              "&:hover": { backgroundColor: "#b91c1c" },
              "&:disabled": { backgroundColor: "#fca5a5", color: "#fee2e2" },
            }}
          >
            {isDeletingUser ? "Deactivating..." : "Deactivate"}
          </Button>
        </DialogActions>
      </Dialog>
    </div>
  );
};

export default Onboarding;