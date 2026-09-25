/* eslint-disable @typescript-eslint/explicit-function-return-type*/
/* eslint-disable @typescript-eslint/no-explicit-any*/
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable react/self-closing-comp */
/* eslint-disable react/no-unescaped-entities */
/* eslint-disable @typescript-eslint/no-var-requires*/
/* eslint-disable @typescript-eslint/no-non-null-assertion */
/* eslint-disable max-lines*/

import * as React from "react";
import { useState, useCallback, useEffect, useRef, useMemo } from "react";
import { sp, Web } from "@pnp/sp/presets/all";

import type { IBuddyAppProps } from "./IBuddyAppProps";
import {
  Home,
  Users,
  Calendar,
  FileText,
  BarChart3,
  X,
  Gift,
  TrendingUp,
  ChevronLeft,
  ChevronRight,
  ZoomIn,
  Plus,
  Upload,
  Star,
  ImagePlus,
  Image,
  MessageSquarePlus,
  EyeOff,
  Eye,
  RefreshCw,
  Headphones,
  Check,
  Clock3,
  Bell,
  BookOpen,
} from "lucide-react";

import { Onboarding } from "./Onboarding";
import { MeetingDetails } from "./Meetingdetails";
import { BuddyQuestionnaire } from "./BuddyQuestionnaire";
import { JoineeQuestionnaire } from "./JoineeQuestionnaire";
import { Holidays } from "./Holidays";
import HelpDesk from "./HelpDesk";
import BannerLogoUpload from "./BannerLogoUpload";
import UserGuides from "./UserGuides";
import BuddyPerformance, {
  ICurrentBuddyRanking,
} from "./BuddyPerformance";
import TopPerformerPopup from "./TopPerformerPopup";
import { SnackbarProvider, useSnackbar } from "./Snackbar";
import { BuddyLoader } from "./Buddyloader";
import styles from "./BuddyApp.module.scss";
import LIST_CONFIG, { getThresholdSafeListItems } from "../../../config/spListConfig";

import PersonAddAlt1Icon from "@mui/icons-material/PersonAddAlt1";
import PersonOffOutlinedIcon from "@mui/icons-material/PersonOffOutlined";
import HowToRegOutlinedIcon from "@mui/icons-material/HowToRegOutlined";
import PersonAddAltOutlinedIcon from "@mui/icons-material/PersonAddAltOutlined";
import TaskAltIcon from "@mui/icons-material/TaskAlt";
import HourglassTopIcon from "@mui/icons-material/HourglassTop";

const NEULAND_LOGO_URL = require("../assets/logo.png");
const REVIEW_PAGE_SIZE = 6;
const MY_REQUEST_PAGE_SIZE = 6;

type ReactionType = "Clap" | "Celebrate" | "Love" | "Like" | "HeartEyes";

const REACTION_OPTIONS: ReadonlyArray<{
  type: ReactionType;
  emoji: string;
  label: string;
}> = [
  { type: "Clap", emoji: "👏", label: "Clap" },
  { type: "Celebrate", emoji: "🎉", label: "Celebrate" },
  { type: "Love", emoji: "❤️", label: "Love" },
  { type: "Like", emoji: "👍", label: "Like" },
  { type: "HeartEyes", emoji: "😍", label: "Heart Eyes" },
];

interface StoryReactionSummary {
  counts: Record<ReactionType, number>;
  currentUserReaction?: ReactionType;
}

interface StoryReactionRecord {
  userEmail: string;
  reactionType: ReactionType;
}

interface ReactionFeedback {
  storyId: number;
  emoji: string;
  mode: "selected" | "removed";
  token: number;
}

const createEmptyReactionCounts = (): Record<ReactionType, number> => ({
  Clap: 0,
  Celebrate: 0,
  Love: 0,
  Like: 0,
  HeartEyes: 0,
});

const toReactionType = (value?: string): ReactionType | undefined => {
  const normalizedValue = String(value || "")
    .replace(/\s+/g, "")
    .toLowerCase();
  const match = REACTION_OPTIONS.find(
    (option) => option.type.toLowerCase() === normalizedValue,
  );
  return match?.type;
};

const parseStoryReactions = (value?: string): StoryReactionRecord[] => {
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    if (!Array.isArray(parsed)) return [];

    const reactionsByUser = new Map<string, StoryReactionRecord>();
    parsed.forEach((item: any) => {
      const userEmail = String(item?.userEmail || item?.UserEmail || "")
        .trim()
        .toLowerCase();
      const reactionType = toReactionType(
        item?.reactionType || item?.ReactionType,
      );
      if (!userEmail || !reactionType) return;

      reactionsByUser.set(userEmail, {
        userEmail,
        reactionType,
      });
    });
    return Array.from(reactionsByUser.values());
  } catch {
    return [];
  }
};

const summarizeStoryReactions = (
  reactions: StoryReactionRecord[],
  currentUserEmail: string,
): StoryReactionSummary => {
  const summary: StoryReactionSummary = {
    counts: createEmptyReactionCounts(),
  };
  reactions.forEach((reaction) => {
    summary.counts[reaction.reactionType] += 1;
    if (reaction.userEmail === currentUserEmail) {
      summary.currentUserReaction = reaction.reactionType;
    }
  });
  return summary;
};


const injectHideCss = (): void => {
  if (typeof document === "undefined") return;

  const styleId = "cms-hide-css";
  let style = document.getElementById(styleId) as HTMLStyleElement | null;

  if (!style) {
    style = document.createElement("style");
    style.id = styleId;
    document.head.appendChild(style);
  }

  style.innerHTML = `
    .headerRow-50 { display: none !important; }
    #spLeftNav,
    #spCommandBar { display: none !important; }
    #CommentsWrapper { display: none !important; }
    .CanvasZoneSectionContainer { margin: 0 auto !important; max-width: 100% !important; width: 100% !important; }
    #RecommendedItems { display: none !important; }
    .headerRow-107 { display: none !important; }
    #O365_MainLink_Settings_container,
    #O365_MainLink_Help_container,
    #TipsNTricksButton_container,
    #GiveFeedbackButton_container { display: none !important; }
    #spSiteHeader { display: none !important; }
    #CommentsWrapper,
    footer { display: none !important; }
    .ControlZone { margin: 0px !important; padding: 0 !important; }
    [id="7a4a60bd-8245-46dd-8a9a-14d24d5ce96c"] { margin: 15px !important; padding: 0 !important; }
    #workbenchCommandBar { display: none !important; }
    [dir="ltr"] .CanvasComponent.LCS .Canvas--edit.Canvas--withLayout,
    .CanvasZone,
    .c_XTVxy_hHQBj,
    .CanvasComponent.LCS .ControlZone.ControlZone--edit {
      margin: 0 !important;
      padding: 0 !important;
    }
    #sp-appBar,
    .CanvasToolboxHint-plusButtonWrapper { display: none !important; }
    .p_U54gn_1x34n,
    .w__2lm3_hHQBj.is-focusVisible,
    .w__2lm3_hHQBj { overflow: hidden !important; }
    #workbenchPageContent{
    max-width: 100% !important;
    width: 100% !important;}
  `;
};

injectHideCss();

interface GalleryItem {
  img: string;
  label: string;
  status?: string;
  itemId?: number;
  fileType?: "image" | "video";
}
interface Story {
  name: string;
  team: string;
  quote: string;
  rating?: number;
  id?: number;
  status?: string;
  buddyName?: string;
  EmailID?: string;
  buddyGlobalID?: string;
  employeeLocation?: string;
  authorRole?: string;
  employeeEmail?: string;
  approvalStatus?: "Pending" | "Approved" | "Rejected";
  reviewedByName?: string;
  reviewedByEmail?: string;
  reviewedOn?: string;
  reviewComment?: string;
  createdOn?: string;
  reactionsJson?: string;
}
interface LightboxState {
  open: boolean;
  img: string;
  label: string;
  fileType?: "image" | "video";
}

type GalleryAction = "activate" | "deactivate";

const cleanTitle = (title: string): string => title.trim().replace(/\s+/g, " ");
const cleanWhitespace = (text: string): string => {
  if (!text) return "";
  return text.trim().replace(/\s+/g, " ");
};

const escSP = (value?: string): string =>
  String(value || "")
    .replace(/'/g, "''")
    .trim();
const normalizeRole = (value?: string): string => String(value || "").trim();
const normalizeStatus = (value?: string): string =>
  String(value || "")
    .trim()
    .toLowerCase();
const isActiveEmployee = (row: any): boolean =>
  normalizeStatus(row?.EmployeeStatus) === "active";
const isAssignedStatus = (row: any): boolean =>
  normalizeStatus(row?.AssignStatus) === "assigned";
const hasRole = (row: any, role: string): boolean =>
  normalizeRole(row?.Role) === role;

const HalfStarRating: React.FC<{
  value: number;
  onChange?: (v: number) => void;
  readonly?: boolean;
  size?: number;
}> = ({ value, onChange, readonly = false, size = 22 }) => {
  const [hovered, setHovered] = useState<number | null>(null);
  const display = hovered !== null ? hovered : value;
  return (
    <div
      style={{
        display: "inline-flex",
        gap: "2px",
        alignItems: "center",
        lineHeight: "1",
      }}
    >
      {[1, 2, 3, 4, 5].map((star) => {
        const full = display >= star;
        const half = !full && display >= star - 0.5;
        return (
          <div
            key={star}
            style={{
              position: "relative",
              width: size,
              height: size,
              cursor: readonly ? "default" : "pointer",
              flexShrink: 0,
            }}
            onMouseLeave={!readonly ? () => setHovered(null) : undefined}
          >
            <span
              style={{
                position: "absolute",
                inset: 0,
                fontSize: size,
                lineHeight: "1",
                color: "#e0e0e0",
                userSelect: "none",
              }}
            >
              ★
            </span>
            {(full || half) && (
              <span
                style={{
                  position: "absolute",
                  inset: 0,
                  fontSize: size,
                  lineHeight: "1",
                  color: "#FFD700",
                  overflow: "hidden",
                  width: full ? "100%" : "50%",
                  userSelect: "none",
                }}
              >
                ★
              </span>
            )}
            {!readonly && (
              <div
                style={{
                  position: "absolute",
                  left: 0,
                  top: 0,
                  width: "50%",
                  height: "100%",
                  zIndex: 2,
                }}
                onMouseEnter={() => setHovered(star - 0.5)}
                onClick={() => onChange?.(star - 0.5)}
              />
            )}
            {!readonly && (
              <div
                style={{
                  position: "absolute",
                  right: 0,
                  top: 0,
                  width: "50%",
                  height: "100%",
                  zIndex: 2,
                }}
                onMouseEnter={() => setHovered(star)}
                onClick={() => onChange?.(star)}
              />
            )}
          </div>
        );
      })}
    </div>
  );
};

const AvatarImg: React.FC<{
  name: string;
  email?: string;
  webAbsoluteUrl: string;
  context?: any;
  photoUrl?: string;
  size?: number;
}> = ({ name, context, photoUrl, size = 46 }) => {
  const [imgError, setImgError] = useState(false);
  const [graphPhotoUrl, setGraphPhotoUrl] = useState("");
  const resolvedPhotoUrl = graphPhotoUrl || String(photoUrl || "").trim();

  useEffect(() => {
    let isActive = true;
    let objectUrl = "";
    setGraphPhotoUrl("");
    if (!context?.msGraphClientFactory) {
      return () => { isActive = false; };
    }

    context.msGraphClientFactory
      .getClient("3")
      .then((client: any) => client
        .api("/me/photo/$value")
        .responseType("blob" as any)
        .get())
      .then((photo: any) => {
        if (!isActive || !photo) return;
        const blob = photo instanceof Blob
          ? photo
          : new Blob([photo], { type: "image/jpeg" });
        if (!blob.size) return;
        objectUrl = URL.createObjectURL(blob);
        setGraphPhotoUrl(objectUrl);
      })
      .catch(() => {
        // The signed SharePoint Picture URL remains the fallback.
      });

    return () => {
      isActive = false;
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [context]);

  useEffect(() => setImgError(false), [resolvedPhotoUrl]);
  const nameParts = (name || "?").trim().split(/\s+/).filter(Boolean);
  const initials =
    nameParts.length >= 2
      ? (nameParts[0][0] + nameParts[nameParts.length - 1][0]).toUpperCase()
      : (name || "?")[0].toUpperCase();
  if (resolvedPhotoUrl && !imgError) {
    return (
      <img
        src={resolvedPhotoUrl}
        alt={name}
        onError={() => setImgError(true)}
        style={{
          width: size,
          height: size,
          borderRadius: "50%",
          objectFit: "cover",
          flexShrink: 0,
          border: "2.5px solid #00D264",
          boxShadow: "0 3px 8px rgba(0,70,50,0.25)",
        }}
      />
    );
  }
  return (
    <div
      style={{
        width: size,
        height: size,
        borderRadius: "50%",
        background: "linear-gradient(135deg, #004632, #00D264)",
        color: "#fff",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        fontWeight: 700,
        fontSize: Math.round(size * 0.35),
        flexShrink: 0,
        boxShadow: "0 3px 8px rgba(0,70,50,0.25)",
      }}
    >
      {initials}
    </div>
  );
};

const EmptyState: React.FC<{
  icon: React.ReactElement;
  title: string;
  subtitle: string;
  action?: { label: string; onClick: () => void };
}> = ({ icon, title, subtitle, action }) => (
  <div
    style={{
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      padding: "2.5rem 1.5rem",
      textAlign: "center",
      gap: "0.75rem",
      flex: 1,
      minHeight: "160px",
    }}
  >
    <div
      style={{
        width: 52,
        height: 52,
        borderRadius: "50%",
        background:
          "linear-gradient(135deg, rgba(0,70,50,0.08), rgba(0,210,100,0.12))",
        border: "1.5px dashed rgba(0,210,100,0.4)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        color: "#00a86b",
        marginBottom: "0.25rem",
      }}
    >
      {icon}
    </div>
    <div>
      <div
        style={{
          fontSize: "0.85rem",
          fontWeight: 700,
          color: "#004632",
          marginBottom: "0.25rem",
        }}
      >
        {title}
      </div>
      <div
        style={{
          fontSize: "0.75rem",
          color: "#6b7280",
          lineHeight: 1.55,
          maxWidth: 260,
        }}
      >
        {subtitle}
      </div>
    </div>
    {action && (
      <button
        onClick={action.onClick}
        style={{
          marginTop: "0.25rem",
          padding: "0.4rem 1rem",
          background: "#004632",
          color: "#fff",
          border: "none",
          borderRadius: "8px",
          fontSize: "0.75rem",
          fontWeight: 600,
          cursor: "pointer",
          display: "inline-flex",
          alignItems: "center",
          gap: "0.3rem",
          fontFamily: "Inter, sans-serif",
        }}
      >
        {action.label}
      </button>
    )}
  </div>
);

interface KpiCard {
  icon: React.ReactElement;
  label: string;
  value: string;
  trend: string;
  dir: "up" | "down";
  bg: string;
  full?: boolean;
}

const DEFAULT_KPI_CARDS: KpiCard[] = [
  {
    icon: <HowToRegOutlinedIcon sx={{ fontSize: "1.1rem" }} />,
    label: "Overview 1",
    value: "—",
    trend: "Loading…",
    dir: "up",
    bg: "linear-gradient(135deg,#005c44,#00a86b)",
  },
  {
    icon: <PersonAddAltOutlinedIcon sx={{ fontSize: "1.1rem" }} />,
    label: "Overview 2",
    value: "—",
    trend: "Loading…",
    dir: "up",
    bg: "linear-gradient(135deg,#00D264,#8CFF8C)",
  },
  {
    icon: <PersonAddAlt1Icon sx={{ fontSize: "1.1rem" }} />,
    label: "Overview 3",
    value: "—",
    trend: "Loading…",
    dir: "up",
    bg: "linear-gradient(135deg,#004632,#00a86b)",
  },
  {
    icon: <PersonOffOutlinedIcon sx={{ fontSize: "1.1rem" }} />,
    label: "Overview 4",
    value: "—",
    trend: "Loading…",
    dir: "up",
    bg: "linear-gradient(135deg,#78736e,#a09b97)",
  },
  {
    icon: <TaskAltIcon sx={{ fontSize: "1.1rem" }} />,
    label: "Overview 5",
    value: "—",
    trend: "Loading…",
    dir: "up",
    bg: "linear-gradient(135deg,#004632,#00D264)",
    full: true,
  },
];


const BuddyAppContent = (props: IBuddyAppProps): React.ReactElement => {
  const { showSnackbar } = useSnackbar();

  const [activeSection, setActiveSection] = useState<string>("home");
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(true);
  const [employeeData, setEmployeeData] = useState<any>(null);
  const [employeeDataLoaded, setEmployeeDataLoaded] = useState<boolean>(false);
  const [kpiCards, setKpiCards] = useState<KpiCard[]>(DEFAULT_KPI_CARDS);
  const [refreshTrigger, setRefreshTrigger] = useState<number>(0);

  const [globalLoader, setGlobalLoader] = useState<{
    visible: boolean;
    variant: "upload" | "submit" | "update" | "default";
    message?: string;
  }>({ visible: false, variant: "default" });
  const [initialDataLoading, setInitialDataLoading] = useState<boolean>(true);

  const [activeGalleryItems, setActiveGalleryItems] = useState<GalleryItem[]>(
    [],
  );
  const [inactiveGalleryItems, setInactiveGalleryItems] = useState<
    GalleryItem[]
  >([]);
  const [galleryLoading, setGalleryLoading] = useState<boolean>(true);
  const [slideIdx, setSlideIdx] = useState(0);
  const activeGalleryRef = useRef<GalleryItem[]>([]);
  useEffect(() => {
    activeGalleryRef.current = activeGalleryItems;
  }, [activeGalleryItems]);

  const slideTimer = useRef<ReturnType<typeof setInterval> | null>(null);
  const [sliderPaused, setSliderPaused] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const isMountedRef = useRef(true);
  const [stories, setStories] = useState<Story[]>([]);
  const [inactiveStories, setInactiveStories] = useState<Story[]>([]);
  const [storyReactionSummaries, setStoryReactionSummaries] = useState<
    Record<number, StoryReactionSummary>
  >({});
  const [reactionSavingStoryIds, setReactionSavingStoryIds] = useState<
    number[]
  >([]);
  const reactionSavingStoryIdsRef = useRef<Set<number>>(new Set());
  const [reactionFeedback, setReactionFeedback] =
    useState<ReactionFeedback | null>(null);
  const reactionFeedbackTimerRef = useRef<ReturnType<
    typeof setTimeout
  > | null>(null);
  const [pendingStories, setPendingStories] = useState<Story[]>([]);
  const [approvedReviewStories, setApprovedReviewStories] = useState<Story[]>(
    [],
  );
  const [rejectedReviewStories, setRejectedReviewStories] = useState<Story[]>(
    [],
  );
  const [mySubmissions, setMySubmissions] = useState<Story[]>([]);
  const [showReviewCenter, setShowReviewCenter] = useState(false);
  const [showMyRequests, setShowMyRequests] = useState(false);
  const [reviewTab, setReviewTab] = useState<
    "Pending" | "Approved" | "Rejected"
  >("Pending");
  const [reviewSearch, setReviewSearch] = useState("");
  const [reviewPage, setReviewPage] = useState(0);
  const [myRequestPage, setMyRequestPage] = useState(0);
  const [showStoryForm, setShowStoryForm] = useState(false);
  const loadStoriesRef = useRef<() => Promise<void>>(async () => undefined);
  const [storiesLoading, setStoriesLoading] = useState<boolean>(true);
  const [heroImageUrl, setHeroImageUrl] = useState<string>("");
  const [buddyLogoUrl, setBuddyLogoUrl] = useState<string>("");

  const [showWelcome, setShowWelcome] = useState(false);
  const [welcomeDismissed, setWelcomeDismissed] = useState(false);
  const [currentBuddyRanking, setCurrentBuddyRanking] =
    useState<ICurrentBuddyRanking | null>(null);
  const [rankingLookupComplete, setRankingLookupComplete] = useState(false);
  const [showRankCelebration, setShowRankCelebration] = useState(false);
  const [pendingWelcomeSection, setPendingWelcomeSection] = useState<
    string | null
  >(null);

  const handleCurrentUserRankingChange = useCallback(
    (ranking: ICurrentBuddyRanking | null): void => {
      const isEligibleBuddy =
        normalizeRole(employeeData?.Role) === "Buddy" &&
        normalizeStatus(employeeData?.EmployeeStatus) === "active";

      setCurrentBuddyRanking(isEligibleBuddy ? ranking : null);
      setRankingLookupComplete(true);
      if (!isEligibleBuddy) setShowRankCelebration(false);
    },
    [employeeData?.EmployeeStatus, employeeData?.Role],
  );

  const dismissWelcome = useCallback((targetSection?: string): void => {
    setShowWelcome(false);
    setWelcomeDismissed(true);
    if (targetSection && targetSection !== "home") {
      setPendingWelcomeSection(targetSection);
    } else if (targetSection === "home") {
      setActiveSection("home");
    }
  }, []);

  const closeRankCelebration = useCallback((): void => {
    setShowRankCelebration(false);
    if (pendingWelcomeSection) {
      setActiveSection(pendingWelcomeSection);
      setPendingWelcomeSection(null);
    }
  }, [pendingWelcomeSection]);

  useEffect(() => {
    if (!employeeDataLoaded || !employeeData || !rankingLookupComplete) return;

    if (currentBuddyRanking) {
      setShowWelcome(false);
      setWelcomeDismissed(true);
      return;
    }

    const key = `welcome_${employeeData.EmailID}_${new Date().toDateString()}`;
    if (!sessionStorage.getItem(key)) {
      sessionStorage.setItem(key, "1");
      setShowWelcome(true);
    } else {
      // The welcome was already viewed in this tab. Allow the ranking reveal
      // to run as soon as the current-month ranking lookup finishes.
      setWelcomeDismissed(true);
    }
  }, [
    currentBuddyRanking,
    employeeDataLoaded,
    employeeData,
    rankingLookupComplete,
  ]);

  useEffect(() => {
    if (
      !welcomeDismissed ||
      !rankingLookupComplete ||
      !employeeData?.EmailID
    ) {
      return;
    }

    const isEligibleBuddy =
      normalizeRole(employeeData?.Role) === "Buddy" &&
      normalizeStatus(employeeData?.EmployeeStatus) === "active";

    if (!isEligibleBuddy || !currentBuddyRanking) {
      setShowRankCelebration(false);
      if (pendingWelcomeSection) {
        setActiveSection(pendingWelcomeSection);
        setPendingWelcomeSection(null);
      }
      return;
    }

    const revealKey = `buddy_rank_reveal_v4_${employeeData.EmailID}_${currentBuddyRanking.monthKey}`;
    if (sessionStorage.getItem(revealKey)) {
      if (pendingWelcomeSection) {
        setActiveSection(pendingWelcomeSection);
        setPendingWelcomeSection(null);
      }
      return;
    }

    sessionStorage.setItem(revealKey, "1");
    setShowRankCelebration(true);
  }, [
    currentBuddyRanking,
    employeeData,
    pendingWelcomeSection,
    rankingLookupComplete,
    welcomeDismissed,
  ]);

  useEffect(() => {
    const isLoading = !employeeDataLoaded || galleryLoading || storiesLoading;
    if (!isLoading && initialDataLoading) {
      const timer = setTimeout(() => setInitialDataLoading(false), 600);
      return () => clearTimeout(timer);
    }
  }, [employeeDataLoaded, galleryLoading, storiesLoading, initialDataLoading]);
/*
   const getWelcomeContent = (role: string, name: string): { greeting: string; message: string; btn: string } => {
     const h = new Date().getHours();
     const gr = h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening';
     const first = (name || '').split(' ')[0] || 'there';
     if (role === 'Super Admin') return {
       greeting: `${gr}, ${first}!`,
       message: 'You have full access to the Neuland Buddy System. Manage all new joiners, buddies, meetings, questionnaires, and holidays across all locations from one place.',
       btn: 'Go to Dashboard',
     };
     if (role === 'L&D Site Admin') return {
       greeting: `${gr}, ${first}!`,
       message: 'Welcome to your L&D dashboard. You can manage onboarding, meetings, and questionnaires for new joiners within your assigned location.',
       btn: 'View Onboarding',
     };
     if (role === 'Buddy') return {
       greeting: `${gr}, ${first}!`,
       message: 'Welcome back! Your new joiners are counting on you. Head over to Meeting Details to track sessions, answer questions, and support your NJs through their onboarding journey.',
       btn: 'View My NJs',
     };
     return {
       greeting: `${gr}, ${first}! Welcome to Neuland`,
       message: "We're so glad you're here! Your buddy has been assigned to guide you through your onboarding. Check your meeting schedule and reach out whenever you need support.",
       btn: 'See My Journey',
     };
   };
*/
  useEffect(() => {
    isMountedRef.current = true;
    (window as any).refreshBuddyApp = () => {
      if (isMountedRef.current) setRefreshTrigger((p) => p + 1);
    };
    return () => {
      isMountedRef.current = false;
      if (reactionFeedbackTimerRef.current) {
        clearTimeout(reactionFeedbackTimerRef.current);
      }
    };
  }, []);

  useEffect(() => {
    const load = async () => {
      try {
        const userEmail = props.userloginDetails?.email;
        if (!userEmail) {
          if (isMountedRef.current) setEmployeeDataLoaded(true);
          return;
        }
        const safeUserEmail = escSP(userEmail);
        console.log("safeUserEmail:", safeUserEmail);
        console.debug("[BuddyApp/loadEmployee] start", { userEmail });
        // const items = await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER).items
        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        const items = await getThresholdSafeListItems({
          web: empMasterWeb,
          listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
          select: "ID,Title,FullName,EmailID,Department,Location,GlobalID,Role,EmployeeStatus,Designation,IsSuccessStoryApprover",
          filter: `EmailID eq '${safeUserEmail}' and EmployeeStatus eq 'Active' and PositionType eq 1`,
          // filter: `EmailID eq '${safeUserEmail}' and EmployeeStatus eq 'Active'`,
          maxItems: 1,
        });
        if (isMountedRef.current) {
          if (items?.length > 0) setEmployeeData(items[0]);
          console.debug("[BuddyApp/loadEmployee] result", {
            fetched: items?.length || 0,
            id: items?.[0]?.ID,
            email: items?.[0]?.EmailID,
            roleRaw: items?.[0]?.Role,
            roleNormalized: normalizeRole(items?.[0]?.Role),
            location: items?.[0]?.Location,
            employeeStatus: items?.[0]?.EmployeeStatus,
          });
          setEmployeeDataLoaded(true);
        }
      } catch (e) {
        console.error("Error loading employee data:", e);
        if (isMountedRef.current) setEmployeeDataLoaded(true);
      }
    };
    setEmployeeDataLoaded(false);
    load();
  }, [props.context, props.userloginDetails?.email, refreshTrigger]);

  const loadHeroImage = useCallback(async () => {
    try {
      const items = await getThresholdSafeListItems({
        web: sp.web,
        listTitle: LIST_CONFIG.LIBRARIES.BuddyBannerTemplate,
        select: "ID,FileRef",
        maxItems: 1,
      });
      if (isMountedRef.current && items.length > 0) {
        const imageUrl = items[0].FileRef;
        setHeroImageUrl(imageUrl);
      }
    } catch (e) {
      console.error("Error loading hero content from SharePoint:", e);
    }
  }, [props.context]);

  useEffect(() => {
    loadHeroImage();
  }, [loadHeroImage]);

  const loadBuddyLogo = useCallback(async () => {
    try {
      const items = await getThresholdSafeListItems({
        web: sp.web,
        listTitle: LIST_CONFIG.LIBRARIES.BuddyLogo,
        select: "ID,FileRef",
        maxItems: 1,
      });
      if (isMountedRef.current && items.length > 0) {
        const logoUrl = items[0].FileRef;
        setBuddyLogoUrl(logoUrl);
      }
    } catch (e) {
      console.error("Error loading buddy logo from SharePoint:", e);
    }
  }, [props.context]);

  useEffect(() => {
    loadBuddyLogo();
  }, [loadBuddyLogo]);

  useEffect(() => {
    const load = async () => {
      if (!employeeData) return;
      try {
        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        const safeLocation = escSP(employeeData.Location);
        const safeEmail = escSP(employeeData.EmailID);
        const safeGlobalId = escSP(employeeData.GlobalID);
        const employeeRole = normalizeRole(employeeData.Role);
        console.debug("[BuddyApp/KPI] start", {
          roleRaw: employeeData.Role,
          roleNormalized: employeeRole,
          email: employeeData.EmailID,
          location: employeeData.Location,
          globalId: employeeData.GlobalID,
        });

        const getMeetingPairs = (row: any) => [
          {
            label: "1st Meeting",
            date: row.FirstInteractionDate,
            status: row.FirstMeetingStatus,
          },
          {
            label: "2nd Meeting",
            date: row.SecondInteractionDate,
            status: row.SecondMeetingStatus,
          },
          {
            label: "3rd Meeting",
            date: row.ThirdInteractionDate,
            status: row.ThirdMeetingStatus,
          },
          {
            label: "4th Meeting",
            date: row.FourthInteractionDate,
            status: row.FourthMeetingStatus,
          },
          {
            label: "5th Meeting",
            date: row.FifthInteractionDate,
            status: row.FifthMeetingStatus,
          },
          {
            label: "6th Meeting",
            date: row.SixthInteractionDate,
            status: row.SixthMeetingStatus,
          },
        ];
        const fmtDate = (v?: string) => {
          if (!v) return "N/A";
          const d = new Date(v);
          return Number.isNaN(d.getTime())
            ? "N/A"
            : d.toLocaleDateString("en-IN", { day: "2-digit", month: "short" });
        };

        if (
          employeeRole === "Super Admin" ||
          employeeRole === "L&D Site Admin"
        ) {
          // Load allowed grades from BuddyNewJoineeCriteria (Status=Yes AND Code in JM1-JM5 only)
          let allowedGrades: string[] = [];
          try {
            const criteriaItems = await getThresholdSafeListItems({
              web: sp.web,
              listTitle: LIST_CONFIG.LISTS.BuddyNewJoineeCriteria,
              select: "ID,Code,Status",
            });

            allowedGrades = (criteriaItems || [])
              .filter((item: any) => {
                const status = String(item.Status || "").trim().toLowerCase();
                const code = String(item.Code || "").trim().toLowerCase();
                return status === "yes" && /^jm[1-5]$/.test(code);
              })
              .map((item: any) => String(item.Code || "").trim().toUpperCase())
              .filter(Boolean);


          } catch (err) {
            console.error("Failed to load BuddyNewJoineeCriteria:", err);
          }

         

          const filter =
            employeeRole === "L&D Site Admin" && safeLocation
              ? `PositionType eq 1 and Location eq '${safeLocation}'`
              : "PositionType eq 1";
          const all = await getThresholdSafeListItems({
            web: empMasterWeb,
            listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
            select: "ID,EmployeeStatus,Role,AssignStatus,Location,Grade",
            filter,
          });
          const activeBuddies = all.filter(
            (e: any) => isActiveEmployee(e) && hasRole(e, "Buddy"),
          ).length;
          const newJoiners = all.filter(
            (e: any) => isActiveEmployee(e) && hasRole(e, "New Joinee"),
          ).length;
          const buddyAssigned = all.filter(
            (e: any) =>
              isActiveEmployee(e) &&
              hasRole(e, "New Joinee") &&
              isAssignedStatus(e),
          ).length;
          const buddyUnassigned = all.filter(
            (e: any) =>
              isActiveEmployee(e) &&
              hasRole(e, "New Joinee") &&
              !isAssignedStatus(e) &&
              (e.Grade && allowedGrades.includes(String(e.Grade).trim().toUpperCase())),
          ).length;
          const blankRoleActive = all.filter(
            (e: any) => isActiveEmployee(e) && !normalizeRole(e.Role),
          ).length;
          const blankAssignStatusNJs = all.filter(
            (e: any) =>
              isActiveEmployee(e) &&
              hasRole(e, "New Joinee") &&
              !normalizeStatus(e.AssignStatus),
          ).length;
          const pct =
            newJoiners > 0 ? Math.round((buddyAssigned / newJoiners) * 100) : 0;
          const scope =
            employeeRole === "L&D Site Admin" && safeLocation
              ? `${safeLocation} location`
              : "All locations";
          console.debug("[BuddyApp/KPI/admin]", {
            filter: filter || "none",
            totalFetched: all.length,
            activeBuddies,
            newJoiners,
            buddyAssigned,
            buddyUnassigned,
            blankRoleActive,
            blankAssignStatusNJs,
            sample: all
              .slice(0, 5)
              .map((e: any) => ({
                ID: e.ID,
                Role: e.Role,
                AssignStatus: e.AssignStatus,
                EmployeeStatus: e.EmployeeStatus,
                Location: e.Location,
              })),
          });
          if (isMountedRef.current)
            setKpiCards([
              {
                icon: <HowToRegOutlinedIcon sx={{ fontSize: "1.1rem" }} />,
                label: "Active Buddies",
                value: activeBuddies.toString(),
                trend: `${scope} · Verified active`,
                dir: "up",
                bg: "linear-gradient(135deg,#005c44,#00a86b)",
              },
              {
                icon: <PersonAddAltOutlinedIcon sx={{ fontSize: "1.1rem" }} />,
                label: "New Joiners",
                value: newJoiners.toString(),
                trend: `${scope} · Pending onboarding`,
                dir: "up",
                bg: "linear-gradient(135deg,#00D264,#8CFF8C)",
              },
              {
                icon: <PersonAddAlt1Icon sx={{ fontSize: "1.1rem" }} />,
                label: "NJ's Assigned",
                value: buddyAssigned.toString(),
                trend: `${buddyAssigned} of ${newJoiners} joiners matched`,
                dir: "up",
                bg: "linear-gradient(135deg,#004632,#00a86b)",
              },
              {
                icon: <PersonOffOutlinedIcon sx={{ fontSize: "1.1rem" }} />,
                label: "NJ's Unassigned",
                value: buddyUnassigned.toString(),
                trend:
                  buddyUnassigned > 0
                    ? `${buddyUnassigned} joiner(s) need a buddy`
                    : "All joiners assigned ✓",
                dir: buddyUnassigned > 0 ? "down" : "up",
                bg: "linear-gradient(135deg,#78736e,#a09b97)",
              },
              {
                icon: <TaskAltIcon sx={{ fontSize: "1.1rem" }} />,
                label: "Assignment Coverage",
                value: `${pct}%`,
                trend:
                  pct >= 80
                    ? "Excellent coverage · Keep it up!"
                    : pct >= 50
                      ? "Good progress · More assignments needed"
                      : "Low coverage · Action required",
                dir: pct >= 50 ? "up" : "down",
                bg: "linear-gradient(135deg,#004632,#00D264)",
                full: true,
              },
            ]);
          return;
        }

        if (employeeRole === "Buddy") {
          const allocs = await getThresholdSafeListItems({
            web: sp.web,
            listTitle: LIST_CONFIG.LISTS.BuddyAllocate,
            select: "ID,GlobalIDofNJ,NameOfTheNewJoinerRefColumn,BuddyEmailID,CurrentMeetingStatus,FirstInteractionDate,SecondInteractionDate,ThirdInteractionDate,FourthInteractionDate,FifthInteractionDate,SixthInteractionDate,FirstMeetingStatus,SecondMeetingStatus,ThirdMeetingStatus,FourthMeetingStatus,FifthMeetingStatus,SixthMeetingStatus",
            filter: `Status eq 'Active' and BuddyEmailID eq '${safeEmail}'`,
          });
          const today = new Date();
          today.setHours(0, 0, 0, 0);
          const weekEnd = new Date(today);
          weekEnd.setDate(weekEnd.getDate() + 7);
          let due = 0,
            overdue = 0,
            done = 0;
          (allocs || []).forEach((row: any) => {
            getMeetingPairs(row).forEach((m) => {
              if (m.status === "Completed") {
                done++;
                return;
              }
              if (!m.date) return;
              const d = new Date(m.date);
              if (Number.isNaN(d.getTime())) return;
              d.setHours(0, 0, 0, 0);
              if (d < today) overdue++;
              else if (d <= weekEnd) due++;
            });
          });
          const njIds = new Set(
            (allocs || []).map((a: any) => String(a.GlobalIDofNJ || "")),
          );
          const drafts = await getThresholdSafeListItems({
            web: sp.web,
            listTitle: LIST_CONFIG.LISTS.MEETING_ANSWERS,
            select: "ID,NJID,Status",
            filter: "Status eq 'Draft'",
          });
          const pending = (drafts || []).filter((a: any) =>
            njIds.has(String(a.NJID || "")),
          ).length;
          const total = allocs.length * 6;
          const rate = total > 0 ? Math.round((done / total) * 100) : 0;
          console.debug("[BuddyApp/KPI/buddy]", {
            filter: `Status eq 'Active' and BuddyEmailID eq '${safeEmail}'`,
            allocations: allocs.length,
            due,
            overdue,
            done,
            pendingDrafts: pending,
            totalMeetings: total,
            completionRate: rate,
          });
          if (isMountedRef.current)
            setKpiCards([
              {
                icon: <HowToRegOutlinedIcon sx={{ fontSize: "1.1rem" }} />,
                label: "My New Joinees",
                value: allocs.length.toString(),
                trend:
                  allocs.length > 0
                    ? `${allocs.length} joiner(s) under your care`
                    : "No active assignments yet",
                dir: allocs.length > 0 ? "up" : "down",
                bg: "linear-gradient(135deg,#005c44,#00a86b)",
              },
              {
                icon: <Calendar size={18} />,
                label: "Due This Week",
                value: due.toString(),
                trend:
                  due > 0
                    ? `${due} meeting(s) due in next 7 days`
                    : "No meetings due this week ✓",
                dir: due > 0 ? "down" : "up",
                bg: "linear-gradient(135deg,#00D264,#8CFF8C)",
              },
              {
                icon: <FileText size={18} />,
                label: "Pending Drafts",
                value: pending.toString(),
                trend:
                  pending > 0
                    ? `${pending} response(s) saved, not submitted`
                    : "All responses submitted ✓",
                dir: pending > 0 ? "down" : "up",
                bg: "linear-gradient(135deg,#004632,#00a86b)",
              },
              {
                icon: <PersonOffOutlinedIcon sx={{ fontSize: "1.1rem" }} />,
                label: "Overdue Meetings",
                value: overdue.toString(),
                trend:
                  overdue > 0
                    ? `${overdue} meeting(s) past scheduled date`
                    : "No overdue meetings ✓",
                dir: overdue > 0 ? "down" : "up",
                bg: "linear-gradient(135deg,#78736e,#a09b97)",
              },
              {
                icon: <TaskAltIcon sx={{ fontSize: "1.1rem" }} />,
                label: "Completion Rate",
                value: `${rate}%`,
                trend:
                  rate >= 80
                    ? `${done} of ${total} sessions done · Great work!`
                    : rate >= 50
                      ? `${done} of ${total} sessions done · Keep going`
                      : `${done} of ${total} sessions done · More needed`,
                dir: rate >= 50 ? "up" : "down",
                bg: "linear-gradient(135deg,#004632,#00D264)",
                full: true,
              },
            ]);
          return;
        }

        if (employeeRole === "New Joinee") {
          const allocs = safeGlobalId
            ? await getThresholdSafeListItems({
                web: sp.web,
                listTitle: LIST_CONFIG.LISTS.BuddyAllocate,
                select: "ID,BuddyNameRefChoiceColumn,BuddyAllocatedOn,CurrentMeetingStatus,GlobalIDofNJ,FirstInteractionDate,SecondInteractionDate,ThirdInteractionDate,FourthInteractionDate,FifthInteractionDate,SixthInteractionDate,FirstMeetingStatus,SecondMeetingStatus,ThirdMeetingStatus,FourthMeetingStatus,FifthMeetingStatus,SixthMeetingStatus",
                filter: `Status eq 'Active' and GlobalIDofNJ eq '${safeGlobalId}'`,
                maxItems: 100,
              })
            : [];
          const row = (allocs || [])[0];
          let done = 0,
            pending = 0,
            upLabel = "No upcoming",
            upTrend = "No active allocation",
            progress = 0;
          if (row) {
            const pairs = getMeetingPairs(row);
            done = pairs.filter((m) => m.status === "Completed").length;
            pending = pairs.filter((m) => m.status !== "Completed").length;
            const next = pairs.find((m) => m.status !== "Completed");
            if (next) {
              upLabel = next.label;
              const nd = fmtDate(next.date);
              upTrend =
                nd !== "N/A" ? `Scheduled on ${nd}` : "Date not scheduled yet";
            }
            progress = Math.round((done / 6) * 100);
          }
          const bName = row?.BuddyNameRefChoiceColumn || "Unassigned";
          const bTrend = row
            ? row.BuddyAllocatedOn
              ? `Assigned on ${fmtDate(row.BuddyAllocatedOn)}`
              : "Buddy has been allocated"
            : "Buddy not assigned yet";
          console.debug("[BuddyApp/KPI/newJoinee]", {
            filter: `Status eq 'Active' and GlobalIDofNJ eq '${safeGlobalId}'`,
            allocations: allocs.length,
            selectedAllocationId: row?.ID,
            done,
            pending,
            upcomingMeeting: upLabel,
            progress,
          });
          if (isMountedRef.current)
            setKpiCards([
              {
                icon: <HowToRegOutlinedIcon sx={{ fontSize: "1.1rem" }} />,
                label: "Assigned Buddy",
                value: bName,
                trend: bTrend,
                dir: row ? "up" : "down",
                bg: "linear-gradient(135deg,#005c44,#00a86b)",
              },
              {
                icon: <Calendar size={18} />,
                label: "Upcoming Meeting",
                value: upLabel,
                trend: upTrend,
                dir: row ? "up" : "down",
                bg: "linear-gradient(135deg,#00D264,#8CFF8C)",
              },
              {
                icon: <TaskAltIcon sx={{ fontSize: "1.1rem" }} />,
                label: "Meetings Completed",
                value: `${done}/6`,
                trend:
                  done === 6
                    ? "All sessions complete 🎉"
                    : done > 0
                      ? `${6 - done} session(s) remaining`
                      : "Your onboarding journey begins",
                dir: "up",
                bg: "linear-gradient(135deg,#004632,#00a86b)",
              },
              {
                icon: <HourglassTopIcon sx={{ fontSize: "1.1rem" }} />,
                label: "Pending Meetings",
                value: pending.toString(),
                trend:
                  pending === 0
                    ? "All meetings completed ✓"
                    : `${pending} session(s) left to complete`,
                dir: pending > 0 ? "down" : "up",
                bg: "linear-gradient(135deg,#78736e,#a09b97)",
              },
              {
                icon: <TrendingUp size={18} />,
                label: "Onboarding Progress",
                value: `${progress}%`,
                trend: row?.CurrentMeetingStatus
                  ? `Current status: ${row.CurrentMeetingStatus}`
                  : progress === 0
                    ? "Ready to start your journey!"
                    : `${progress}% of onboarding complete`,
                dir: "up",
                bg: "linear-gradient(135deg,#004632,#00D264)",
                full: true,
              },
            ]);
          return;
        }

        if (isMountedRef.current) setKpiCards(DEFAULT_KPI_CARDS);
      } catch (e) {
        console.error("Error loading KPI:", e);
        if (isMountedRef.current) setKpiCards(DEFAULT_KPI_CARDS);
      }
    };
    load();
  }, [employeeData]);

  const resetTimer = useCallback(() => {
    if (slideTimer.current) clearInterval(slideTimer.current);
    if (!sliderPaused) {
      slideTimer.current = setInterval(() => {
        const len = activeGalleryRef.current.length;
        if (len > 0) setSlideIdx((p) => (p + 1) % len);
      }, 4000);
    }
  }, [sliderPaused]);

  useEffect(() => {
    resetTimer();
    return () => {
      if (slideTimer.current) clearInterval(slideTimer.current);
    };
  }, [resetTimer, sliderPaused, activeGalleryItems.length]);

  const pauseAllSliderVideos = () => {
    const videos = document.querySelectorAll('video');
    videos.forEach((v) => {
      if (!v.paused && !v.ended) v.pause();
    });
  };

  const prevSlide = useCallback(() => {
    const l = activeGalleryRef.current.length;
    if (!l) return;
    setSlideIdx((p) => (p - 1 + l) % l);
    setSliderPaused(false); 
    pauseAllSliderVideos();
    resetTimer();
  }, [resetTimer]);

  const nextSlide = useCallback(() => {
    const l = activeGalleryRef.current.length;
    if (!l) return;
    setSlideIdx((p) => (p + 1) % l);
    setSliderPaused(false); 
    pauseAllSliderVideos();
    resetTimer();
  }, [resetTimer]);
  const goToSlide = useCallback(
    (i: number) => {
      setSlideIdx(i);
      resetTimer();
    },
    [resetTimer],
  );

  useEffect(() => {
    if (activeSection === "home") {
      setRefreshTrigger((p) => p + 1);
    }
  }, [activeSection]);

  useEffect(() => {
    const link = document.createElement("link");
    link.href =
      "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap";
    link.rel = "stylesheet";
    document.head.appendChild(link);
    return () => {
      if (document.head.contains(link)) document.head.removeChild(link);
    };
  }, []);

  const getRequestDigest = useCallback(async (): Promise<string> => {
    const digest = (props.context.pageContext as any).legacyPageContext
      ?.formDigestValue;
    if (digest) return digest;
    const webUrl = props.context.pageContext.web.absoluteUrl;
    const res = await fetch(`${webUrl}/_api/contextinfo`, { method: "POST" });
    const data = await res.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  }, [props.context]);

  const webAbsoluteUrl = props.context.pageContext.web.absoluteUrl;
  const employeeRole = normalizeRole(employeeData?.Role);
  const isSuperAdmin = employeeRole === "Super Admin";
  const isSuccessStoryApprover =
    isSuperAdmin &&
    (employeeData?.IsSuccessStoryApprover === true ||
      employeeData?.IsSuccessStoryApprover === "Yes");
  const loginEmployeeStatus = normalizeStatus(employeeData?.EmployeeStatus);
  const isLoginEmployeeActive = loginEmployeeStatus === "active";
  const loggedInUserEmail = String(
    props.userloginDetails?.email || employeeData?.EmailID || "",
  )
    .trim()
    .toLowerCase();

  const loadGallery = useCallback(async () => {
    setGalleryLoading(true);
    try {
      const galleryRows = await getThresholdSafeListItems({
        web: sp.web,
        listTitle: LIST_CONFIG.LIBRARIES.BuddyGallery,
        select: "ID,FileRef,Title,ImgTitle,Status",
        filter: isSuperAdmin ? undefined : "Status eq 'Active'",
      });
      const allItems: GalleryItem[] = (galleryRows || [])
        .filter((i: any) => i.Status && i.Status.trim() !== "")
        .map((i: any) => {
          const fileRef: string = i.FileRef || "";
          const isVideo = fileRef.toLowerCase().endsWith(".mp4");
          return {
            img: fileRef,
            label: i.ImgTitle || i.Title || "Untitled",
            status: i.Status || "Active",
            itemId: i.ID,
            fileType: isVideo ? "video" : "image",
          };
        });
      if (isMountedRef.current) {
        const active = allItems.filter((i) => i.status === "Active");
        const inactive = allItems.filter((i) => i.status !== "Active");
        setActiveGalleryItems(active);
        activeGalleryRef.current = active;
        setInactiveGalleryItems(inactive);
        setSlideIdx((p) =>
          active.length > 0 ? Math.min(p, active.length - 1) : 0,
        );
        setGalleryLoading(false);
      }
    } catch (e) {
      console.error("Error loading gallery:", e);
      if (isMountedRef.current) setGalleryLoading(false);
    }
  }, [webAbsoluteUrl, isSuperAdmin]);

  useEffect(() => {
    loadGallery();
  }, [loadGallery]);

  const [uploadModal, setUploadModal] = useState(false);
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [uploadPreview, setUploadPreview] = useState<string>("");
  const [uploadTitle, setUploadTitle] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const dropRef = useRef<HTMLDivElement>(null);

  const openUploadModal = useCallback(() => {
    setUploadFile(null);
    setUploadPreview("");
    setUploadTitle("");
    setIsDragging(false);
    setIsUploading(false);
    setUploadModal(true);
  }, []);
  const closeUploadModal = useCallback(() => {
    setUploadModal(false);
    setUploadFile(null);
    setUploadPreview("");
    setUploadTitle("");
    setIsUploading(false);
  }, []);
  const IMAGE_EXTENSIONS = ["jpg", "jpeg", "png", "gif", "webp"];
  const IMAGE_MAX_SIZE = 10 * 1024 * 1024; 
  const VIDEO_MAX_SIZE = 25 * 1024 * 1024; 
  const processFile = useCallback((file: File) => {
    const ext = file.name.split(".").pop()?.toLowerCase() || "";
    const isImage = IMAGE_EXTENSIONS.includes(ext);
    const isVideo = file.type === "video/mp4" || ext === "mp4";
    if (!isImage && !isVideo) {
      showSnackbar("Only image files (WEBP, JPG, JPEG, PNG, GIF) and MP4 videos are allowed.", "error", 4000);
      return;
    }
    if (isImage && file.size > IMAGE_MAX_SIZE) {
      const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
      showSnackbar(`Image size (${sizeMB} MB) exceeds the 10 MB limit. Please choose a smaller image.`, "error", 4000);
      return;
    }
    if (isVideo && file.size > VIDEO_MAX_SIZE) {
      const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
      showSnackbar(`Video size (${sizeMB} MB) exceeds the 25 MB limit. Please choose a smaller video.`, "error", 4000);
      return;
    }
    setUploadFile(file);
    const reader = new FileReader();
    reader.onload = (e) => {
      if (isMountedRef.current) setUploadPreview(e.target?.result as string);
    };
    reader.readAsDataURL(file);
  }, [showSnackbar]);
  const handleFileInput = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const f = e.target.files?.[0];
      if (f) processFile(f);
      e.target.value = "";
    },
    [processFile],
  );
  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);
  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);
  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      const f = e.dataTransfer.files[0];
      if (f) processFile(f);
    },
    [processFile],
  );

  const handleUploadConfirm = useCallback(async () => {
    if (!uploadFile || !uploadPreview) return;
    const cleaned = cleanTitle(uploadTitle);
    if (!cleaned) {
      showSnackbar("Please enter a title for the file", "error", 3000);
      return;
    }
    
    const ext = uploadFile.name.split(".").pop()?.toLowerCase() || "";
    const isImage = IMAGE_EXTENSIONS.includes(ext);
    const isVideo = uploadFile.type === "video/mp4" || ext === "mp4";
    if (isImage && uploadFile.size > IMAGE_MAX_SIZE) {
      const sizeMB = (uploadFile.size / (1024 * 1024)).toFixed(2);
      showSnackbar(`Image size (${sizeMB} MB) exceeds the 10 MB limit. Please choose a smaller image.`, "error", 4000);
      return;
    }
    if (isVideo && uploadFile.size > VIDEO_MAX_SIZE) {
      const sizeMB = (uploadFile.size / (1024 * 1024)).toFixed(2);
      showSnackbar(`Video size (${sizeMB} MB) exceeds the 25 MB limit. Please choose a smaller video.`, "error", 4000);
      return;
    }
    
    setIsUploading(true);
    // const isVideo = uploadFile.type === "video/mp4";
    const uploadMessage = "Uploading content to gallery…";
    setGlobalLoader({
      visible: true,
      variant: "upload",
      message: uploadMessage,
    });
    const webUrl = props.context.pageContext.web.absoluteUrl;
    const webServerRelativeUrl =
      props.context.pageContext.web.serverRelativeUrl;
    try {
      const fileBuffer = await uploadFile.arrayBuffer();
      const randomNum = Math.floor(Math.random() * 900000) + 100000;
      const parts = uploadFile.name.split(".");
      const ext = parts.pop();
      const base = parts.join(".");
      const uniqueName = `${base}-${randomNum}.${ext}`;
      const folderPath = `${webServerRelativeUrl}/${LIST_CONFIG.LIBRARIES.BuddyGallery}`;
      const digest = await getRequestDigest();
      const uploadUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')/Files/add(url='${encodeURIComponent(uniqueName)}',overwrite=true)`;
      const uploadResp = await fetch(uploadUrl, {
        method: "POST",
        headers: {
          Accept: "application/json",
          "Content-Type": "application/octet-stream",
          "X-RequestDigest": digest,
        },
        body: fileBuffer,
      });

      if (!uploadResp.ok) {
        throw new Error(`Upload failed: ${uploadResp.status}`);
      }

      const uploadedFile = await uploadResp.json();
      const fileServerRelUrl =
        uploadedFile.ServerRelativeUrl || uploadedFile.d?.ServerRelativeUrl;

      if (!fileServerRelUrl) {
        throw new Error("Upload completed but the file URL could not be determined.");
      }

      await fetch(
        `${webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(fileServerRelUrl)}')/ListItemAllFields`,
        {
          method: "PATCH",
          headers: {
            Accept: "application/json",
            "Content-Type": "application/json",
            "X-RequestDigest": digest,
            "If-Match": "*",
          },
          body: JSON.stringify({ ImgTitle: cleaned, Status: "Active" }),
        },
      );
      if (!isMountedRef.current) return;
      setGlobalLoader({ visible: false, variant: "upload" });
      closeUploadModal();
      showSnackbar("Content uploaded successfully!", "success", 3000);
      await loadGallery();
    } catch (e) {
      console.error("Upload error:", e);
      if (isMountedRef.current) {
        setIsUploading(false);
        setGlobalLoader({ visible: false, variant: "upload" });
        showSnackbar("Failed to upload content. Please try again.", "error", 4000);
      }
    }
  }, [
    uploadFile,
    uploadPreview,
    uploadTitle,
    props.context,
    getRequestDigest,
    resetTimer,
    closeUploadModal,
    showSnackbar,
  ]);

  type ConfirmAction =
    | GalleryAction
    | "deactivate-story"
    | "activate-story"
    | "approve-story"
    | "reject-story";
  const [reviewComment, setReviewComment] = useState("");
  const [confirmDialog, setConfirmDialog] = useState<{
    open: boolean;
    type: "gallery" | "story";
    action: ConfirmAction | "";
    itemIdx: number | null;
    storyId: number | null;
    fromInactive: boolean;
  }>({
    open: false,
    type: "gallery",
    action: "",
    itemIdx: null,
    storyId: null,
    fromInactive: false,
  });

  const cancelAction = useCallback(() => {
    setReviewComment("");
    setConfirmDialog({
      open: false,
      type: "gallery",
      action: "",
      itemIdx: null,
      storyId: null,
      fromInactive: false,
    });
  }, []);

  const handleDeactivateGallery = useCallback((idx: number) => {
    setConfirmDialog({
      open: true,
      type: "gallery",
      action: "deactivate",
      itemIdx: idx,
      storyId: null,
      fromInactive: false,
    });
  }, []);
  const handleActivateGallery = useCallback((idx: number) => {
    setConfirmDialog({
      open: true,
      type: "gallery",
      action: "activate",
      itemIdx: idx,
      storyId: null,
      fromInactive: true,
    });
  }, []);
  const handleDeactivateStory = useCallback((storyId: number) => {
    setConfirmDialog({
      open: true,
      type: "story",
      action: "deactivate-story",
      itemIdx: null,
      storyId,
      fromInactive: false,
    });
  }, []);
  const handleActivateStory = useCallback((storyId: number) => {
    setConfirmDialog({
      open: true,
      type: "story",
      action: "activate-story",
      itemIdx: null,
      storyId,
      fromInactive: true,
    });
  }, []);
  const handleReviewStory = useCallback(
    (storyId: number, decision: "approve-story" | "reject-story") => {
      setReviewComment("");
      setConfirmDialog({
        open: true,
        type: "story",
        action: decision,
        itemIdx: null,
        storyId,
        fromInactive: false,
      });
    },
    [],
  );

  const confirmAction = useCallback(async () => {
    const { type, action, itemIdx, storyId, fromInactive } = confirmDialog;
    if (!action) {
      cancelAction();
      return;
    }
    if (action === "reject-story" && !reviewComment.trim()) {
      showSnackbar("Rejection reason is required.", "warning", 3000);
      return;
    }
    if (
      (action === "approve-story" || action === "reject-story") &&
      !isSuccessStoryApprover
    ) {
      cancelAction();
      showSnackbar("You are not authorized to review stories.", "error", 3000);
      return;
    }
    const webUrl = props.context.pageContext.web.absoluteUrl;
    setConfirmDialog({
      open: false,
      type: "gallery",
      action: "",
      itemIdx: null,
      storyId: null,
      fromInactive: false,
    });
    setGlobalLoader({
      visible: true,
      variant: "update",
      message: "Applying changes…",
    });

    try {
      if (type === "gallery" && itemIdx !== null) {
        const source = fromInactive ? inactiveGalleryItems : activeGalleryItems;
        const item = source[itemIdx];
        if (!item || !item.itemId) {
          showSnackbar("Cannot update: item ID not found", "error", 3000);
          return;
        }
        const newStatus = action === "activate" ? "Active" : "Inactive";
        const digest = await getRequestDigest();
        const itemUrl = `${webUrl}/_api/web/lists/getbytitle('${LIST_CONFIG.LIBRARIES.BuddyGallery}')/items(${item.itemId})`;
        const res = await fetch(itemUrl, {
          method: "PATCH",
          headers: {
            Accept: "application/json",
            "Content-Type": "application/json",
            "X-RequestDigest": digest,
            "X-HTTP-Method": "MERGE",
            "If-Match": "*",
          },
          body: JSON.stringify({ Status: newStatus }),
        });
        if (!res.ok) throw new Error(`Status update failed: ${res.status}`);
        if (!isMountedRef.current) return;
        if (action === "deactivate")
          showSnackbar(
            "Content set to Inactive — hidden from gallery",
            "success",
            3000,
          );
        else showSnackbar("Content restored to Active gallery", "success", 3000);
        setGlobalLoader({ visible: false, variant: "update" });
        await loadGallery();
        return;
      }

      if (type === "story" && storyId !== null) {
        if (action === "approve-story" || action === "reject-story") {
          const storyItemUrl = `${webUrl}/_api/web/lists/getbytitle('${LIST_CONFIG.LISTS.MemberSuccessStories}')/items(${storyId})`;
          const currentResponse = await fetch(
            `${storyItemUrl}?$select=ID,ApprovalStatus`,
            { headers: { Accept: "application/json" } },
          );
          if (!currentResponse.ok) {
            throw new Error(
              `Unable to read current review state: ${currentResponse.status}`,
            );
          }
          const current = await currentResponse.json();
          const currentDecision = String(
            current?.ApprovalStatus || "Approved",
          ).trim();

          if (currentDecision !== "Pending") {
            if (isMountedRef.current) {
              setGlobalLoader({ visible: false, variant: "update" });
              setPendingStories((prev) =>
                prev.filter((story) => story.id !== storyId),
              );
              showSnackbar(
                `This story was already ${currentDecision.toLowerCase()} by another reviewer.`,
                "warning",
                4000,
              );
            }
            await loadStoriesRef.current();
            return;
          }

          const eTag = currentResponse.headers.get("ETag") || "";
          if (!eTag) {
            throw new Error("Unable to enforce first-reviewer concurrency.");
          }

          const approved = action === "approve-story";
          const digest = await getRequestDigest();
          const updateResponse = await fetch(storyItemUrl, {
            method: "PATCH",
            headers: {
              Accept: "application/json",
              "Content-Type": "application/json",
              "X-RequestDigest": digest,
              "X-HTTP-Method": "MERGE",
              "If-Match": eTag,
            },
            body: JSON.stringify({
              ApprovalStatus: approved ? "Approved" : "Rejected",
              Status: approved ? "Active" : "Inactive",
              ReviewedByName: employeeData?.FullName || "",
              ReviewedByEmail: employeeData?.EmailID || "",
              ReviewedOn: new Date().toISOString(),
              ReviewComment: reviewComment.trim(),
              RunWF: "Yes",
            }),
          });
          if (!updateResponse.ok) {
            throw new Error(`Review update failed: ${updateResponse.status}`);
          }

          if (!isMountedRef.current) return;
          setGlobalLoader({ visible: false, variant: "update" });
          setPendingStories((prev) =>
            prev.filter((story) => story.id !== storyId),
          );
          showSnackbar(
            approved
              ? "Story approved and published successfully."
              : "Story rejected successfully.",
            "success",
            3500,
          );
          await loadStoriesRef.current();
          return;
        }

        const newStatus = action === "activate-story" ? "Active" : "Inactive";
        await sp.web.lists
          .getByTitle(LIST_CONFIG.LISTS.MemberSuccessStories)
          .items.getById(storyId)
          .update({ Status: newStatus });
        if (!isMountedRef.current) return;
        if (action === "deactivate-story") {
          setStories((prev) => {
            const moved = prev.find((s) => s.id === storyId);
            if (moved)
              setInactiveStories((ip) => [
                { ...moved, status: "Inactive".toLowerCase() },
                ...ip,
              ]);
            return prev.filter((s) => s.id !== storyId);
          });
          showSnackbar(
            "Story set to Inactive — hidden from all users",
            "success",
            3000,
          );
        } else {
          setInactiveStories((prev) => {
            const moved = prev.find((s) => s.id === storyId);
            if (moved)
              setStories((ap) => [
                { ...moved, status: "Active".toLowerCase() },
                ...ap,
              ]);
            return prev.filter((s) => s.id !== storyId);
          });
          showSnackbar("Story restored to Active", "success", 3000);
        }
        setGlobalLoader({ visible: false, variant: "update" });
      }
    } catch (e) {
      console.error("Action error:", e);
      if (isMountedRef.current) {
        setGlobalLoader({ visible: false, variant: "update" });
        if (action === "approve-story" || action === "reject-story") {
          showSnackbar(
            "Review could not be saved. Another reviewer may have acted first; the queue has been refreshed.",
            "warning",
            4500,
          );
          await loadStoriesRef.current();
        } else {
          showSnackbar("Action failed. Please try again.", "error", 3000);
        }
      }
    }
  }, [
    confirmDialog,
    activeGalleryItems,
    inactiveGalleryItems,
    props.context,
    getRequestDigest,
    showSnackbar,
    cancelAction,
    reviewComment,
    isSuccessStoryApprover,
    employeeData,
  ]);

  const [lightbox, setLightbox] = useState<LightboxState>({
    open: false,
    img: "",
    label: "",
    fileType: "image",
  });
  const openLightbox = useCallback(
    (img: string, label: string, fileType: "image" | "video" = "image") =>
      setLightbox({ open: true, img, label, fileType }),
    [],
  );
  const closeLightbox = useCallback(
    () => setLightbox((p) => ({ ...p, open: false })),
    [],
  );
  useEffect(() => {
    if (!lightbox.open && !uploadModal) return;
    const h = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        closeLightbox();
        closeUploadModal();
      }
    };
    window.addEventListener("keydown", h);
    return () => window.removeEventListener("keydown", h);
  }, [lightbox.open, uploadModal, closeLightbox, closeUploadModal]);

  const loadStoriesFromSP = useCallback(async () => {
    setStoriesLoading(true);
    try {
      const currentEmail = String(employeeData?.EmailID || "")
        .trim()
        .toLowerCase();
      const emailFilter = escSP(employeeData?.EmailID || "");
      const storyFilter = isSuperAdmin
        ? undefined
        : emailFilter
          ? `(Status eq 'Active' or EmployeeEmail eq '${emailFilter}')`
          : "Status eq 'Active'";
      const items = await getThresholdSafeListItems({
        web: sp.web,
        listTitle: LIST_CONFIG.LISTS.MemberSuccessStories,
        select: "ID,EmployeeName,EmployeeEmail,SuccessStory,Department,Rating,BuddyName,BuddyEmailID,BuddyGlobalID,EmployeeLocation,Status,AuthorRole,ApprovalStatus,ReviewedByName,ReviewedByEmail,ReviewedOn,ReviewComment,ReactionsJson,Created",
        filter: storyFilter,
      });
      if (isMountedRef.current) {
        const mapStory = (i: any): Story => ({
          id: i.ID,
          name: i.EmployeeName || "Unknown",
          team: i.Department || "Team",
          quote: i.SuccessStory || "",
          rating: i.Rating || 5,
          status: (i.Status || "Active").trim(),
          buddyName: i.BuddyName || "",
          EmailID: i.BuddyEmailID || "",
          buddyGlobalID: i.BuddyGlobalID || "",
          employeeLocation: i.EmployeeLocation || "",
          authorRole: i.AuthorRole || "Team Member",
          employeeEmail: i.EmployeeEmail || "",
          // Blank means a legacy story created before the approval process.
          approvalStatus: (i.ApprovalStatus || "Approved") as
            | "Pending"
            | "Approved"
            | "Rejected",
          reviewedByName: i.ReviewedByName || "",
          reviewedByEmail: i.ReviewedByEmail || "",
          reviewedOn: i.ReviewedOn || "",
          reviewComment: i.ReviewComment || "",
          createdOn: i.Created || "",
          reactionsJson: i.ReactionsJson || "",
        });
        const all = (items || []).map(mapStory);
        const publishedStories = all.filter(
          (s) =>
            s.approvalStatus === "Approved" &&
            (s.status || "").toLowerCase() === "active",
        );
        setStories(publishedStories);
        const nextReactionSummaries: Record<number, StoryReactionSummary> = {};
        publishedStories.forEach((story) => {
          if (!story.id) return;
          nextReactionSummaries[story.id] = summarizeStoryReactions(
            parseStoryReactions(story.reactionsJson),
            loggedInUserEmail,
          );
        });
        setStoryReactionSummaries(nextReactionSummaries);
        setInactiveStories(
          all.filter(
            (s) =>
              s.approvalStatus === "Approved" &&
              (s.status || "").toLowerCase() !== "active",
          ),
        );
        setPendingStories(
          isSuccessStoryApprover
            ? all.filter((s) => s.approvalStatus === "Pending")
            : [],
        );
        setApprovedReviewStories(
          isSuccessStoryApprover
            ? all.filter(
                (s) =>
                  s.approvalStatus === "Approved" &&
                  Boolean(s.reviewedOn || s.reviewedByEmail),
              )
            : [],
        );
        setRejectedReviewStories(
          isSuccessStoryApprover
            ? all.filter((s) => s.approvalStatus === "Rejected")
            : [],
        );
        setMySubmissions(
          all.filter(
            (s) =>
              String(s.employeeEmail || "").trim().toLowerCase() ===
                currentEmail &&
              (s.approvalStatus === "Pending" ||
                s.approvalStatus === "Rejected"),
          ),
        );
        setStoriesLoading(false);
      }
    } catch (e) {
      console.error("Error loading stories:", e);
      if (isMountedRef.current) setStoriesLoading(false);
    }
  }, [
    isSuperAdmin,
    isSuccessStoryApprover,
    employeeData?.EmailID,
    loggedInUserEmail,
  ]);

  useEffect(() => {
    loadStoriesRef.current = loadStoriesFromSP;
  }, [loadStoriesFromSP]);

  useEffect(() => {
    loadStoriesFromSP();
  }, [loadStoriesFromSP]); 

  const toggleReviewCenter = useCallback(() => {
    const nextOpen = !showReviewCenter;
    setShowReviewCenter(nextOpen);
    if (nextOpen) {
      setShowStoryForm(false);
      setShowMyRequests(false);
      setReviewTab("Pending");
      setReviewPage(0);
    }
  }, [showReviewCenter]);

  const toggleMyRequests = useCallback(() => {
    const nextOpen = !showMyRequests;
    setShowMyRequests(nextOpen);
    if (nextOpen) {
      setShowStoryForm(false);
      setShowReviewCenter(false);
      setMyRequestPage(0);
    }
  }, [showMyRequests]);

  const playReactionFeedback = useCallback(
    (storyId: number, reactionType: ReactionType, removed: boolean): void => {
      const emoji =
        REACTION_OPTIONS.find((option) => option.type === reactionType)?.emoji ||
        "✨";
      if (reactionFeedbackTimerRef.current) {
        clearTimeout(reactionFeedbackTimerRef.current);
      }
      setReactionFeedback({
        storyId,
        emoji,
        mode: removed ? "removed" : "selected",
        token: Date.now(),
      });
      reactionFeedbackTimerRef.current = setTimeout(() => {
        if (isMountedRef.current) setReactionFeedback(null);
      }, 850);
    },
    [],
  );

  const toggleStoryReaction = useCallback(
    async (storyId: number, reactionType: ReactionType): Promise<void> => {
      if (!loggedInUserEmail) {
        showSnackbar("Please sign in to react to a story.", "warning", 3000);
        return;
      }

      if (reactionSavingStoryIdsRef.current.has(storyId)) return;
      reactionSavingStoryIdsRef.current.add(storyId);
      setReactionSavingStoryIds((current) => [...current, storyId]);

      try {
        const digest = await getRequestDigest();
        const storyItemUrl = `${webAbsoluteUrl}/_api/web/lists/getbytitle('${LIST_CONFIG.LISTS.MemberSuccessStories}')/items(${storyId})`;
        let savedReactions: StoryReactionRecord[] | null = null;
        let removedReaction = false;

        for (let attempt = 0; attempt < 3; attempt += 1) {
          const currentResponse = await fetch(
            `${storyItemUrl}?$select=ID,ReactionsJson`,
            { headers: { Accept: "application/json" } },
          );
          if (!currentResponse.ok) {
            throw new Error(
              `Unable to read story reactions: ${currentResponse.status}`,
            );
          }

          const responseJson = await currentResponse.json();
          const currentItem = responseJson?.d || responseJson;
          const eTag =
            currentResponse.headers.get("ETag") ||
            currentItem?.["@odata.etag"] ||
            currentItem?.__metadata?.etag;
          if (!eTag) {
            throw new Error("Unable to protect the reaction update.");
          }

          const currentReactions = parseStoryReactions(
            currentItem?.ReactionsJson,
          );
          const existingReaction = currentReactions.find(
            (reaction) => reaction.userEmail === loggedInUserEmail,
          );
          let nextReactions: StoryReactionRecord[];

          removedReaction = existingReaction?.reactionType === reactionType;
          if (removedReaction) {
            nextReactions = currentReactions.filter(
              (reaction) => reaction.userEmail !== loggedInUserEmail,
            );
          } else {
            const nextReaction: StoryReactionRecord = {
              userEmail: loggedInUserEmail,
              reactionType,
            };
            nextReactions = existingReaction
              ? currentReactions.map((reaction) =>
                  reaction.userEmail === loggedInUserEmail
                    ? nextReaction
                    : reaction,
                )
              : [...currentReactions, nextReaction];
          }

          const updateResponse = await fetch(storyItemUrl, {
            method: "PATCH",
            headers: {
              Accept: "application/json",
              "Content-Type": "application/json",
              "X-RequestDigest": digest,
              "X-HTTP-Method": "MERGE",
              "If-Match": eTag,
            },
            body: JSON.stringify({
              ReactionsJson: JSON.stringify(nextReactions),
            }),
          });

          if (updateResponse.status === 412) continue;
          if (!updateResponse.ok) {
            throw new Error(
              `Reaction update failed: ${updateResponse.status}`,
            );
          }

          savedReactions = nextReactions;
          break;
        }

        if (!savedReactions) {
          throw new Error(
            "Reaction changed repeatedly by another user. Please retry.",
          );
        }

        const reactionsJson = JSON.stringify(savedReactions);
        setStoryReactionSummaries((current) => ({
          ...current,
          [storyId]: summarizeStoryReactions(
            savedReactions!,
            loggedInUserEmail,
          ),
        }));
        setStories((current) =>
          current.map((story) =>
            story.id === storyId ? { ...story, reactionsJson } : story,
          ),
        );
        playReactionFeedback(storyId, reactionType, removedReaction);
      } catch (error) {
        console.error("Unable to save success story reaction:", error);
        showSnackbar(
          "Reaction could not be saved. Please try again.",
          "error",
          3500,
        );
      } finally {
        reactionSavingStoryIdsRef.current.delete(storyId);
        if (isMountedRef.current) {
          setReactionSavingStoryIds((current) =>
            current.filter((id) => id !== storyId),
          );
        }
      }
    },
    [
      loggedInUserEmail,
      getRequestDigest,
      playReactionFeedback,
      showSnackbar,
      webAbsoluteUrl,
    ],
  );

  const [newStory, setNewStory] = useState<Story>({
    name: "",
    team: "",
    quote: "",
    rating: 5,
    buddyName: "",
    EmailID: "",
    buddyGlobalID: "",
    authorRole: "",
  });
  const [storyRating, setStoryRating] = useState<number>(5);
  const [buddySearchResults, setBuddySearchResults] = useState<any[]>([]);
  const [showBuddySearch, setShowBuddySearch] = useState(false);
  const buddySearchRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!showBuddySearch) return;
    const h = (e: MouseEvent) => {
      if (
        buddySearchRef.current &&
        !buddySearchRef.current.contains(e.target as Node)
      )
        setShowBuddySearch(false);
    };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, [showBuddySearch]);

  useEffect(() => {
    if (showStoryForm && employeeData && !newStory.name) {
      setNewStory((p) => ({
        ...p,
        name: employeeData.FullName || "",
        team: employeeData.Department || "",
        rating: 5,
        authorRole: employeeData.Role || "Team Member",
      }));
      setStoryRating(5);
    }
  }, [showStoryForm, employeeData]);

  const handleAddStory = useCallback(async () => {
    if (!newStory.name.trim() || !newStory.quote.trim()) return;
    if (newStory.buddyName && !newStory.EmailID) {
      showSnackbar(
        "Buddy selected but email not found. Please pick from dropdown.",
        "error",
        4000,
      );
      return;
    }
    if (
      newStory.EmailID &&
      employeeData?.EmailID &&
      newStory.EmailID.toLowerCase() === employeeData.EmailID.toLowerCase()
    ) {
      showSnackbar(
        "Buddy email cannot be the same as your own.",
        "error",
        4000,
      );
      return;
    }
    if (!employeeData) {
      showSnackbar("Employee data not loaded", "error", 3000);
      return;
    }
    setGlobalLoader({
      visible: true,
      variant: "submit",
      message: "Saving your success story…",
    });
    try {
      const result = await sp.web.lists
        .getByTitle(LIST_CONFIG.LISTS.MemberSuccessStories)
        .items.add({
          GlobalID: employeeData.GlobalID || "",
          Title: cleanWhitespace(newStory.name),
          EmployeeName: cleanWhitespace(newStory.name),
          EmployeeEmail: employeeData.EmailID || "",
          Department:
            cleanWhitespace(newStory.team) || employeeData.Department || "",
          Designation: employeeData.Designation || "",
          EmployeeLocation: employeeData.Location || "",
          SuccessStory: cleanWhitespace(newStory.quote),
          Rating: storyRating,
          BuddyName: newStory.buddyName
            ? cleanWhitespace(newStory.buddyName)
            : "",
          BuddyEmailID: newStory.EmailID || "",
          BuddyGlobalID: newStory.buddyGlobalID || "",
          AuthorRole: newStory.authorRole || employeeData.Role || "Team Member",
          ApprovalStatus: "Pending",
          Status: "Inactive",
          RunWF: "Yes",
        });
      setMySubmissions((prev) => [
        {
          ...newStory,
          id: result.data.Id,
          rating: storyRating,
          status: "Inactive",
          employeeEmail: employeeData.EmailID || "",
          approvalStatus: "Pending",
          authorRole: newStory.authorRole || employeeData.Role || "Team Member",
        },
        ...prev,
      ]);
      setNewStory({
        name: "",
        team: "",
        quote: "",
        rating: 5,
        buddyName: "",
        EmailID: "",
        buddyGlobalID: "",
        authorRole: "",
      });
      setStoryRating(5);
      setShowStoryForm(false);
      setGlobalLoader({ visible: false, variant: "submit" });
      showSnackbar(
        "Success story submitted for approval.",
        "success",
        3500,
      );
      await loadStoriesFromSP();
    } catch (e) {
      console.error("Error saving story:", e);
      setGlobalLoader({ visible: false, variant: "submit" });
      showSnackbar("Failed to save story.", "error", 3000);
    }
  }, [newStory, storyRating, employeeData, showSnackbar, loadStoriesFromSP]);

  const handleBuddySearch = useCallback(
    async (term: string) => {
      if (!term || term.length < 2) {
        setBuddySearchResults([]);
        return;
      }
      if (!employeeData?.Location || !employeeData?.EmailID) {
        setBuddySearchResults([]);
        return;
      }
      try {
        const t = term.replace(/'/g, "''");
        const location = String(employeeData.Location).replace(/'/g, "''");
        const currentUserEmail = String(employeeData.EmailID).replace(
          /'/g,
          "''",
        );
        // const res = await sp.web.lists.getByTitle(LIST_CONFIG.LISTS.EMPLOYEE_MASTER).items
        const empMasterWeb = Web(LIST_CONFIG.SITENAME.BUDDY_EMP_MASTER_SITE);
        const res = await getThresholdSafeListItems({
          web: empMasterWeb,
          listTitle: LIST_CONFIG.LISTS.EMPLOYEE_MASTER,
          select: "ID,FullName,EmailID,GlobalID,Department,Location,EmployeeStatus",
          filter: `PositionType eq 1 and EmployeeStatus eq 'Active' and (substringof('${t}',FullName) or substringof('${t}',EmailID) or substringof('${t}',GlobalID)) and Location eq '${location}' and EmailID ne '${currentUserEmail}'`,
          maxItems: 5,
        });
        if (isMountedRef.current) setBuddySearchResults(res || []);
      } catch (e) {
        console.error("Error searching buddies:", e);
      }
    },
    [employeeData],
  );

  
  const handleBuddySelect = useCallback((buddy: any) => {
    setNewStory((p) => ({
      ...p,
      buddyName: buddy.FullName,
      EmailID: buddy.EmailID,
      buddyGlobalID: buddy.GlobalID,
    }));
    setBuddySearchResults([]);
    setShowBuddySearch(false);
  }, []);

  const [showInactiveGallery, setShowInactiveGallery] = useState(false);
  const [showInactiveStories, setShowInactiveStories] = useState(false);

  const handleNavigation = useCallback((sectionId: string) => {
    if (sectionId === "reports") {
      const employeeRole = normalizeRole(employeeData?.Role);
      
      // Location-to-report mapping
      const locationReportMap: Record<string, string> = {
        "R&D": "RnDReport.aspx",
        "Corporate Office": "CorporateOfficeReport.aspx",
        "Unit 1": "Unit1Report.aspx",
        "Unit 2": "Unit2Report.aspx",
        "Unit 3": "Unit3Report.aspx",
      };

      // For Super Admin, show reports without location restriction
      if (employeeRole === "Super Admin") {
        const reportPageUrl = `${webAbsoluteUrl}/SitePages/neuland-report.aspx`;
        window.open(reportPageUrl, '_blank');
        return;
      }

      // For L&D Site Admin, route to location-specific report
      if (employeeRole === "L&D Site Admin") {
        const userLocation = employeeData?.Location;
        if (!userLocation) {
          showSnackbar("No location assigned to your account. Cannot access reports.", "error", 3000);
          return;
        }

        const reportFileName = locationReportMap[userLocation];
        if (!reportFileName) {
          showSnackbar(
            `No report available for location: ${userLocation}. Please contact your administrator.`,
            "error",
            3000
          );
          return;
        }

        const reportPageUrl = `${webAbsoluteUrl}/SitePages/${reportFileName}`;
        window.open(reportPageUrl, '_blank');
        return;
      }

      // For other roles, show "No Report Available"
      showSnackbar("You do not have access to reports.", "error", 3000);
    } else {
      setActiveSection(sectionId);
    }
  }, [webAbsoluteUrl, employeeData, showSnackbar]);
  const allNavigationItems = useMemo(
    () => [
      { id: "home", label: "Home", icon: Home },
      { id: "onboarding", label: "Onboarding", icon: Users },
      { id: "meetings", label: "Meeting Details", icon: Calendar },
      { id: "questionnaire", label: "Buddy Questionnaire", icon: FileText },
      {
        id: "joineequestionnaire",
        label: "Joinee Questionnaire",
        icon: FileText,
      },
      { id: "holidays", label: "Holidays", icon: Gift },
      { id: "userguides", label: "BuddyPedia", icon: BookOpen },
      { id: "imagemanagement", label: "Image Management", icon: Image },
      { id: "reports", label: "Reports", icon: BarChart3 },
      { id: "helpdesk", label: "Help Desk", icon: Headphones },
    ],
    [],
  );

  const navigationItems = useMemo(() => {
    const r = normalizeRole(employeeData?.Role);
    return allNavigationItems.filter((item) => {
      if (r === "Super Admin") return true;
      if (r === "L&D Site Admin")
        return ["home", "onboarding", "meetings", "userguides", "helpdesk", "reports"].includes(item.id);
      if (r === "Buddy")
        return ["home", "meetings", "userguides", "helpdesk"].includes(item.id);
      if (r === "New Joinee")
        return ["home", "meetings", "userguides", "helpdesk"].includes(item.id);
      return ["home", "userguides", "helpdesk"].includes(item.id);
    });
  }, [employeeRole, allNavigationItems]);

  useEffect(() => {
    const ids = navigationItems.map((i) => i.id);
    if (!ids.includes(activeSection)) setActiveSection("home");
  }, [navigationItems, activeSection]);

  const filteredReviewStories = useMemo(() => {
    const source =
      reviewTab === "Pending"
        ? pendingStories
        : reviewTab === "Approved"
          ? approvedReviewStories
          : rejectedReviewStories;
    const search = reviewSearch.trim().toLowerCase();
    if (!search) return source;
    return source.filter((story) =>
      [
        story.name,
        story.team,
        story.employeeEmail,
        story.quote,
        story.reviewedByName,
        story.reviewComment,
      ]
        .join(" ")
        .toLowerCase()
        .includes(search),
    );
  }, [
    reviewTab,
    reviewSearch,
    pendingStories,
    approvedReviewStories,
    rejectedReviewStories,
  ]);

  const reviewTotalPages = Math.max(
    1,
    Math.ceil(filteredReviewStories.length / REVIEW_PAGE_SIZE),
  );
  const visibleReviewStories = useMemo(
    () =>
      filteredReviewStories.slice(
        reviewPage * REVIEW_PAGE_SIZE,
        reviewPage * REVIEW_PAGE_SIZE + REVIEW_PAGE_SIZE,
      ),
    [filteredReviewStories, reviewPage],
  );

  useEffect(() => {
    setReviewPage(0);
  }, [reviewTab, reviewSearch]);

  useEffect(() => {
    if (reviewPage >= reviewTotalPages) {
      setReviewPage(Math.max(0, reviewTotalPages - 1));
    }
  }, [reviewPage, reviewTotalPages]);

  const myRequestTotalPages = Math.max(
    1,
    Math.ceil(mySubmissions.length / MY_REQUEST_PAGE_SIZE),
  );
  const myPendingCount = useMemo(
    () =>
      mySubmissions.filter((story) => story.approvalStatus === "Pending")
        .length,
    [mySubmissions],
  );
  const visibleMySubmissions = useMemo(
    () =>
      mySubmissions.slice(
        myRequestPage * MY_REQUEST_PAGE_SIZE,
        myRequestPage * MY_REQUEST_PAGE_SIZE + MY_REQUEST_PAGE_SIZE,
      ),
    [mySubmissions, myRequestPage],
  );

  useEffect(() => {
    if (myRequestPage >= myRequestTotalPages) {
      setMyRequestPage(Math.max(0, myRequestTotalPages - 1));
    }
  }, [myRequestPage, myRequestTotalPages]);

  const renderStoryCard = (s: Story, i: number, isInactive = false) => {
    const initials = s.name ? s.name[0].toUpperCase() : "";
    const reactionSummary = s.id
      ? storyReactionSummaries[s.id]
      : undefined;
    const isReactionSaving = Boolean(
      s.id && reactionSavingStoryIds.includes(s.id),
    );
    
    const authorRoleDisplay = s.authorRole ? ` (${s.authorRole})` : "";
    const displayName = s.buddyName 
      ? `${s.name}${authorRoleDisplay} appreciated the support from ${s.buddyName}`
      : `${s.name}${authorRoleDisplay}`;

    return (
      <div
        key={s.id || i}
        className={`${styles.storyCard} ${isInactive ? styles.storyCardInactive : ""}`}
      >
        {s.id && isSuperAdmin && isInactive && (
          <button
            className={styles.restoreBtn + ' ' + styles.storyRestoreBottomBtn}
            onClick={(e) => {
              e.stopPropagation();
              handleActivateStory(s.id!);
            }}
            title="Restore story"
            aria-label="Restore story"
          >
            <RefreshCw size={12} style={{ marginRight: 4 }} /> Restore
          </button>
        )}
        {s.id && isSuperAdmin && !isInactive && (
          <div className={styles.storyAdminBtns}>
            <button
              className={styles.storyDeleteBtn}
              onClick={(e) => {
                e.stopPropagation();
                handleDeactivateStory(s.id!);
              }}
              title="Hide story"
              aria-label="Hide story"
            >
              <EyeOff size={10} />
            </button>
          </div>
        )}
        <div className={styles.storyCardMain}>
          <div
            className={styles.storyAvatar}
            style={{
              background: isInactive
                ? "linear-gradient(135deg, #78736e, #a09b97)"
                : "linear-gradient(135deg, #004632, #00D264)",
            }}
          >
            {initials}
          </div>
          <div className={styles.storyBody}>
            <div className={styles.storyName}>{displayName}</div>
            <div className={styles.storyTeam}>{s.team}</div>
            <div className={styles.storyQuote} title={s.quote}>
              {s.quote}
            </div>
            <div className={styles.storyRating}>
              <HalfStarRating value={s.rating ?? 5} readonly size={16} />
            </div>
          </div>
        </div>
        {!isInactive && s.id && (
          <div
            className={styles.storyReactionBar}
            role="group"
            aria-label={`Reactions for ${s.name}'s success story`}
          >
            {reactionFeedback?.storyId === s.id && (
              <span
                key={reactionFeedback.token}
                className={`${styles.storyReactionFeedback} ${
                  reactionFeedback.mode === "removed"
                    ? styles.storyReactionFeedbackRemoved
                    : styles.storyReactionFeedbackSelected
                }`}
                aria-hidden="true"
              >
                {reactionFeedback.emoji}
              </span>
            )}
            {REACTION_OPTIONS.map((reaction) => {
              const count = reactionSummary?.counts[reaction.type] || 0;
              const selected =
                reactionSummary?.currentUserReaction === reaction.type;
              return (
                <button
                  key={reaction.type}
                  type="button"
                  className={`${styles.storyReactionBtn} ${
                    selected ? styles.storyReactionBtnSelected : ""
                  }`}
                  onClick={(event) => {
                    event.stopPropagation();
                    toggleStoryReaction(s.id!, reaction.type);
                  }}
                  disabled={isReactionSaving || !loggedInUserEmail}
                  aria-pressed={selected}
                  aria-label={`${reaction.label}: ${count}. ${
                    selected
                      ? "Click to remove your reaction"
                      : reactionSummary?.currentUserReaction
                        ? `Click to change your reaction to ${reaction.label}`
                        : `Click to react with ${reaction.label}`
                  }`}
                  title={
                    selected
                      ? `Remove ${reaction.label}`
                      : `React with ${reaction.label}`
                  }
                >
                  <span aria-hidden="true">{reaction.emoji}</span>
                  <strong>{count}</strong>
                </button>
              );
            })}
          </div>
        )}
      </div>
    );
  };

  const formatReviewDate = (value?: string): string => {
    if (!value) return "";
    const date = new Date(value);
    return Number.isNaN(date.getTime())
      ? ""
      : date.toLocaleString(undefined, {
          day: "2-digit",
          month: "short",
          year: "numeric",
          hour: "2-digit",
          minute: "2-digit",
        });
  };

  const renderPendingStoryCard = (story: Story, index: number) => (
    <div
      key={story.id || index}
      className={styles.reviewGridRow}
      role="row"
    >
      <div className={styles.reviewGridCell} role="cell">
        <strong>{story.name}</strong>
        <span>{story.team || "Team"}</span>
        {story.employeeEmail && <small>{story.employeeEmail}</small>}
      </div>
      <div
        className={`${styles.reviewGridCell} ${styles.reviewGridStoryCell}`}
        role="cell"
        title={story.quote}
      >
        <span>{story.quote}</span>
      </div>
      <div className={styles.reviewGridCell} role="cell">
        <HalfStarRating value={story.rating ?? 5} readonly size={14} />
      </div>
      <div className={styles.reviewGridCell} role="cell">
        <span
          className={`${styles.storyStatusBadge} ${styles.storyStatusPending}`}
        >
          <Clock3 size={11} /> Pending
        </span>
      </div>
      <div className={styles.reviewGridCell} role="cell">
        <span className={styles.reviewGridMuted}>Awaiting decision</span>
      </div>
      <div
        className={`${styles.reviewGridCell} ${styles.reviewGridActions}`}
        role="cell"
      >
        <button
          className={styles.storyRejectBtn}
          onClick={() => story.id && handleReviewStory(story.id, "reject-story")}
          disabled={!story.id}
        >
          <X size={13} /> Reject
        </button>
        <button
          className={styles.storyApproveBtn}
          onClick={() => story.id && handleReviewStory(story.id, "approve-story")}
          disabled={!story.id}
        >
          <Check size={13} /> Approve &amp; Publish
        </button>
      </div>
    </div>
  );

  const renderReviewHistoryCard = (story: Story, index: number) => {
    const approved = story.approvalStatus === "Approved";
    return (
      <div
        key={story.id || index}
        className={styles.reviewGridRow}
        role="row"
      >
        <div className={styles.reviewGridCell} role="cell">
          <strong>{story.name}</strong>
          <span>{story.team || "Team"}</span>
          {story.employeeEmail && <small>{story.employeeEmail}</small>}
        </div>
        <div
          className={`${styles.reviewGridCell} ${styles.reviewGridStoryCell}`}
          role="cell"
          title={story.quote}
        >
          <span>{story.quote}</span>
          {!approved && story.reviewComment && (
            <small className={styles.reviewGridReason}>
              Reason: {story.reviewComment}
            </small>
          )}
        </div>
        <div className={styles.reviewGridCell} role="cell">
          <HalfStarRating value={story.rating ?? 5} readonly size={14} />
        </div>
        <div className={styles.reviewGridCell} role="cell">
          <span
            className={`${styles.storyStatusBadge} ${
              approved
                ? styles.storyStatusApproved
                : styles.storyStatusRejected
            }`}
          >
            {approved ? <Check size={11} /> : <X size={11} />}
            {story.approvalStatus}
          </span>
        </div>
        <div className={styles.reviewGridCell} role="cell">
          <span>{story.reviewedByName || "—"}</span>
          {story.reviewedByEmail && <small>{story.reviewedByEmail}</small>}
        </div>
        <div className={styles.reviewGridCell} role="cell">
          <span>{formatReviewDate(story.reviewedOn) || "—"}</span>
        </div>
      </div>
    );
  };

  const renderMySubmissionCard = (story: Story, index: number) => {
    const rejected = story.approvalStatus === "Rejected";
    return (
      <div
        key={story.id || index}
        className={styles.myRequestGridRow}
        role="row"
      >
        <div className={styles.reviewGridCell} role="cell">
          <strong>{story.name}</strong>
          <span>{story.team || "Team"}</span>
        </div>
        <div
          className={`${styles.reviewGridCell} ${styles.reviewGridStoryCell}`}
          role="cell"
          title={story.quote}
        >
          <span>{story.quote}</span>
          {rejected && (
            <small className={styles.reviewGridReason}>
              Reason: {story.reviewComment || "—"}
            </small>
          )}
        </div>
        <div className={styles.reviewGridCell} role="cell">
          <HalfStarRating value={story.rating ?? 5} readonly size={14} />
        </div>
        <div className={styles.reviewGridCell} role="cell">
          <span
            className={`${styles.storyStatusBadge} ${
              rejected
                ? styles.storyStatusRejected
                : styles.storyStatusPending
            }`}
          >
            {rejected ? <X size={11} /> : <Clock3 size={11} />}
            {story.approvalStatus}
          </span>
        </div>
        <div className={styles.reviewGridCell} role="cell">
          {rejected ? (
            <>
              <span>
                {story.reviewedByName
                  ? `Reviewed by ${story.reviewedByName}`
                  : "Reviewed"}
              </span>
              {story.reviewedOn && (
                <small>{formatReviewDate(story.reviewedOn)}</small>
              )}
            </>
          ) : (
            <span className={styles.reviewGridMuted}>Waiting for review</span>
          )}
        </div>
      </div>
    );
  };

  if (!employeeDataLoaded) {
    return (
      <div
        className={styles.portalWrapper}
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
        }}
      >
        <div style={{ textAlign: "center", color: "#6b7280" }}>
          <div style={{ fontSize: "0.9rem", marginBottom: "0.5rem" }}>
            Loading application...
          </div>
        </div>
      </div>
    );
  }

  const userExists = !!employeeData;
  const userIsActive = employeeData ? isLoginEmployeeActive : false;

  if (employeeDataLoaded && (!userExists || !userIsActive)) {
    let deniedMessage = "You don't have access to this application.";
    let deniedReason = "";

    if (!userExists) {
      deniedMessage = "User not found in the system.";
      deniedReason = "Your account does not exist in the Database pls contact administrator.";
    } else if (!userIsActive) {
      deniedMessage = "Your account is inactive.";
      deniedReason = "Inactive accounts cannot access the Buddy System.";
    }

    return (
      <div className={styles.accessDeniedWrapper}>
        <div className={styles.accessDeniedContainer}>
          <div className={styles.accessDeniedIcon}>
            <svg
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect>
              <path d="M7 11V7a5 5 0 0 1 10 0v4"></path>
            </svg>
          </div>
          <h1 className={styles.accessDeniedTitle}>No Access</h1>
          <p className={styles.accessDeniedMessage}>{deniedMessage}</p>
          <p className={styles.accessDeniedSubtext}>
            {deniedReason ||
              "Please contact your administrator if you believe this is incorrect."}
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.portalWrapper}>
      <aside
        className={`${styles.sidebar} ${sidebarOpen ? styles.open : styles.closed}`}
      >
        <div className={styles.sidebarInner}>
          
          <div className={styles.logo}>
            {sidebarOpen ? (
              <div className={styles.logoFull}>
                <div className={styles.logoIconWrap}>
                  <img src={`${NEULAND_LOGO_URL}`} alt="Neuland logo" />
                </div>
                <div>
                  <div className={styles.logoText}>neuland</div>
                  <div className={styles.logoSubtext}>Buddy System</div>
                </div>
              </div>
            ) : (
              <div className={styles.logoIconOnly}>
                <img src={`${NEULAND_LOGO_URL}`} alt="Neuland" />
              </div>
            )}
          </div>
          
          <nav className={styles.nav}>
            {navigationItems.map((item) => {
              const Icon = item.icon;
              const isActive = activeSection === item.id;
              return (
                <button
                  key={item.id}
                  onClick={() => handleNavigation(item.id)}
                  className={`${styles.navItem} ${isActive ? styles.active : ""}`}
                  title={!sidebarOpen ? item.label : undefined}
                >
                  <Icon size={18} />
                  {sidebarOpen && <span>{item.label}</span>}
                  {isActive && sidebarOpen && (
                    <span className={styles.navActiveBar} />
                  )}
                </button>
              );
            })}
          </nav>
          {/* {sidebarOpen && (
            <div className={styles.userBlock}>
              <AvatarImg
                name={employeeData?.FullName || "User"}
                email={employeeData?.EmailID}
                webAbsoluteUrl={webAbsoluteUrl}
                size={40}
              />
              <div className={styles.userInfo}>
                <div className={styles.userName}>
                  {employeeData?.FullName || "User"}
                </div>
                <div className={styles.userRole}>
                  {employeeData?.Role || "User"}
                </div>
              </div>
            </div>
          )} */}
        </div>
        <button
          className={`${styles.toggleBtn} ${sidebarOpen ? styles.toggleBtnOpen : styles.toggleBtnClosed}`}
          onClick={() => setSidebarOpen((p) => !p)}
          title={sidebarOpen ? "Collapse sidebar" : "Expand sidebar"}
          aria-label={sidebarOpen ? "Collapse sidebar" : "Expand sidebar"}
        >
          {sidebarOpen ? <ChevronLeft size={13} /> : <ChevronRight size={13} />}
        </button>
      </aside>

      <main className={styles.main}>
        <div
          className={styles.content}
          style={
            ["onboarding", "meetings", "questionnaire", "joineequestionnaire", "holidays", "userguides", "imagemanagement"].includes(
              activeSection,
            )
              ? {
                  flex: "1",
                  minHeight: 0,
                  overflow: "hidden",
                  padding: 0,
                  display: "flex",
                  flexDirection: "column",
                }
              : { display: "block" }
          }
        >
          {activeSection === "home" ? (
            <>
              <div className={styles.heroBanner}>
                <img
                  src={heroImageUrl || require("../assets/banner_image.jpg")}
                  alt="neuland team"
                  className={styles.heroBannerImg}
                />
                <div className={styles.heroOverlay} />
                <div className={styles.heroSummit} aria-hidden="true">
                  <svg
                    viewBox="0 0 400 300"
                    fill="none"
                    preserveAspectRatio="none"
                  >
                    <path
                      d="M-20 300 L180 0 L220 60 L420 300Z"
                      fill="rgba(0,210,100,0.10)"
                    />
                    <path
                      d="M80 300 L260 0 L300 50 L500 300Z"
                      fill="rgba(0,70,50,0.16)"
                    />
                  </svg>
                </div>
                <div className={styles.heroContent} style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", position: "relative" }}>
                  <div className={styles.heroTextSide}>
                    {/* <div className={styles.heroBadge}><span className={styles.heroBadgeDot}></span>Neuland Buddy System</div> */}
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: "1.2rem",
                        marginBottom: "0.5rem",
                      }}
                    >
                      <div
                        className={styles.buddyLogoPanel}
                        style={{
                          width: "110px",
                          height: "100px",
                          margin: 0,
                          flexShrink: 0,
                        }}
                      >
                        <div
                          className={styles.buddyLogoPanelRing}
                          aria-hidden="true"
                        />
                        <img
                          src={buddyLogoUrl || require("../assets/BuddyLogo-preview.png")}
                          alt="Buddy — The Guiding Star at Neuland"
                          className={styles.buddyLogoHero}
                          onError={(e) => {
                            (e.target as HTMLImageElement).style.display =
                              "none";
                          }}
                        />
                      </div>
                      <h1
                        className={styles.heroTitle}
                        style={{ margin: 0, fontSize: "2rem", lineHeight: 1.1 }}
                      >
                        Welcome to neuland
                        <br />
                        <span className={styles.heroTitleAccent}>
                          Buddy System
                        </span>
                      </h1>
                    </div>
                    <p className={styles.heroSubtitle}>
                      Empowering new joiners through meaningful connections and
                      guided support
                      {/* ksdhsajkhfjk */}
                    </p>
                  </div>
                  <div className={styles.userBlock} style={{ position: "absolute", top: "0px", right: "0px" }}>
                    
                    <div className={styles.userInfo}>
                      <div className={styles.userName}>
                        {employeeData?.FullName || "User"}
                      </div>
                      <div className={styles.userRole}>
                        {employeeData?.Role || "User"}
                      </div>
                    </div>
                    <AvatarImg
                      name={employeeData?.FullName || "User"}
                      email={employeeData?.EmailID || props.userloginDetails?.email}
                      webAbsoluteUrl={webAbsoluteUrl}
                      context={props.context}
                      photoUrl={props.userloginDetails?.photoUrl}
                      size={40}
                    />
                  </div>
                </div>
              </div>

              <div className={styles.topRow}>
                <div className={styles.halfCol}>
                  <h3 className={styles.sectionTitle}>Performance Overview</h3>
                  <div className={styles.kpiGrid}>
                    {kpiCards.map((card, idx) => (
                      <div
                        key={idx}
                        className={`${styles.kpiCard} ${card.full ? styles.kpiCardFull : ""}`}
                      >
                        <div
                          className={styles.kpiIcon}
                          style={{ background: card.bg }}
                        >
                          {card.icon}
                        </div>
                        <div className={styles.kpiBody}>
                          <div className={styles.kpiLabel}>{card.label}</div>
                          <div
                            className={
                              card.value.length > 6
                                ? styles.kpiValueText
                                : styles.kpiValue
                            }
                          >
                            {card.value}
                          </div>
                          <div
                            className={`${styles.kpiTrend} ${styles[card.dir]}`}
                          >
                            <TrendingUp size={11} />
                            <span>{card.trend}</span>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                <div className={`${styles.halfCol} ${styles.halfColGallery}`}>
                  <div className={styles.galleryHeader}>
                    <h3 className={styles.sectionTitle}>Gallery</h3>
                    {isSuperAdmin && (
                      <button
                        className={styles.uploadBtn}
                        onClick={openUploadModal}
                      >
                        <Upload size={13} />
                        <span>Add Media</span>
                      </button>
                    )}
                  </div>

                  <div className={styles.sliderWrapper}>
                    {galleryLoading ? (
                      <div className={styles.galleryEmptyState}>
                        <div className={styles.galleryLoadingState}>
                          <div className={styles.galleryLoadingSpinner} />
                          <span>Loading gallery…</span>
                        </div>
                      </div>
                    ) : activeGalleryItems.length === 0 ? (
                      <div className={styles.galleryEmptyState}>
                        <EmptyState
                          icon={<Image size={22} />}
                          title="No active media"
                          subtitle={
                            isSuperAdmin
                              ? "Upload new images or restore inactive ones from the panel below."
                              : "No gallery images to display yet."
                          }
                        />
                      </div>
                    ) : (
                      <>
                        <div className={styles.slider}>
                          <div
                            className={styles.sliderTrack}
                            onClick={() =>
                              activeGalleryItems[slideIdx] &&
                              openLightbox(
                                activeGalleryItems[slideIdx].img,
                                activeGalleryItems[slideIdx].label,
                                activeGalleryItems[slideIdx].fileType || "image",
                              )
                            }
                          >
                            {activeGalleryItems.map((item, idx) => (
                              <div
                                key={item.itemId || idx}
                                className={`${styles.sliderSlide} ${idx === slideIdx ? styles.sliderSlideActive : ""}`}
                              >
                                {item.fileType === "video" ? (
                                  <video
                                    src={item.img}
                                    controls
                                    crossOrigin="anonymous"
                                    preload="metadata"
                                    className={styles.sliderVideo}
                                    onPlay={() => setSliderPaused(true)}
                                    onPause={e => {
                                      setTimeout(() => {
                                        const videos = document.querySelectorAll('video');
                                        const anyPlaying = Array.from(videos).some(v => !v.paused && !v.ended);
                                        if (!anyPlaying) setSliderPaused(false);
                                      }, 100);
                                    }}
                                    onEnded={e => {
                                      setTimeout(() => {
                                        const videos = document.querySelectorAll('video');
                                        const anyPlaying = Array.from(videos).some(v => !v.paused && !v.ended);
                                        if (!anyPlaying) setSliderPaused(false);
                                      }, 100);
                                    }}
                                    onMouseEnter={() => setSliderPaused(true)}
                                    onMouseLeave={() => setSliderPaused(false)}
                                  />
                                ) : (
                                  <img
                                    src={item.img}
                                    alt={item.label}
                                    className={styles.sliderImg}
                                    onMouseEnter={() => setSliderPaused(true)}
                                    onMouseLeave={() => setSliderPaused(false)}
                                  />
                                )}
                                <div
                                  className={styles.sliderLabel}
                                  title={item.label}
                                >
                                  {item.label}
                                </div>
                                <div className={styles.sliderZoom}>
                                  <ZoomIn size={16} />
                                </div>
                              </div>
                            ))}
                          </div>
                          {activeGalleryItems.length > 1 && (
                            <>
                              <button
                                className={`${styles.sliderArrow} ${styles.sliderPrev}`}
                                onClick={prevSlide}
                                aria-label="Previous image"
                              >
                                <ChevronLeft size={20} />
                              </button>
                              <button
                                className={`${styles.sliderArrow} ${styles.sliderNext}`}
                                onClick={nextSlide}
                                aria-label="Next image"
                              >
                                <ChevronRight size={20} />
                              </button>
                            </>
                          )}
                          <div className={styles.sliderDots}>
                            {activeGalleryItems.map((_, idx) => (
                              <button
                                key={idx}
                                className={`${styles.sliderDot} ${idx === slideIdx ? styles.sliderDotActive : ""}`}
                                onClick={(e) => {
                                  e.stopPropagation();
                                  goToSlide(idx);
                                }}
                                aria-label={`Go to slide ${idx + 1}`}
                              />
                            ))}
                          </div>
                        </div>

                        <div className={styles.sliderThumbs}>
                          {activeGalleryItems.map((item, idx) => (
                            <div
                              key={item.itemId || idx}
                              className={styles.sliderThumbWrapper}
                            >
                              <button
                                className={`${styles.sliderThumb} ${idx === slideIdx ? styles.sliderThumbActive : ""}`}
                                onClick={() => goToSlide(idx)}
                                title={item.label}
                              >
                                {item.fileType === "video" ? (
                                  <div
                                    style={{
                                      position: "relative",
                                      width: "100%",
                                      height: "100%",
                                      backgroundColor: "#000",
                                      display: "flex",
                                      alignItems: "center",
                                      justifyContent: "center",
                                      overflow: "hidden",
                                    }}
                                  >
                                    <video
                                      src={item.img}
                                      crossOrigin="anonymous"
                                      preload="metadata"
                                      className={styles.sliderVideo}
                                    />
                                    <div
                                      style={{
                                        position: "absolute",
                                        color: "#fff",
                                        fontSize: "10px",
                                      }}
                                    >
                                      Play
                                    </div>
                                  </div>
                                ) : (
                                  <img src={item.img} alt={item.label} />
                                )}
                              </button>
                              {isSuperAdmin && (
                                <div className={styles.sliderThumbMenu}>
                                  <button
                                    className={styles.sliderThumbMenuBtn}
                                    onClick={() => handleDeactivateGallery(idx)}
                                    title="Hide image (set Inactive)"
                                    aria-label="Hide image"
                                  >
                                    <EyeOff size={11} />
                                  </button>
                                </div>
                              )}
                            </div>
                          ))}
                        </div>
                      </>
                    )}
                  </div>
                </div>
              </div>

              <div className={styles.buddyPerformanceHome}>
                <BuddyPerformance
                  context={props.context}
                  isSuperAdmin={isSuperAdmin}
                  userLocation={employeeData?.Location || ""}
                  currentUserEmail={
                    employeeData?.EmailID || props.userloginDetails?.email || ""
                  }
                  currentUserGlobalId={employeeData?.GlobalID || ""}
                  currentUserName={employeeData?.FullName || ""}
                  onCurrentUserRankingChange={
                    handleCurrentUserRankingChange
                  }
                />
              </div>

              {isSuperAdmin && (
                <div className={styles.adminPanel}>
                  <button
                    className={styles.adminPanelToggle}
                    onClick={() => setShowInactiveGallery((p) => !p)}
                  >
                    <span className={styles.adminPanelToggleLeft}>
                      <EyeOff size={14} />
                      <span>Inactive Gallery Content</span>
                      {inactiveGalleryItems.length > 0 && (
                        <span className={styles.adminBadge}>
                          {inactiveGalleryItems.length}
                        </span>
                      )}
                    </span>
                    <ChevronRight
                      size={14}
                      className={showInactiveGallery ? styles.chevronDown : ""}
                    />
                  </button>
                  {showInactiveGallery && (
                    <div className={styles.adminPanelBody}>
                      {inactiveGalleryItems.length === 0 ? (
                        <div className={styles.adminEmptyNote}>
                          <Eye size={14} /> No inactive images — all gallery
                          items are currently visible.
                        </div>
                      ) : (
                        <div className={styles.inactiveGalleryGrid}>
                          {inactiveGalleryItems.map((item, idx) => (
                            <div
                              key={item.itemId || idx}
                              className={styles.inactiveGalleryCard}
                            >
                              <div className={styles.inactiveGalleryImgWrap}>
                                {item.fileType === "video" ? (
                                  <video
                                    src={item.img}
                                    crossOrigin="anonymous"
                                    preload="metadata"
                                    className={styles.inactiveGalleryVideo}
                                  />
                                ) : (
                                  <img
                                    src={item.img}
                                    alt={item.label}
                                    className={styles.inactiveGalleryImg}
                                  />
                                )}
                                <div className={styles.inactiveGalleryOverlay}>
                                  <EyeOff size={16} />
                                </div>
                              </div>
                              <div className={styles.inactiveGalleryMeta}>
                                <span
                                  className={styles.inactiveGalleryLabel}
                                  title={item.label}
                                >
                                  {item.label}
                                </span>
                                <button
                                  className={styles.restoreBtn}
                                  onClick={() => handleActivateGallery(idx)}
                                  title="Restore to gallery"
                                >
                                  {/* <Eye size={11} /><span>Restore</span> */}
                                  <RefreshCw size={10} />
                                  <span>Restore</span>
                                </button>
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}

              <div className={styles.storiesSection}>
                <div className={styles.storiesHeader}>
                  <h3 className={styles.sectionTitle}>Success Stories</h3>
                  <div className={styles.storyHeaderActions}>
                    {isSuccessStoryApprover && (
                      <button
                        className={`${styles.reviewNotificationBtn} ${
                          showReviewCenter
                            ? styles.reviewNotificationBtnActive
                            : ""
                        }`}
                        onClick={toggleReviewCenter}
                        aria-expanded={showReviewCenter}
                        aria-label={`Open success story review center. ${pendingStories.length} pending requests.`}
                        title="Open success story review center"
                      >
                        <Bell size={14} />
                        <span>Reviews</span>
                        {pendingStories.length > 0 && (
                          <span className={styles.reviewNotificationBadge}>
                            {pendingStories.length > 99
                              ? "99+"
                              : pendingStories.length}
                          </span>
                        )}
                      </button>
                    )}
                    <button
                      className={`${styles.reviewNotificationBtn} ${
                        showMyRequests
                          ? styles.reviewNotificationBtnActive
                          : ""
                      }`}
                      onClick={toggleMyRequests}
                      aria-expanded={showMyRequests}
                      aria-label={`Open my success story requests. ${myPendingCount} pending requests.`}
                      title="Open my success story requests"
                    >
                      <FileText size={14} />
                      <span>My Requests</span>
                      {myPendingCount > 0 && (
                        <span
                          className={`${styles.reviewNotificationBadge} ${styles.myRequestNotificationBadge}`}
                        >
                          {myPendingCount > 99
                            ? "99+"
                            : myPendingCount}
                        </span>
                      )}
                    </button>
                    <button
                      className={styles.addStoryBtn}
                      onClick={() => {
                        const nextOpen = !showStoryForm;
                        setShowStoryForm(nextOpen);
                        if (nextOpen) {
                          setShowReviewCenter(false);
                          setShowMyRequests(false);
                        }
                      }}
                    >
                      {showStoryForm ? <X size={14} /> : <Plus size={14} />}
                      <span>{showStoryForm ? "Cancel" : "Add Story"}</span>
                    </button>
                  </div>
                </div>

                {showStoryForm && (
                  <div className={styles.storyForm}>
                    <div
                      style={{
                        display: "grid",
                        gridTemplateColumns: "1fr 1fr 1fr 1fr",
                        gap: "12px",
                        marginBottom: "12px",
                      }}
                    >
                      <div className={styles.storyFormField}>
                        <label className={styles.storyFormLabel}>
                          Name <span className={styles.asterisk}>*</span>
                        </label>
                        <input
                          type="text"
                          className={styles.storyFormInput}
                          placeholder="e.g. Priya & Ravi"
                          value={newStory.name}
                          disabled
                          onChange={(e) =>
                            setNewStory((p) => ({ ...p, name: e.target.value }))
                          }
                        />
                      </div>
                      <div className={styles.storyFormField}>
                        <label className={styles.storyFormLabel}>Team</label>
                        <input
                          type="text"
                          className={styles.storyFormInput}
                          placeholder="e.g. Software Development"
                          value={newStory.team}
                          disabled
                          onChange={(e) =>
                            setNewStory((p) => ({ ...p, team: e.target.value }))
                          }
                        />
                      </div>
                      <div className={styles.storyFormField}>
                        <label className={styles.storyFormLabel}>
                          {/* Buddy Name (Optional)  */}
                          Appreciate To
                        </label>
                        <div
                          style={{ position: "relative" }}
                          ref={buddySearchRef}
                        >
                          <input
                            type="text"
                            placeholder="Search the user..."
                            value={newStory.buddyName}
                            onChange={(e) => {
                              const v = e.target.value;
                              setNewStory((p) => ({
                                ...p,
                                buddyName: v,
                                EmailID: "",
                                buddyGlobalID: "",
                              }));
                              if (v.length >= 2) {
                                handleBuddySearch(v);
                                setShowBuddySearch(true);
                              } else {
                                setShowBuddySearch(false);
                                setBuddySearchResults([]);
                              }
                            }}
                            style={{
                              width: "100%",
                              padding: "10px",
                              border: "1px solid #ddd",
                              borderRadius: "4px",
                              fontSize: "14px",
                              fontFamily: "inherit",
                              boxSizing: "border-box",
                            }}
                          />
                          {showBuddySearch && buddySearchResults.length > 0 && (
                            <div
                              style={{
                                position: "absolute",
                                top: "100%",
                                left: 0,
                                right: 0,
                                backgroundColor: "white",
                                border: "1px solid #ddd",
                                borderTop: "none",
                                maxHeight: "150px",
                                overflowY: "auto",
                                zIndex: 1000,
                              }}
                            >
                              {buddySearchResults.map((buddy, idx) => (
                                <div
                                  key={idx}
                                  onClick={() => handleBuddySelect(buddy)}
                                  style={{
                                    padding: "10px",
                                    cursor: "pointer",
                                    borderBottom: "1px solid #eee",
                                    transition: "background-color 0.2s",
                                  }}
                                  onMouseOver={(e) =>
                                    (e.currentTarget.style.backgroundColor =
                                      "#f0f0f0")
                                  }
                                  onMouseOut={(e) =>
                                    (e.currentTarget.style.backgroundColor =
                                      "white")
                                  }
                                >
                                  <div
                                    style={{ fontWeight: "600" }}
                                    title={`${buddy.FullName} - ${buddy.EmailID}`}
                                  >
                                    {buddy.FullName}
                                  </div>
                                  <div
                                    style={{
                                      fontSize: "0.85rem",
                                      color: "#666",
                                    }}
                                  >
                                    {buddy.Department}
                                  </div>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      </div>
                      <div className={styles.storyFormField}>
                        <label className={styles.storyFormLabel}>
                          Rating <span className={styles.asterisk}>*</span>
                        </label>
                        <div
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: "6px",
                            paddingTop: "4px",
                          }}
                        >
                          <HalfStarRating
                            value={storyRating}
                            onChange={setStoryRating}
                            size={24}
                          />
                          <span
                            style={{
                              fontWeight: "600",
                              fontSize: "12px",
                              color: "#374151",
                              minWidth: "36px",
                            }}
                          >
                            {storyRating}/5
                          </span>
                        </div>
                      </div>
                    </div>
                    <div className={styles.storyFormField}>
                      <label className={styles.storyFormLabel}>
                        Quote <span className={styles.asterisk}>*</span>
                      </label>
                      <textarea
                        className={styles.storyFormTextarea}
                        rows={3}
                        placeholder='"Type the success story here..."'
                        value={newStory.quote}
                        onChange={(e) =>
                          setNewStory((p) => ({ ...p, quote: e.target.value }))
                        }
                      />
                    </div>
                    <input
                      type="hidden"
                      value={newStory.EmailID}
                      readOnly
                      style={{ display: "none" }}
                    />
                    <div className={styles.storyFormActions}>
                      <button
                        className={styles.storyFormCancel}
                        onClick={() => setShowStoryForm(false)}
                      >
                        Cancel
                      </button>
                      <button
                        className={styles.storyFormSave}
                        onClick={handleAddStory}
                        disabled={
                          !newStory.name.trim() || !newStory.quote.trim()
                        }
                      >
                        <Star size={13} /> Add Story
                      </button>
                    </div>
                  </div>
                )}

                {isSuccessStoryApprover && showReviewCenter && (
                  <div className={styles.reviewCenter}>
                    <div
                      className={`${styles.reviewCenterHeader} ${styles.workflowThemedHeader}`}
                    >
                      <div className={styles.workflowHeaderIdentity}>
                        <span className={styles.workflowHeaderIcon}>
                          <Bell size={20} />
                        </span>
                        <div>
                          <h4>Success Story Review Center</h4>
                          <p>
                            Review pending requests and track completed
                            decisions.
                          </p>
                        </div>
                      </div>
                      <button
                        className={styles.reviewCenterClose}
                        onClick={toggleReviewCenter}
                        title="Close review center"
                        aria-label="Close review center"
                      >
                        <X size={15} />
                      </button>
                    </div>

                    <div className={styles.reviewCenterToolbar}>
                      <div className={styles.reviewTabs}>
                        {(
                          ["Pending", "Approved", "Rejected"] as const
                        ).map((tab) => {
                          const count =
                            tab === "Pending"
                              ? pendingStories.length
                              : tab === "Approved"
                                ? approvedReviewStories.length
                                : rejectedReviewStories.length;
                          return (
                            <button
                              key={tab}
                              className={`${styles.reviewTab} ${
                                reviewTab === tab
                                  ? styles.reviewTabActive
                                  : ""
                              }`}
                              onClick={() => setReviewTab(tab)}
                            >
                              {tab}
                              <span>{count}</span>
                            </button>
                          );
                        })}
                      </div>
                      <input
                        className={styles.reviewSearchInput}
                        type="search"
                        value={reviewSearch}
                        onChange={(e) => setReviewSearch(e.target.value)}
                        placeholder="Search name, team, email or story..."
                        aria-label="Search review stories"
                      />
                    </div>

                    {visibleReviewStories.length === 0 ? (
                      <div className={styles.storyWorkflowEmpty}>
                        {reviewSearch.trim() ? (
                          <>No matching {reviewTab.toLowerCase()} stories.</>
                        ) : reviewTab === "Pending" ? (
                          <>
                            <Check size={14} /> No stories waiting for approval.
                          </>
                        ) : (
                          <>No {reviewTab.toLowerCase()} review history yet.</>
                        )}
                      </div>
                    ) : (
                      <div className={styles.reviewGridScroller}>
                        <div
                          className={`${styles.reviewGridTable} ${
                            reviewTab === "Pending"
                              ? styles.reviewGridTablePending
                              : ""
                          }`}
                          role="table"
                          aria-label={`${reviewTab} success story requests`}
                        >
                          <div
                            className={styles.reviewGridHeaderRow}
                            role="row"
                          >
                            <div role="columnheader">Employee</div>
                            <div role="columnheader">Story</div>
                            <div role="columnheader">Rating</div>
                            <div role="columnheader">Status</div>
                            <div role="columnheader">Reviewer</div>
                            <div role="columnheader">
                              {reviewTab === "Pending"
                                ? "Actions"
                                : "Reviewed On"}
                            </div>
                          </div>
                          {visibleReviewStories.map((story, index) =>
                            reviewTab === "Pending"
                              ? renderPendingStoryCard(story, index)
                              : renderReviewHistoryCard(story, index),
                          )}
                        </div>
                      </div>
                    )}

                    {filteredReviewStories.length > REVIEW_PAGE_SIZE && (
                      <div className={styles.reviewPagination}>
                        <button
                          onClick={() =>
                            setReviewPage((page) => Math.max(0, page - 1))
                          }
                          disabled={reviewPage === 0}
                          aria-label="Previous review page"
                        >
                          <ChevronLeft size={13} /> Previous
                        </button>
                        <span>
                          Page {reviewPage + 1} of {reviewTotalPages}
                        </span>
                        <button
                          onClick={() =>
                            setReviewPage((page) =>
                              Math.min(reviewTotalPages - 1, page + 1),
                            )
                          }
                          disabled={reviewPage >= reviewTotalPages - 1}
                          aria-label="Next review page"
                        >
                          Next <ChevronRight size={13} />
                        </button>
                      </div>
                    )}
                  </div>
                )}

                {showMyRequests && (
                  <div
                    className={`${styles.storyWorkflowSection} ${styles.myRequestsPanel}`}
                  >
                    <div
                      className={`${styles.storyWorkflowHeader} ${styles.workflowThemedHeader}`}
                    >
                      <div className={styles.workflowHeaderIdentity}>
                        <span className={styles.workflowHeaderIcon}>
                          <FileText size={20} />
                        </span>
                        <div>
                          <h4>My Submissions</h4>
                          <p>Track your pending and rejected stories.</p>
                        </div>
                      </div>
                      <button
                        className={styles.reviewCenterClose}
                        onClick={toggleMyRequests}
                        title="Close my submissions"
                        aria-label="Close my submissions"
                      >
                        <X size={15} />
                      </button>
                    </div>
                    {visibleMySubmissions.length === 0 ? (
                      <div className={styles.storyWorkflowEmpty}>
                        <FileText size={14} /> You have no pending or rejected
                        submissions.
                      </div>
                    ) : (
                      <div className={styles.reviewGridScroller}>
                        <div
                          className={styles.myRequestGridTable}
                          role="table"
                          aria-label="My success story requests"
                        >
                          <div
                            className={styles.myRequestGridHeaderRow}
                            role="row"
                          >
                            <div role="columnheader">Employee</div>
                            <div role="columnheader">Story</div>
                            <div role="columnheader">Rating</div>
                            <div role="columnheader">Status</div>
                            <div role="columnheader">Review Details</div>
                          </div>
                          {visibleMySubmissions.map(renderMySubmissionCard)}
                        </div>
                      </div>
                    )}
                    {mySubmissions.length > MY_REQUEST_PAGE_SIZE && (
                      <div className={styles.reviewPagination}>
                        <button
                          onClick={() =>
                            setMyRequestPage((page) => Math.max(0, page - 1))
                          }
                          disabled={myRequestPage === 0}
                          aria-label="Previous submissions page"
                        >
                          <ChevronLeft size={13} /> Previous
                        </button>
                        <span>
                          Page {myRequestPage + 1} of {myRequestTotalPages}
                        </span>
                        <button
                          onClick={() =>
                            setMyRequestPage((page) =>
                              Math.min(myRequestTotalPages - 1, page + 1),
                            )
                          }
                          disabled={myRequestPage >= myRequestTotalPages - 1}
                          aria-label="Next submissions page"
                        >
                          Next <ChevronRight size={13} />
                        </button>
                      </div>
                    )}
                  </div>
                )}

                {storiesLoading ? (
                  <div className={styles.storiesLoadingState}>
                    <div className={styles.galleryLoadingSpinner} />
                    <span>Loading stories…</span>
                  </div>
                ) : stories.length === 0 ? (
                  <div className={styles.storiesEmptyWrapper}>
                    <EmptyState
                      icon={<MessageSquarePlus size={22} />}
                      title="No success stories yet"
                      subtitle="Be the first to share your buddy experience. Your story can inspire others!"
                    />
                  </div>
                ) : (
                  <div className={styles.storiesScroll}>
                    {stories.map((s, i) => renderStoryCard(s, i, false))}
                  </div>
                )}

                {isSuperAdmin && (
                  <div
                    className={`${styles.adminPanel} ${styles.adminPanelStories}`}
                  >
                    <button
                      className={styles.adminPanelToggle}
                      onClick={() => setShowInactiveStories((p) => !p)}
                    >
                      <span className={styles.adminPanelToggleLeft}>
                        <EyeOff size={14} />
                        <span>Inactive Success Stories</span>
                        {inactiveStories.length > 0 && (
                          <span className={styles.adminBadge}>
                            {inactiveStories.length}
                          </span>
                        )}
                      </span>
                      <ChevronRight
                        size={14}
                        className={
                          showInactiveStories ? styles.chevronDown : ""
                        }
                      />
                    </button>
                    {showInactiveStories && (
                      <div className={styles.adminPanelBody}>
                        {inactiveStories.length === 0 ? (
                          <div className={styles.adminEmptyNote}>
                            <Eye size={14} /> No inactive stories — all stories
                            are currently visible.
                          </div>
                        ) : (
                          <div className={styles.storiesScroll}>
                            {inactiveStories.map((s, i) =>
                              renderStoryCard(s, i, true),
                            )}
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                )}
              </div>
            </>
          ) : activeSection === "onboarding" ? (
            <Onboarding {...props} />
          ) : activeSection === "meetings" ? (
            <MeetingDetails {...props} />
          ) : activeSection === "questionnaire" ? (
            <BuddyQuestionnaire {...props} />
          ) : activeSection === "joineequestionnaire" ? (
            <JoineeQuestionnaire {...props} />
          ) : activeSection === "helpdesk" ? (
            <HelpDesk {...props} />
          ) : activeSection === "holidays" ? (
            <Holidays {...props} />
          ) : activeSection === "userguides" ? (
            <UserGuides
              context={props.context}
              isSuperAdmin={isSuperAdmin}
            />
          ) : activeSection === "imagemanagement" ? (
            // Image Management is restricted to Super Admin only
            isSuperAdmin ? (
              <BannerLogoUpload
                context={props.context}
                showSnackbar={showSnackbar}
                onUploadSuccess={() => {
                  loadHeroImage();
                  loadBuddyLogo();
                }}
              />
            ) : (
              <div className={styles.cardsGrid} style={{ padding: "2rem", textAlign: "center" }}>
                <div style={{ color: "#6b7280", fontSize: "0.95rem" }}>You don't have access to this section.</div>
              </div>
            )
          ) : (
            <div className={styles.cardsGrid}>
              {navigationItems.slice(1).map((item) => {
                const Icon = item.icon;
                const isActive = activeSection === item.id;
                return (
                  <div
                    key={item.id}
                    onClick={() => handleNavigation(item.id)}
                    className={`${styles.card} ${isActive ? styles.cardActive : ""}`}
                  >
                    <div className={styles.cardIcon}>
                      <Icon size={24} />
                    </div>
                    <h3 className={styles.cardTitle}>{item.label}</h3>
                    <p className={styles.cardDesc}>
                      {item.id === "onboarding" &&
                        "Set up new buddies and manage their onboarding journey"}
                      {item.id === "meetings" &&
                        "Schedule and track buddy meeting sessions"}
                      {item.id === "questionnaire" &&
                        "Create and manage buddy assessment forms"}
                      {item.id === "joineequestionnaire" &&
                        "Create and manage joinee onboarding questions"}
                      {item.id === "holidays" &&
                        "Manage organisation holidays and seasonal celebrations"}
                      {item.id === "userguides" &&
                        "Browse trusted guides, templates and learning resources"}
                      {item.id === "reports" &&
                        "View analytics and insights on buddy relationships"}
                    </p>
                    <button className={styles.cardBtn}>
                      {isActive ? "Active" : "Open"} →
                    </button>
                    {isActive && <div className={styles.cardIndicator} />}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </main>

      {uploadModal && (
        <div className={styles.modalBackdrop} onClick={closeUploadModal}>
          <div
            className={styles.uploadModalBox}
            onClick={(e) => e.stopPropagation()}
          >
            <div className={styles.uploadModalHeader}>
              <div className={styles.uploadModalHeaderLeft}>
                <ImagePlus size={18} />
                <span>Upload Gallery Content</span>
              </div>
              <button
                className={styles.uploadModalClose}
                onClick={closeUploadModal}
                aria-label="Close modal"
              >
                <X size={16} />
              </button>
            </div>
            <div className={styles.uploadModalBody}>
              <div
                ref={dropRef}
                className={`${styles.dropZone} ${isDragging ? styles.dropZoneActive : ""} ${uploadPreview ? styles.dropZoneHasFile : ""}`}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                onClick={() => !uploadPreview && fileInputRef.current?.click()}
              >
                {uploadPreview ? (
                  <div className={styles.dropZonePreview}>
                    {uploadFile?.type === "video/mp4" ? (
                      <video
                        src={uploadPreview}
                        controls
                        crossOrigin="anonymous"
                        preload="metadata"
                        className={styles.dropZonePreviewVideo}
                      />
                    ) : (
                      <img
                        src={uploadPreview}
                        alt="Preview"
                        className={styles.dropZonePreviewImg}
                      />
                    )}
                    <button
                      className={styles.dropZoneRemove}
                      onClick={(e) => {
                        e.stopPropagation();
                        setUploadFile(null);
                        setUploadPreview("");
                        setUploadTitle("");
                      }}
                    >
                      <X size={13} /> Remove
                    </button>
                  </div>
                ) : (
                  <div className={styles.dropZoneEmpty}>
                    <div className={styles.dropZoneIcon}>
                      <Upload size={28} />
                    </div>
                    <p className={styles.dropZoneTitle}>
                      {isDragging
                        ? "Drop your image or video here"
                        : "Drag & drop an image or video here"}
                    </p>
                    <p className={styles.dropZoneSub}>
                      or click to browse from your computer
                    </p>
                    <p className={styles.dropZoneFormats}>
                      Supports: JPG, JPEG, PNG, GIF, WEBP (Max 10 MB) · MP4 Videos (Max 25 MB)
                    </p>
                  </div>
                )}
              </div>
              <input
                ref={fileInputRef}
                type="file"
                accept="image/*,.mp4,video/mp4"
                className={styles.hiddenInput}
                onChange={handleFileInput}
              />
              <div className={styles.uploadTitleField}>
                <label className={styles.uploadTitleLabel}>
                  Title{" "}
                  <span className={styles.uploadTitleHint}>
                    (shown in gallery)
                  </span>
                </label>
                <input
                  type="text"
                  className={styles.uploadTitleInput}
                  placeholder="e.g. Team Outing 2024"
                  value={uploadTitle}
                  onChange={(e) => setUploadTitle(e.target.value)}
                  maxLength={60}
                />
                {uploadTitle && (
                  <span className={styles.uploadTitleCount}>
                    {uploadTitle.length}/60
                  </span>
                )}
              </div>
            </div>
            <div className={styles.uploadModalFooter}>
              <button
                className={styles.uploadModalCancel}
                onClick={closeUploadModal}
                disabled={isUploading}
              >
                Cancel
              </button>
              <button
                className={styles.uploadModalConfirm}
                onClick={handleUploadConfirm}
                disabled={!uploadFile || !uploadTitle.trim() || isUploading}
              >
                <Upload size={14} />
                {isUploading ? "Uploading…" : "Add to Gallery"}
              </button>
            </div>
          </div>
        </div>
      )}

      {confirmDialog.open && (
        <div className={styles.modalBackdrop} onClick={cancelAction}>
          <div
            className={styles.confirmDialogBox}
            onClick={(e) => e.stopPropagation()}
          >
            <div className={styles.confirmDialogHeader}>
              <span>
                {confirmDialog.action === "deactivate" && "Hide Image"}
                {confirmDialog.action === "activate" && "Restore Image"}
                {confirmDialog.action === "deactivate-story" && "Hide Story"}
                {confirmDialog.action === "activate-story" && "Restore Story"}
                {confirmDialog.action === "approve-story" && "Approve Story"}
                {confirmDialog.action === "reject-story" && "Reject Story"}
              </span>
            </div>
            <div className={styles.confirmDialogBody}>
              {confirmDialog.action === "deactivate" && (
                <p>
                  This content will be set to <strong>Inactive</strong> and hidden
                  from the gallery. You can restore it anytime from the Inactive
                  content panel below.
                </p>
              )}
              {confirmDialog.action === "activate" && (
                <p>
                  This content will be set back to <strong>Active</strong> and
                  will appear in the gallery again.
                </p>
              )}
              {confirmDialog.action === "deactivate-story" && (
                <p>
                  This story will be set to <strong>Inactive</strong> and hidden
                  from all users. You can restore it anytime from the Inactive
                  Stories panel below.
                </p>
              )}
              {confirmDialog.action === "activate-story" && (
                <p>
                  This story will be set back to <strong>Active</strong> and
                  visible to all users again.
                </p>
              )}
              {confirmDialog.action === "approve-story" && (
                <p>
                  Approving will immediately <strong>publish</strong> this story.
                </p>
              )}
              {confirmDialog.action === "reject-story" && (
                <>
                  <label className={styles.reviewReasonLabel}>
                    Rejection reason <span>*</span>
                  </label>
                  <textarea
                    className={styles.reviewReasonInput}
                    rows={4}
                    value={reviewComment}
                    maxLength={1000}
                    onChange={(e) => setReviewComment(e.target.value)}
                    placeholder="Explain why this story is being rejected..."
                    autoFocus
                  />
                  <div className={styles.reviewReasonCount}>
                    {reviewComment.length}/1000
                  </div>
                </>
              )}
            </div>
            <div className={styles.confirmDialogFooter}>
              <button
                className={styles.confirmDialogCancel}
                onClick={cancelAction}
              >
                Cancel
              </button>
              <button
                className={`${styles.confirmDialogConfirm} ${confirmDialog.action === "deactivate" || confirmDialog.action === "deactivate-story" || confirmDialog.action === "reject-story" ? styles.confirmDialogDanger : styles.confirmDialogSuccess}`}
                onClick={confirmAction}
                disabled={
                  confirmDialog.action === "reject-story" &&
                  !reviewComment.trim()
                }
              >
                {(confirmDialog.action === "deactivate" ||
                  confirmDialog.action === "deactivate-story") && (
                  <>
                    <EyeOff size={13} /> Set Inactive
                  </>
                )}
                {(confirmDialog.action === "activate" ||
                  confirmDialog.action === "activate-story") && (
                  <>
                    <Eye size={13} /> Restore
                  </>
                )}
                {confirmDialog.action === "approve-story" && (
                  <>
                    <Check size={13} /> Approve &amp; Publish
                  </>
                )}
                {confirmDialog.action === "reject-story" && (
                  <>
                    <X size={13} /> Reject Story
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {lightbox.open && (
        <div className={styles.lightboxBackdrop} onClick={closeLightbox}>
          <div
            className={styles.lightboxInner}
            onClick={(e) => e.stopPropagation()}
          >
            <button
              className={styles.lightboxClose}
              onClick={closeLightbox}
              aria-label="Close lightbox"
            >
              <X size={16} />
            </button>
            {lightbox.fileType === "video" ? (
              <video
                src={lightbox.img}
                controls
                crossOrigin="anonymous"
                preload="metadata"
                className={styles.lightboxVideo}
              />
            ) : (
              <img
                src={lightbox.img}
                alt={lightbox.label}
                className={styles.lightboxImg}
              />
            )}
            <div className={styles.lightboxCaption}>{lightbox.label}</div>
          </div>
        </div>
      )}

      {showWelcome &&
        employeeData &&
        (() => {
          const h = new Date().getHours();
          const gr =
            h < 12
              ? "Good morning"
              : h < 17
                ? "Good afternoon"
                : "Good evening";
          const first = (employeeData.FullName || "").split(" ")[0] || "there";

          const content: Record<
            string,
            {
              badge: string;
              nameLine1: string;
              nameLine2: string;
              msg: string;
              btn: string;
            }
          > = {
            "Super Admin": {
              badge: "Super Admin",
              nameLine1: gr + ",",
              nameLine2: first + "!",
              msg: "You have full access to the Neuland Buddy System. Manage all new joiners, buddies, meetings, questionnaires, and holidays across all locations.",
              btn: "Go to Dashboard",
            },
            "L&D Site Admin": {
              badge: "L&D Site Admin",
              nameLine1: gr + ",",
              nameLine2: "Welcome back!",
              msg: "Your L&D dashboard is ready. Manage onboarding, meetings, and questionnaires for new joiners within your assigned location.",
              btn: "View Onboarding",
            },
            Buddy: {
              badge: "Buddy",
              nameLine1: gr + ",",
              nameLine2: first + "!",
              msg: "Your new joiners are counting on you today. Head to Meeting Details to track sessions and support your NJs on their journey.",
              btn: "View My NJs",
            },
            "New Joinee": {
              badge: "New Joinee",
              nameLine1: "Welcome to",
              nameLine2: "neuland!",
              msg: "We're so glad you're here! Your buddy is ready to guide you through onboarding. Check your meeting schedule and begin your journey.",
              btn: "Begin Journey",
            },
            Employee: {
              badge: "User",
              nameLine1: gr + ",",
              nameLine2: first + "!",
              msg: "Welcome to the Neuland Buddy System. Explore available resources and updates from your dashboard.",
              btn: "Go to Home",
            },
          };

          const welcomeRole = normalizeRole(employeeData.Role);
          const c = content[welcomeRole] ?? content.Employee;

          return (
            <div
              onClick={() => dismissWelcome()}
              style={{
                position: "fixed",
                inset: 0,
                background: "rgba(0,0,0,0.6)",
                zIndex: 9999,
                display: "flex",
                alignItems: "flex-start",
                justifyContent: "center",
                padding: "1.5rem",
                fontFamily: '"Inter",-apple-system,sans-serif',
                overflowY: "auto",
              }}
            >
              <style>{`
        @keyframes spFloat  { 0%,100%{transform:translateY(0)} 50%{transform:translateY(-14px)} }
        @keyframes spPopIn  { 0%{opacity:0;transform:scale(.72) translateY(32px)} 70%{transform:scale(1.03) translateY(-4px)} 100%{opacity:1;transform:scale(1) translateY(0)} }
        @keyframes spSlide  { from{opacity:0;transform:translateY(16px)} to{opacity:1;transform:translateY(0)} }
        @keyframes spSpin   { from{transform:rotate(0deg)} to{transform:rotate(360deg)} }
        @keyframes spRevSpin{ from{transform:rotate(0deg)} to{transform:rotate(-360deg)} }
        @keyframes spSpark  { 0%,100%{opacity:0;transform:scale(0) rotate(0deg)} 40%{opacity:1;transform:scale(1) rotate(180deg)} 70%{opacity:.6;transform:scale(.8) rotate(300deg)} }
        @keyframes spRing   { 0%{transform:scale(.6);opacity:.7} 100%{transform:scale(2.4);opacity:0} }
        @keyframes spParticle{ 0%{opacity:1;transform:translate(0,0)} 100%{opacity:0;transform:translate(var(--px),var(--py))} }
        @keyframes spOrbit1 { from{transform:rotate(0deg)   translateX(52px)} to{transform:rotate(360deg)  translateX(52px)} }
        @keyframes spOrbit2 { from{transform:rotate(120deg) translateX(44px)} to{transform:rotate(480deg)  translateX(44px)} }
        @keyframes spOrbit3 { from{transform:rotate(240deg) translateX(38px)} to{transform:rotate(600deg)  translateX(38px)} }
        @keyframes spDot    { 0%,80%,100%{transform:scale(.6);opacity:.4} 40%{transform:scale(1.2);opacity:1} }
        @keyframes spShimmer{ 0%{transform:translateX(-100%)} 100%{transform:translateX(200%)} }
        @keyframes spHex    { from{transform:rotate(45deg)} to{transform:rotate(405deg)} }
        .sp-hex   { position:absolute; border:1.5px solid rgba(0,210,100,0.14); border-radius:4px; transform:rotate(45deg); }
        .sp-spark { position:absolute; animation:spSpark 2.4s ease-in-out infinite; color:#00D264; font-size:13px; pointer-events:none; }
        .sp-ptcl  { position:absolute; width:5px; height:5px; border-radius:50%; animation:spParticle 3s ease-out infinite; }
        .sp-ring  { position:absolute; border-radius:50%; border:1.5px solid rgba(0,210,100,0.18); animation:spRing 2.8s ease-out infinite; }
        .sp-od    { position:absolute; top:50%; left:50%; border-radius:50%; }
        .sp-card::-webkit-scrollbar { display:none; }
        .sp-btn:hover { transform:scale(1.05) !important; }
      `}</style>

              <div
                onClick={(e) => e.stopPropagation()}
                className="sp-card"
                style={{
                  background: "#004632",
                  borderRadius: 24,
                  width: "100%",
                  maxWidth: 420,
                  maxHeight: "calc(100vh - 3rem)",
                  overflowY: "auto",
                  overflowX: "hidden",
                  animation: "spPopIn .55s cubic-bezier(.22,1,.36,1) both",
                  position: "relative",
                  padding: "2.25rem 1.75rem 1.75rem",
                  display: "flex",
                  flexDirection: "column",
                  alignItems: "center",
                  margin: "auto 0",
                }}
              >
                <button
                  type="button"
                  onClick={() => dismissWelcome()}
                  aria-label="Close welcome message"
                  style={{
                    position: "absolute",
                    top: 14,
                    right: 14,
                    zIndex: 12,
                    width: 38,
                    height: 38,
                    padding: 0,
                    border: "1px solid rgba(255,255,255,0.2)",
                    borderRadius: 11,
                    background: "rgba(255,255,255,0.08)",
                    color: "#fff",
                    display: "inline-flex",
                    alignItems: "center",
                    justifyContent: "center",
                    cursor: "pointer",
                  }}
                >
                  <X size={19} />
                </button>

                {[
                  { w: 60, h: 42, s: "top:5%;left:5%", d: "14s" },
                  { w: 78, h: 54, s: "top:4%;right:6%", d: "20s" },
                  { w: 45, h: 32, s: "bottom:10%;left:2%", d: "16s" },
                  { w: 66, h: 46, s: "bottom:7%;right:4%", d: "22s" },
                  { w: 40, h: 28, s: "top:38%;left:-1%", d: "18s" },
                  { w: 50, h: 35, s: "top:33%;right:-1%", d: "15s" },
                ].map((hx, i) => (
                  <div
                    key={i}
                    className="sp-hex"
                    style={{
                      width: hx.w,
                      height: hx.h,
                      ...Object.fromEntries(
                        hx.s
                          .split(";")
                          .map((s) => {
                            const [k, v] = s.split(":");
                            return [
                              k
                                ?.trim()
                                .replace(/-([a-z])/g, (_: string, l: string) =>
                                  l.toUpperCase(),
                                ),
                              v?.trim(),
                            ];
                          })
                          .filter(([k]) => k),
                      ),
                      animation: `spSpin ${hx.d} linear infinite ${i % 2 ? "reverse" : ""}`,
                    }}
                  />
                ))}
                {[
                  ["12%", "15%", "0s"],
                  ["80%", "10%", ".4s"],
                  ["68%", "75%", ".7s"],
                  ["20%", "68%", "1s"],
                  ["48%", "6%", "1.3s"],
                  ["88%", "48%", "1.6s"],
                  ["30%", "90%", "1.9s"],
                ].map(([t, l, del], i) => (
                  <div
                    key={i}
                    className="sp-spark"
                    style={{
                      top: t,
                      left: l,
                      animationDelay: del,
                      animationDuration: `${2 + i * 0.2}s`,
                    }}
                  >
                    ✦
                  </div>
                ))}
                {[
                  { c: "#00D264", px: "-45px", py: "-55px" },
                  { c: "#8CFF8C", px: "50px", py: "-40px" },
                  { c: "rgba(255,255,255,.5)", px: "-55px", py: "45px" },
                  { c: "#00D264", px: "48px", py: "52px" },
                  { c: "#8CFF8C", px: "-30px", py: "60px" },
                  { c: "rgba(255,255,255,.4)", px: "55px", py: "-30px" },
                ].map((p, i) => (
                  <div
                    key={i}
                    className="sp-ptcl"
                    style={{
                      top: `${20 + i * 12}%`,
                      left: `${8 + i * 14}%`,
                      background: p.c,
                      ["--px" as any]: p.px,
                      ["--py" as any]: p.py,
                      animationDelay: `${i * 0.45}s`,
                      animationDuration: `${3 + i * 0.3}s`,
                    }}
                  />
                ))}
                {[
                  { s: "90px", d: "0s" },
                  { s: "160px", d: ".9s" },
                  { s: "240px", d: "1.8s" },
                ].map((r, i) => (
                  <div
                    key={i}
                    className="sp-ring"
                    style={{
                      width: r.s,
                      height: r.s,
                      top: `calc(38% - ${parseInt(r.s) / 2}px)`,
                      left: `calc(50% - ${parseInt(r.s) / 2}px)`,
                      animationDelay: r.d,
                    }}
                  />
                ))}

                <div
                  style={{
                    position: "relative",
                    width: 120,
                    height: 120,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    animation: "spFloat 3.2s ease-in-out infinite",
                    marginBottom: "1.375rem",
                    zIndex: 5,
                  }}
                >
                  <div
                    style={{
                      position: "absolute",
                      inset: -18,
                      borderRadius: "50%",
                      background:
                        "radial-gradient(circle,rgba(0,210,100,0.2) 0%,transparent 70%)",
                    }}
                  />
                  <div
                    style={{
                      position: "absolute",
                      inset: -6,
                      borderRadius: "50%",
                      border: "2px dashed rgba(0,210,100,0.32)",
                      animation: "spSpin 12s linear infinite",
                    }}
                  />
                  <div
                    style={{
                      position: "absolute",
                      inset: -14,
                      borderRadius: "50%",
                      border: "1.5px dashed rgba(140,255,140,0.18)",
                      animation: "spRevSpin 18s linear infinite",
                    }}
                  />
                  {[
                    {
                      cls: "od1",
                      bg: "#8CFF8C",
                      s: "10px",
                      tx: "52px",
                      dur: "3.5s",
                      del: "0s",
                    },
                    {
                      cls: "od2",
                      bg: "#00D264",
                      s: "8px",
                      tx: "44px",
                      dur: "2.8s",
                      del: "0s",
                    },
                    {
                      cls: "od3",
                      bg: "rgba(255,255,255,.55)",
                      s: "7px",
                      tx: "38px",
                      dur: "4.2s",
                      del: "0s",
                    },
                  ].map((o, i) => (
                    <div
                      key={i}
                      className="sp-od"
                      style={{
                        width: o.s,
                        height: o.s,
                        background: o.bg,
                        marginTop: `-${parseInt(o.s) / 2}px`,
                        marginLeft: `-${parseInt(o.s) / 2}px`,
                        animation: `spOrbit${i + 1} ${o.dur} linear infinite`,
                      }}
                    />
                  ))}
                  <div
                    style={{
                      width: 110,
                      height: 110,
                      borderRadius: "50%",
                      background: "linear-gradient(135deg,#00D264,#004632)",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      border: "3px solid rgba(255,255,255,0.22)",
                      position: "relative",
                      zIndex: 2,
                      boxShadow:
                        "0 0 0 8px rgba(0,210,100,0.12),0 0 0 18px rgba(0,210,100,0.06)",
                    }}
                  >
                    <svg
                      width="50"
                      height="50"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="#fff"
                      strokeWidth="1.8"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    >
                      <circle cx="12" cy="8" r="4" />
                      <path d="M4 20c0-4 3.6-7 8-7s8 3 8 7" />
                    </svg>
                  </div>
                </div>

                <div
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 6,
                    background: "rgba(0,210,100,0.16)",
                    border: "1px solid rgba(0,210,100,0.35)",
                    borderRadius: 999,
                    padding: "4px 14px",
                    fontSize: 11,
                    fontWeight: 600,
                    color: "#8CFF8C",
                    letterSpacing: ".07em",
                    textTransform: "uppercase",
                    marginBottom: "0.75rem",
                    animation: "spSlide .4s .18s both",
                  }}
                >
                  <div
                    style={{
                      width: 5,
                      height: 5,
                      borderRadius: "50%",
                      background: "#00D264",
                      animation: "spDot 1.4s infinite",
                    }}
                  />
                  {c.badge}
                </div>

                <div
                  style={{
                    textAlign: "center",
                    marginBottom: ".5rem",
                    animation: "spSlide .45s .28s both",
                    zIndex: 5,
                  }}
                >
                  <div
                    style={{
                      fontSize: 15,
                      color: "rgba(255,255,255,0.6)",
                      fontWeight: 400,
                      marginBottom: 2,
                    }}
                  >
                    {c.nameLine1}
                  </div>
                  <div
                    style={{
                      fontSize: 28,
                      fontWeight: 700,
                      color: "#00D264",
                      lineHeight: 1.2,
                      letterSpacing: "-.02em",
                    }}
                  >
                    {c.nameLine2}
                  </div>
                </div>

                <p
                  style={{
                    fontSize: 13,
                    color: "rgba(255,255,255,0.65)",
                    textAlign: "center",
                    lineHeight: 1.72,
                    maxWidth: 300,
                    margin: "0 0 1.5rem",
                    animation: "spSlide .45s .38s both",
                    zIndex: 5,
                  }}
                >
                  {c.msg}
                </p>

                {/* <div style={{ display:'flex', gap:6, marginBottom:'1.375rem', animation:'spSlide .4s .46s both' }}>
          {[true,false,false].map((active,i) => (
            <div key={i} style={{ width:active?22:8, height:8, borderRadius:4, background:active?'#00D264':'rgba(255,255,255,0.2)', transition:'width .3s' }} />
          ))}
        </div> */}

                <button
                  className="sp-btn"
                  onClick={() => {
                    const role = normalizeRole(employeeData.Role);
                    const sectionMap: Record<string, string> = {
                      "Super Admin": "home",
                      "L&D Site Admin": "onboarding",
                      Buddy: "meetings",
                      "New Joinee": "meetings",
                    };
                    const targetSection = sectionMap[role] || "home";
                    dismissWelcome(targetSection);
                  }}
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 8,
                    padding: ".7rem 2.25rem",
                    background: "#00D264",
                    color: "#004632",
                    border: "none",
                    borderRadius: 999,
                    fontSize: 14,
                    fontWeight: 700,
                    cursor: "pointer",
                    fontFamily: "inherit",
                    animation: "spSlide .4s .54s both",
                    position: "relative",
                    overflow: "hidden",
                    transition: "transform .18s",
                    zIndex: 5,
                  }}
                >
                  <span>{c.btn}</span>
                  <span style={{ fontSize: 16 }}>→</span>
                  <div
                    style={{
                      position: "absolute",
                      inset: 0,
                      background:
                        "linear-gradient(90deg,transparent,rgba(255,255,255,0.22),transparent)",
                      transform: "translateX(-100%)",
                      animation: "spShimmer 2.5s 1.5s infinite",
                    }}
                  />
                </button>

                {/* <button onClick={() => setShowWelcome(false)} style={{ marginTop:'.625rem', fontSize:12, color:'rgba(255,255,255,0.35)', background:'none', border:'none', cursor:'pointer', fontFamily:'inherit', animation:'spSlide .4s .62s both', transition:'color .15s' }}>
          Maybe later
        </button> */}
              </div>
            </div>
          );
        })()}

      {showRankCelebration &&
        currentBuddyRanking &&
        normalizeRole(employeeData?.Role) === "Buddy" &&
        normalizeStatus(employeeData?.EmployeeStatus) === "active" && (
        <TopPerformerPopup
          ranking={currentBuddyRanking}
          context={props.context}
          photoUrl={props.userloginDetails?.photoUrl}
          onClose={closeRankCelebration}
        />
      )}

      <BuddyLoader
        visible={initialDataLoading}
        variant="default"
        message="Loading dashboard…"
        subMessage="Preparing your Buddy System dashboard."
      />

      <BuddyLoader
        visible={globalLoader.visible}
        variant={globalLoader.variant}
        message={globalLoader.message}
      />
    </div>
  );
};

const BuddyApp: React.FC<IBuddyAppProps> = (props) => (
  <SnackbarProvider>
    <BuddyAppContent {...props} />
  </SnackbarProvider>
);

export default BuddyApp;