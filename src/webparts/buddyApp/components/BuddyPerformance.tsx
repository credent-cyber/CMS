/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-var-requires */
/* eslint-disable @rushstack/no-new-null */
/* eslint-disable max-lines */
/* eslint-disable no-void */


import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
import {
  Activity,
  CalendarCheck2,
  ChevronLeft,
  ChevronRight,
  Info,
  MapPin,
  RefreshCw,
  Sparkles,
  Star,
  Trophy,
  Users,
} from "lucide-react";

import LIST_CONFIG, {
  getThresholdSafeListItems,
} from "../../../config/spListConfig";
import styles from "./BuddyPerformance.module.scss";

const SCORE_WEIGHTS = {
  rating: 15,
  onTime: 60,
  completion: 15,
  newJoiners: 10,
} as const;

const MAX_NEW_JOINERS_FOR_FULL_SCORE = 4;
const TIMELINE_TOLERANCE_DAYS = 5;

const RATING_FIELDS = [
  "Rating",
  "NJRRating",
  "NJRating",
  "NewJoineeRating",
  "OverallRating",
  "FeedbackRating",
  "InteractionRating",
  "Score",
];

const ANSWER_DATE_FIELDS = [
  "InteractionDate",
  "ActualInteractionDate",
  "CompletedOn",
  "SubmittedOn",
  "Modified",
  "Created",
];

const ANSWER_NJ_FIELDS = ["NJID", "GlobalIDofNJ", "NewJoineeID", "NJGlobalID"];
const ANSWER_BUDDY_EMAIL_FIELDS = [
  "BuddyEmailID",
  "BuddyEmail",
  "AssignedBuddyEmail",
];

export interface ICurrentBuddyRanking {
  rank: number;
  name: string;
  email: string;
  location: string;
  score: number;
  averageRating: number;
  completed: number;
  monthKey: string;
  monthLabel: string;
}

interface IBuddyPerformanceProps {
  context: any;
  isSuperAdmin: boolean;
  userLocation?: string;
  currentUserEmail?: string;
  currentUserGlobalId?: string;
  currentUserName?: string;
  onCurrentUserRankingChange?: (ranking: ICurrentBuddyRanking | null) => void;
}

interface IMeetingPair {
  planned?: string;
  actual?: string;
  status?: string;
}

interface IBuddySource {
  key: string;
  name: string;
  email: string;
  globalId: string;
  location: string;
}

interface IMutablePerformance extends IBuddySource {
  scheduled: number;
  completed: number;
  onTime: number;
  outsideTimeline: number;
  missingPlannedDates: number;
  ratings: number[];
  newJoinerIds: Set<string>;
}

interface IBuddyRanking extends IBuddySource {
  rank: number;
  score: number;
  averageRating: number;
  ratingCount: number;
  scheduled: number;
  completed: number;
  onTime: number;
  outsideTimeline: number;
  missingPlannedDates: number;
  completionRate: number;
  onTimeRate: number;
  newJoiners: number;
  ratingPoints: number;
  onTimePoints: number;
  completionPoints: number;
  newJoinerPoints: number;
}

interface ILocationBoard {
  location: string;
  buddies: IBuddyRanking[];
}

const normalize = (value?: unknown): string =>
  String(value || "").trim().toLowerCase();

const firstValue = (item: any, fields: string[]): any => {
  for (const field of fields) {
    const value = item?.[field];
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      return value;
    }
  }
  return undefined;
};

const parseDate = (value?: unknown): Date | null => {
  if (!value) return null;
  const date = new Date(String(value));
  return Number.isNaN(date.getTime()) ? null : date;
};

const isInMonth = (value: unknown, month: Date): boolean => {
  const date = parseDate(value);
  return Boolean(
    date &&
      date.getFullYear() === month.getFullYear() &&
      date.getMonth() === month.getMonth(),
  );
};

const dayDifference = (leftValue?: unknown, rightValue?: unknown): number => {
  const left = parseDate(leftValue);
  const right = parseDate(rightValue);
  if (!left || !right) return Number.POSITIVE_INFINITY;
  return Math.abs(left.getTime() - right.getTime()) / 86400000;
};

const toRating = (item: any): number | null => {
  const directValue = firstValue(item, RATING_FIELDS);
  let candidate = directValue !== undefined
    ? directValue
    : Object.keys(item || {}).find((key) => /rating/i.test(key))
      ? item[Object.keys(item).find((key) => /rating/i.test(key)) as string]
      : undefined;
  const questionText = Object.keys(item || {})
    .filter((key) => /question|prompt|label/i.test(key))
    .map((key) => String(item[key] || ""))
    .join(" ");
  if (candidate === undefined && /rating|rate\s+(the\s+)?buddy/i.test(questionText)) {
    candidate = firstValue(item, ["Answer", "Response", "Value", "AnswerValue"]);
  }
  const rating = Number(candidate);
  return Number.isFinite(rating) && rating >= 1 && rating <= 5 ? rating : null;
};

const isRatingFeedbackMilestone = (item: any): boolean => {
  const meeting = normalize(
    firstValue(item, ["Meetings", "Meeting", "Milestone"]),
  );
  return /^(3rd|6th)(?:\s+(?:meeting\s+)?feedback)?$/.test(meeting);
};

const getMeetingPairs = (row: any): IMeetingPair[] => [
  { planned: row.FirstInteractionDate, actual: row.FirstActualInteractionDate, status: row.FirstMeetingStatus },
  { planned: row.SecondInteractionDate, actual: row.SecondActualInteractionDate, status: row.SecondMeetingStatus },
  { planned: row.ThirdInteractionDate, actual: row.ThirdActualInteractionDate, status: row.ThirdMeetingStatus },
  { planned: row.FourthInteractionDate, actual: row.FourthActualInteractionDate, status: row.FourthMeetingStatus },
  { planned: row.FifthInteractionDate, actual: row.FifthActualInteractionDate, status: row.FifthMeetingStatus },
  { planned: row.SixthInteractionDate, actual: row.SixthActualInteractionDate, status: row.SixthMeetingStatus },
];

const getBuddySource = (row: any): IBuddySource => {
  const email = String(row.BuddyEmailID || "").trim();
  const globalId = String(row.GlobalIDofBuddy || "").trim();
  const name = String(row.BuddyNameRefChoiceColumn || "Buddy").trim();
  return {
    key: normalize(email || globalId || name),
    name,
    email,
    globalId,
    location: String(row.LocationofBuddy || "Unassigned").trim() || "Unassigned",
  };
};

const monthLabel = (month: Date): string =>
  month.toLocaleDateString("en-IN", { month: "long", year: "numeric" });

const initials = (name: string): string =>
  name
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase())
    .join("") || "B";

const BuddyAvatar: React.FC<{
  name: string;
  email: string;
  context: any;
  rank: number;
}> = ({ name, email, context, rank }) => {
  const [photoUrl, setPhotoUrl] = useState("");
  const rankAvatarClass =
    rank === 1
      ? styles.rankAvatar1
      : rank === 2
        ? styles.rankAvatar2
        : styles.rankAvatar3;
  const accountName = useMemo(() => {
    const rawAccountName = String(email || "").trim();
    return (
      rawAccountName.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i)?.[0] ||
      rawAccountName.split("|").pop()?.replace(/^mailto:/i, "").trim() ||
      ""
    );
  }, [email]);

  useEffect(() => {
    let isActive = true;
    let objectUrl = "";
    setPhotoUrl("");

    if (!accountName || !context?.msGraphClientFactory) {
      return () => {
        isActive = false;
      };
    }

    context.msGraphClientFactory
      .getClient("3")
      .then((client: any) =>
        client
          .api(`/users/${encodeURIComponent(accountName)}/photo/$value`)
          .responseType("blob" as any)
          .get(),
      )
      .then((photo: any) => {
        if (!isActive || !photo) return;
        const photoBlob =
          photo instanceof Blob
            ? photo
            : new Blob([photo], { type: "image/jpeg" });
        if (!photoBlob.size) return;
        objectUrl = URL.createObjectURL(photoBlob);
        setPhotoUrl(objectUrl);
      })
      .catch(() => {
        // Keep initials when no Microsoft 365 profile photo is available.
      });

    return () => {
      isActive = false;
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [accountName, context]);

  return (
    <div className={`${styles.avatarWrap} ${rankAvatarClass}`}>
      {photoUrl ? (
        <img
          src={photoUrl}
          alt={name}
          loading="eager"
          decoding="async"
          onError={() => setPhotoUrl("")}
        />
      ) : (
        <span>{initials(name)}</span>
      )}
    </div>
  );
};

const rankClass = (rank: number): string =>
  rank === 1 ? styles.rank1 : rank === 2 ? styles.rank2 : styles.rank3;

const podiumPosition = (rank: number): number =>
  rank === 2 ? 0 : rank === 1 ? 1 : 2;

const Medal: React.FC<{ rank: number }> = ({ rank }) => (
  <span className={`${styles.rankMedal} ${rankClass(rank)}`}>
    <span
      className={`${styles.medalLaurel} ${styles.medalLaurelLeft}`}
      aria-hidden="true"
    >
      <svg viewBox="0 0 18 30" focusable="false">
        <path d="M14.5 27C7.8 22.2 4.9 13.7 6.2 3.5" />
        <ellipse cx="6.5" cy="6" rx="2.1" ry="3.8" transform="rotate(-24 6.5 6)" />
        <ellipse cx="6.2" cy="12.2" rx="2.1" ry="3.8" transform="rotate(-39 6.2 12.2)" />
        <ellipse cx="8" cy="18.2" rx="2.1" ry="3.8" transform="rotate(-51 8 18.2)" />
        <ellipse cx="11.2" cy="23.4" rx="2.1" ry="3.8" transform="rotate(-62 11.2 23.4)" />
      </svg>
    </span>
    <strong>{rank}</strong>
    <span
      className={`${styles.medalLaurel} ${styles.medalLaurelRight}`}
      aria-hidden="true"
    >
      <svg viewBox="0 0 18 30" focusable="false">
        <path d="M14.5 27C7.8 22.2 4.9 13.7 6.2 3.5" />
        <ellipse cx="6.5" cy="6" rx="2.1" ry="3.8" transform="rotate(-24 6.5 6)" />
        <ellipse cx="6.2" cy="12.2" rx="2.1" ry="3.8" transform="rotate(-39 6.2 12.2)" />
        <ellipse cx="8" cy="18.2" rx="2.1" ry="3.8" transform="rotate(-51 8 18.2)" />
        <ellipse cx="11.2" cy="23.4" rx="2.1" ry="3.8" transform="rotate(-62 11.2 23.4)" />
      </svg>
    </span>
  </span>
);

const PosterMetric: React.FC<{
  icon: React.ReactNode;
  value: string;
  label: string;
}> = ({ icon, value, label }) => (
  <div className={styles.posterMetric}>
    <span>{icon}</span>
    <strong>{value}</strong>
    <small>{label}</small>
  </div>
);

const Metric: React.FC<{
  icon: React.ReactNode;
  label: string;
  value: string;
}> = ({ icon, label, value }) => (
  <div className={styles.detailMetric}>
    <span>{icon}</span>
    <div>
      <strong>{value}</strong>
      <small>{label}</small>
    </div>
  </div>
);

export const BuddyPerformance: React.FC<IBuddyPerformanceProps> = ({
  context,
  isSuperAdmin,
  userLocation = "",
  currentUserEmail = "",
  currentUserGlobalId = "",
  currentUserName = "",
  onCurrentUserRankingChange,
}) => {
  const isMountedRef = useRef(true);
  const [selectedMonth, setSelectedMonth] = useState(
    () => new Date(new Date().getFullYear(), new Date().getMonth(), 1),
  );
  const [boards, setBoards] = useState<ILocationBoard[]>([]);
  const [locationFilter, setLocationFilter] = useState("All Locations");
  const defaultLocationAppliedRef = useRef(false);
  const [loading, setLoading] = useState(true);
  const [errorMessage, setErrorMessage] = useState("");
  const currentMonth = useMemo(() => {
    const today = new Date();
    return new Date(today.getFullYear(), today.getMonth(), 1);
  }, []);
  const earliestMonth = useMemo(
    () => new Date(currentMonth.getFullYear() - 1, currentMonth.getMonth(), 1),
    [currentMonth],
  );
  const canMoveToPreviousMonth =
    selectedMonth.getTime() > earliestMonth.getTime();
  const canMoveToNextMonth = selectedMonth.getTime() < currentMonth.getTime();

  const isCurrentUserBuddy = useCallback(
    (buddy: IBuddyRanking): boolean => {
      const normalizedEmail = normalize(currentUserEmail);
      if (normalizedEmail) return normalize(buddy.email) === normalizedEmail;

      const normalizedGlobalId = normalize(currentUserGlobalId);
      if (normalizedGlobalId) {
        return normalize(buddy.globalId) === normalizedGlobalId;
      }

      const normalizedName = normalize(currentUserName);
      return Boolean(
        normalizedName && normalize(buddy.name) === normalizedName,
      );
    },
    [currentUserEmail, currentUserGlobalId, currentUserName],
  );

  useEffect(() => {
    isMountedRef.current = true;
    return () => {
      isMountedRef.current = false;
    };
  }, []);

  const loadRankings = useCallback(async (): Promise<void> => {
    setLoading(true);
    setErrorMessage("");
    try {
      if (!isSuperAdmin && !String(userLocation || "").trim()) {
        if (isMountedRef.current) setBoards([]);
        return;
      }
      const allocationSelect =
        "ID,GlobalIDofBuddy,LocationofBuddy,BuddyEmailID,BuddyNameRefChoiceColumn," +
        "GlobalIDofNJ,NJEmailID,Status," +
        "FirstInteractionDate,SecondInteractionDate,ThirdInteractionDate," +
        "FourthInteractionDate,FifthInteractionDate,SixthInteractionDate," +
        "FirstActualInteractionDate,SecondActualInteractionDate,ThirdActualInteractionDate," +
        "FourthActualInteractionDate,FifthActualInteractionDate,SixthActualInteractionDate," +
        "FirstMeetingStatus,SecondMeetingStatus,ThirdMeetingStatus," +
        "FourthMeetingStatus,FifthMeetingStatus,SixthMeetingStatus";

      const safeUserLocation = String(userLocation || "").replace(/'/g, "''");
      const allocationRows = await getThresholdSafeListItems({
        web: sp.web,
        listTitle: LIST_CONFIG.LISTS.BuddyAllocate,
        select: allocationSelect,
        filter: !isSuperAdmin && safeUserLocation
          ? `LocationofBuddy eq '${safeUserLocation}'`
          : undefined,
      });
      let answerRows: any[] = [];
      try {
        answerRows = await getThresholdSafeListItems({
          web: sp.web,
          listTitle: LIST_CONFIG.LISTS.MEETING_ANSWERS,
          select: "*",
        });
      } catch (ratingError) {
        console.warn("Buddy Performance ratings are temporarily unavailable:", ratingError);
      }

      const performanceByBuddy = new Map<string, IMutablePerformance>();
      const buddyByNewJoiner = new Map<string, IBuddySource>();
      const buddyByEmail = new Map<string, IBuddySource>();

      const ensurePerformance = (source: IBuddySource): IMutablePerformance => {
        const existing = performanceByBuddy.get(source.key);
        if (existing) return existing;
        const created: IMutablePerformance = {
          ...source,
          scheduled: 0,
          completed: 0,
          onTime: 0,
          outsideTimeline: 0,
          missingPlannedDates: 0,
          ratings: [],
          newJoinerIds: new Set<string>(),
        };
        performanceByBuddy.set(source.key, created);
        return created;
      };

      (allocationRows || []).forEach((row: any) => {
        if (normalize(row.Status) !== "active") return;

        const source = getBuddySource(row);
        if (!source.key) return;
        const njId = normalize(row.GlobalIDofNJ || row.NJEmailID);
        if (njId) buddyByNewJoiner.set(njId, source);
        if (source.email) buddyByEmail.set(normalize(source.email), source);

        const monthlyPairs = getMeetingPairs(row).filter((pair) =>
          isInMonth(parseDate(pair.actual) ? pair.actual : pair.planned, selectedMonth),
        );
        if (!monthlyPairs.length) return;

        const performance = ensurePerformance(source);
        if (njId) performance.newJoinerIds.add(njId);
        monthlyPairs.forEach((pair) => {
          performance.scheduled += 1;
          const actualDate = parseDate(pair.actual);
          const plannedDate = parseDate(pair.planned);
          if (actualDate) {
            performance.completed += 1;
            if (!plannedDate) {
              performance.missingPlannedDates += 1;
            } else if (dayDifference(plannedDate, actualDate) <= TIMELINE_TOLERANCE_DAYS) {
              performance.onTime += 1;
            } else {
              performance.outsideTimeline += 1;
            }
          }
        });
      });

      (answerRows || []).forEach((row: any) => {
        if (normalize(row.Status) !== "completed") return;
        if (!isRatingFeedbackMilestone(row)) return;
        const rating = toRating(row);
        const answerDate = firstValue(row, ANSWER_DATE_FIELDS);
        if (rating === null || !isInMonth(answerDate, selectedMonth)) return;

        const buddyEmail = normalize(firstValue(row, ANSWER_BUDDY_EMAIL_FIELDS));
        const njId = normalize(firstValue(row, ANSWER_NJ_FIELDS));
        const source =
          (buddyEmail && buddyByEmail.get(buddyEmail)) ||
          (njId && buddyByNewJoiner.get(njId));
        if (!source) return;
        ensurePerformance(source).ratings.push(rating);
      });

      const rankings = Array.from(performanceByBuddy.values()).map((item) => {
        const averageRating = item.ratings.length
          ? item.ratings.reduce((total, rating) => total + rating, 0) /
            item.ratings.length
          : 0;
        const scheduledBase = Math.max(item.scheduled, item.completed);
        const completionRate = scheduledBase
          ? Math.min(100, (item.completed / scheduledBase) * 100)
          : 0;
        const onTimeRate = scheduledBase
          ? (item.onTime / scheduledBase) * 100
          : 0;
        const ratingPoints = (averageRating / 5) * SCORE_WEIGHTS.rating;
        const onTimePoints = (onTimeRate / 100) * SCORE_WEIGHTS.onTime;
        const completionPoints =
          (completionRate / 100) * SCORE_WEIGHTS.completion;
        const newJoinerPoints =
          Math.min(
            1,
            item.newJoinerIds.size / MAX_NEW_JOINERS_FOR_FULL_SCORE,
          ) * SCORE_WEIGHTS.newJoiners;
        const score =
          ratingPoints + onTimePoints + completionPoints + newJoinerPoints;

        return {
          ...item,
          rank: 0,
          score: Number(score.toFixed(1)),
          averageRating: Number(averageRating.toFixed(1)),
          ratingCount: item.ratings.length,
          onTime: item.onTime,
          outsideTimeline: item.outsideTimeline,
          missingPlannedDates: item.missingPlannedDates,
          completionRate: Math.round(completionRate),
          onTimeRate: Math.round(onTimeRate),
          newJoiners: item.newJoinerIds.size,
          ratingPoints: Number(ratingPoints.toFixed(1)),
          onTimePoints: Number(onTimePoints.toFixed(1)),
          completionPoints: Number(completionPoints.toFixed(1)),
          newJoinerPoints: Number(newJoinerPoints.toFixed(1)),
        } as IBuddyRanking;
      });

      const byLocation = new Map<string, IBuddyRanking[]>();
      rankings.forEach((ranking) => {
        const location = ranking.location || "Unassigned";
        const group = byLocation.get(location) || [];
        group.push(ranking);
        byLocation.set(location, group);
      });

      const nextBoards = Array.from(byLocation.entries())
        .map(([location, buddies]) => ({
          location,
          buddies: buddies
            .sort(
              (left, right) =>
                right.score - left.score ||
                right.averageRating - left.averageRating ||
                right.completed - left.completed ||
                left.name.localeCompare(right.name),
            )
            .slice(0, 3)
            .map((buddy, index) => ({ ...buddy, rank: index + 1 })),
        }))
        .sort((left, right) => left.location.localeCompare(right.location));

      if (isMountedRef.current) setBoards(nextBoards);
    } catch (error) {
      console.error("Unable to load monthly Buddy Performance:", error);
      if (isMountedRef.current) {
        setErrorMessage(
          "We couldn't load the monthly Buddy Performance right now. Please try again shortly.",
        );
      }
    } finally {
      if (isMountedRef.current) setLoading(false);
    }
  }, [isSuperAdmin, selectedMonth, userLocation]);

  useEffect(() => {
    void loadRankings();
  }, [loadRankings]);

  const locations = useMemo(
    () => boards.map((board) => board.location),
    [boards],
  );

  useEffect(() => {
    if (!isSuperAdmin || defaultLocationAppliedRef.current) return;

    const currentLocation = normalize(userLocation);
    if (!currentLocation || locations.length === 0) return;

    const matchingLocation = locations.find(
      (location) => normalize(location) === currentLocation,
    );
    if (!matchingLocation) return;

    setLocationFilter(matchingLocation);
    defaultLocationAppliedRef.current = true;
  }, [isSuperAdmin, locations, userLocation]);

  const visibleBoards = useMemo(() => {
    if (isSuperAdmin) {
      return locationFilter === "All Locations"
        ? boards
        : boards.filter((board) => board.location === locationFilter);
    }
    const currentLocation = normalize(userLocation);
    return currentLocation
      ? boards.filter((board) => normalize(board.location) === currentLocation)
      : [];
  }, [boards, isSuperAdmin, locationFilter, userLocation]);

  useEffect(() => {
    if (!onCurrentUserRankingChange || loading) return;

    const today = new Date();
    const isCurrentMonth =
      selectedMonth.getFullYear() === today.getFullYear() &&
      selectedMonth.getMonth() === today.getMonth();
    const hasCurrentUserIdentity = Boolean(
      normalize(currentUserEmail) ||
        normalize(currentUserGlobalId) ||
        normalize(currentUserName),
    );

    if (!isCurrentMonth || !hasCurrentUserIdentity) {
      onCurrentUserRankingChange(null);
      return;
    }

    const currentRanking = boards
      .reduce<IBuddyRanking[]>((all, board) => all.concat(board.buddies), [])
      .find(isCurrentUserBuddy);

    onCurrentUserRankingChange(
      currentRanking
        ? {
            rank: currentRanking.rank,
            name: currentRanking.name,
            email: currentRanking.email,
            location: currentRanking.location,
            score: currentRanking.score,
            averageRating: currentRanking.averageRating,
            completed: currentRanking.completed,
            monthKey: `${selectedMonth.getFullYear()}-${
              selectedMonth.getMonth() + 1 < 10 ? "0" : ""
            }${selectedMonth.getMonth() + 1}`,
            monthLabel: monthLabel(selectedMonth),
          }
        : null,
    );
  }, [
    boards,
    currentUserEmail,
    currentUserGlobalId,
    currentUserName,
    isCurrentUserBuddy,
    loading,
    onCurrentUserRankingChange,
    selectedMonth,
  ]);

  const boardRailRef = useRef<HTMLDivElement>(null);

  const scrollBoardRail = (direction: -1 | 1): void => {
    const rail = boardRailRef.current;
    if (!rail) return;
    rail.scrollBy({
      left: direction * Math.max(320, rail.clientWidth * 0.92),
      behavior: "smooth",
    });
  };

  const showAllLocationSlides =
    isSuperAdmin && locationFilter === "All Locations" && visibleBoards.length > 1;

  const summary = useMemo(() => {
    const rankedBuddies = boards.reduce(
      (total, board) => total + board.buddies.length,
      0,
    );
    const ratingValues = boards.flatMap((board) =>
      board.buddies
        .filter((buddy) => buddy.averageRating > 0)
        .map((buddy) => buddy.averageRating),
    );
    return {
      locations: boards.length,
      buddies: rankedBuddies,
      averageRating: ratingValues.length
        ? (
            ratingValues.reduce((total, rating) => total + rating, 0) /
            ratingValues.length
          ).toFixed(1)
        : "—",
    };
  }, [boards]);

  const moveMonth = (offset: number): void => {
    setSelectedMonth((current) => {
      const requestedMonth = new Date(
        current.getFullYear(),
        current.getMonth() + offset,
        1,
      );
      if (requestedMonth.getTime() < earliestMonth.getTime()) {
        return earliestMonth;
      }
      if (requestedMonth.getTime() > currentMonth.getTime()) {
        return currentMonth;
      }
      return requestedMonth;
    });
  };

  return (
    <section className={styles.performanceSection}>
      <div className={styles.performanceHeader}>
        <div className={styles.headerIdentity}>
          <span className={styles.headerIcon}><Trophy size={21} /></span>
          <div>
            <p>Monthly Recognition</p>
            <h2>Buddy Performance</h2>
            <span>Celebrating consistent support, timely interactions and New Joiner feedback.</span>
          </div>
        </div>

        <div className={styles.headerTools}>
          {isSuperAdmin && locations.length > 0 && (
            <label className={styles.locationSelect}>
              <MapPin size={14} />
              <select
                value={locationFilter}
                onChange={(event) => setLocationFilter(event.target.value)}
                aria-label="Filter performance by location"
              >
                <option>All Locations</option>
                {locations.map((location) => (
                  <option key={location}>{location}</option>
                ))}
              </select>
            </label>
          )}
          <div className={styles.monthControl}>
            <button
              type="button"
              onClick={() => moveMonth(-1)}
              disabled={!canMoveToPreviousMonth}
              aria-label="Previous month"
            >
              <ChevronLeft size={15} />
            </button>
            <div><CalendarCheck2 size={14} /><strong>{monthLabel(selectedMonth)}</strong></div>
            <button
              type="button"
              onClick={() => moveMonth(1)}
              disabled={!canMoveToNextMonth}
              aria-label="Next month"
            >
              <ChevronRight size={15} />
            </button>
          </div>
          <button
            type="button"
            className={styles.refreshButton}
            onClick={() => { void loadRankings(); }}
            disabled={loading}
            aria-label="Refresh rankings"
            style={{ cursor: loading ? "default" : "pointer" }}
          >
            <RefreshCw
              size={14}
              className={loading ? styles.spinning : ""}
              style={{ pointerEvents: "none" }}
            />
          </button>
        </div>
      </div>

      {isSuperAdmin && (
        <div className={styles.summaryStrip}>
          <div><MapPin size={14} /><span><strong>{summary.locations}</strong> locations</span></div>
          <div><Trophy size={14} /><span><strong>{summary.buddies}</strong> monthly leaders</span></div>
          <div><Star size={14} /><span><strong>{summary.averageRating}</strong> average rating</span></div>
          <p><Sparkles size={13} /> Timeline consistency carries the highest score weight.</p>
        </div>
      )}

      {errorMessage ? (
        <div className={styles.messageState}>
          <Info size={20} />
          <p>{errorMessage}</p>
          <button type="button" onClick={() => { void loadRankings(); }}>Try again</button>
        </div>
      ) : loading ? (
        <div className={styles.loadingBoards}>
          {[1, 2].map((item) => <div key={item} className={styles.boardSkeleton} />)}
        </div>
      ) : visibleBoards.length === 0 ? (
        <div className={styles.messageState}>
          <Trophy size={22} />
          <h3>No rankings available yet</h3>
          <p>
            {isSuperAdmin || userLocation
              ? `Completed interactions and New Joiner ratings for ${monthLabel(selectedMonth)} will appear here.`
              : "Your location is not available in the employee profile. Please contact the Buddy System administrator."}
          </p>
        </div>
      ) : (
        <div className={styles.boardScroller}>
          {showAllLocationSlides && (
            <div className={styles.boardSliderNav}>
              <span>Swipe or use arrows to view each location</span>
              <div>
                <button
                  type="button"
                  onClick={() => scrollBoardRail(-1)}
                  aria-label="Previous location"
                >
                  <ChevronLeft size={18} />
                </button>
                <button
                  type="button"
                  onClick={() => scrollBoardRail(1)}
                  aria-label="Next location"
                >
                  <ChevronRight size={18} />
                </button>
              </div>
            </div>
          )}
          <div
            ref={boardRailRef}
            className={`${styles.locationBoards} ${
              showAllLocationSlides ? styles.locationBoardsHorizontal : ""
            }`}
          >
            {visibleBoards.map((board) => (
              <article key={board.location} className={styles.locationBoard}>
                <div className={styles.posterHeading}>
                  <h3>Top Performers</h3>
                  <span>{monthLabel(selectedMonth)}</span>
                  <small><MapPin size={12} /> {board.location}</small>
                </div>

                <div
                  className={styles.podium}
                  role="list"
                  aria-label={`${board.location} buddy rankings`}
                >
                  {board.buddies
                    .slice()
                    .sort(
                      (left, right) =>
                        podiumPosition(left.rank) - podiumPosition(right.rank),
                    )
                    .map((buddy) => {
                      const isCurrentUser = isCurrentUserBuddy(buddy);
                      return (
                        <div
                          key={buddy.key}
                          className={`${styles.rankCard} ${rankClass(buddy.rank)} ${
                            isSuperAdmin ? styles.flippableCard : ""
                          } ${isCurrentUser ? styles.currentUserCard : ""}`}
                          role="listitem"
                          tabIndex={isSuperAdmin ? 0 : undefined}
                          aria-label={`${buddy.name}, rank ${buddy.rank}${
                            isCurrentUser ? ", your ranking" : ""
                          }`}
                        >
                          {isCurrentUser && <span className={styles.youBadge}>You</span>}
                          <div className={styles.portraitStage}>
                            <BuddyAvatar
                              name={buddy.name}
                              email={buddy.email}
                              context={context}
                              rank={buddy.rank}
                            />
                            <Medal rank={buddy.rank} />
                          </div>

                          <div className={styles.posterPanel}>
                            <div className={styles.posterPanelInner}>
                              <div className={styles.cardFront}>
                                <div className={styles.buddyIdentity}>
                                  <h4 title={buddy.name}>{buddy.name}</h4>
                                  <p><MapPin size={11} /> {buddy.location}</p>
                                </div>

                                <div className={styles.posterMetrics}>
                                  {isSuperAdmin ? (
                                    <>
                                      <PosterMetric
                                        icon={<CalendarCheck2 size={22} />}
                                        value={`${buddy.onTimeRate}%`}
                                        label="On-Time"
                                      />
                                      <PosterMetric
                                        icon={<Activity size={22} />}
                                        value={`${buddy.completionRate}%`}
                                        label="Completion"
                                      />
                                    </>
                                  ) : (
                                    <>
                                      <PosterMetric
                                        icon={<Trophy size={22} />}
                                        value={`#${buddy.rank}`}
                                        label="Monthly Rank"
                                      />
                                      <PosterMetric
                                        icon={<CalendarCheck2 size={22} />}
                                        value={selectedMonth.toLocaleDateString("en-IN", { month: "short" })}
                                        label="Recognition Month"
                                      />
                                    </>
                                  )}
                                  <PosterMetric
                                    icon={<Star size={22} />}
                                        value={`${buddy.score} pts`}
                                        label="Performance Points"
                                  />
                                </div>
                              </div>

                              {isSuperAdmin && (
                                <div className={styles.cardBack}>
                                  <div className={styles.detailTitle}>
                                    <Info size={13} />
                                    <span>Points Breakdown</span>
                                    <strong>{buddy.score}<small>/100 points</small></strong>
                                  </div>
                                  <div className={styles.detailGrid}>
                                    <Metric icon={<Star size={13} />} label={`${buddy.ratingCount} ratings`} value={buddy.averageRating ? `${buddy.averageRating}/5` : "No rating"} />
                                    <Metric icon={<Users size={13} />} label="New Joiners" value={String(buddy.newJoiners)} />
                                    <Metric icon={<CalendarCheck2 size={13} />} label="Within ±5 days" value={`${buddy.onTime}/${Math.max(buddy.scheduled, buddy.completed)}`} />
                                    <Metric icon={<Activity size={13} />} label={`${buddy.completed}/${Math.max(buddy.scheduled, buddy.completed)} completed`} value={`${buddy.completionRate}%`} />
                                  </div>
                                  <p>
                                    Timeline {buddy.onTimePoints}/{SCORE_WEIGHTS.onTime} + Rating {buddy.ratingPoints}/{SCORE_WEIGHTS.rating} + Completion {buddy.completionPoints}/{SCORE_WEIGHTS.completion} + NJ support {buddy.newJoinerPoints}/{SCORE_WEIGHTS.newJoiners}
                                  </p>
                                  {buddy.missingPlannedDates > 0 && (
                                    <em>{buddy.missingPlannedDates} completed interaction(s) missing a planned date.</em>
                                  )}
                                </div>
                              )}
                            </div>
                          </div>
                        </div>
                      );
                    })}
                </div>
              </article>
            ))}
          </div>
        </div>
      )}
    </section>
  );
};

export default BuddyPerformance;