/* eslint-disable @typescript-eslint/explicit-function-return-type */
/*   eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import { useEffect, useState } from "react";
import { Crown, Sparkles, Trophy, X } from "lucide-react";

import { ICurrentBuddyRanking } from "./BuddyPerformance";
import styles from "./TopPerformerPopup.module.scss";

export interface ITopPerformerPopupProps {
  ranking: ICurrentBuddyRanking;
  context: any;
  photoUrl?: string;
  onClose: () => void;
}

const getInitials = (name?: string): string =>
  String(name || "Buddy")
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part.charAt(0).toUpperCase())
    .join("") || "B";

const TopPerformerPopup: React.FC<ITopPerformerPopupProps> = ({
  ranking,
  context,
  photoUrl,
  onClose,
}) => {
  const [photoFailed, setPhotoFailed] = useState(false);
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

  useEffect(() => setPhotoFailed(false), [resolvedPhotoUrl]);
  const rankThemeClass =
    ranking.rank === 1
      ? styles.gold
      : ranking.rank === 2
        ? styles.silver
        : styles.bronze;

  return (
    <div className={styles.backdrop} onClick={onClose} role="presentation">
      <section
        className={`${styles.card} ${rankThemeClass}`}
        onClick={(event) => event.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-labelledby="top-performer-title"
      >
        <button
          type="button"
          className={styles.closeButton}
          onClick={onClose}
          aria-label="Close Top Performer celebration"
        >
          <X size={18} />
        </button>

        <div className={styles.confetti} aria-hidden="true">
          {Array.from({ length: 14 }).map((_, index) => (
            <span key={index} />
          ))}
        </div>

        <div className={styles.eyebrow}>
          <Sparkles size={13} /> Top Performer · {ranking.monthLabel}
        </div>

        <div className={styles.heading}>
          <p>Congratulations!</p>
          <h2 id="top-performer-title">You’re a Top Performer</h2>
        </div>

        <div className={styles.stage}>
          <div className={styles.glow} />
          <Crown
            className={styles.crown}
            size={60}
            strokeWidth={1.8}
            aria-hidden="true"
          />
          <div className={styles.avatar}>
            <span>{getInitials(ranking.name)}</span>
            {resolvedPhotoUrl && !photoFailed && (
              <img
                src={resolvedPhotoUrl}
                alt={ranking.name}
                onError={() => setPhotoFailed(true)}
              />
            )}
          </div>
        </div>

        <div className={styles.rankBand} aria-label={`Monthly rank ${ranking.rank}`}>
          <span className={styles.rankDiamond}>
            <span className={styles.rankDiamondInner}>
              <Trophy size={16} />
              <strong>#{ranking.rank}</strong>
            </span>
          </span>
        </div>

        <div className={styles.identity}>
          <h3>{ranking.name}</h3>
          <p>
            Your timely support and consistent engagement placed you among
            this month’s leading Buddies.
          </p>
        </div>
      </section>
    </div>
  );
};

export default TopPerformerPopup;
