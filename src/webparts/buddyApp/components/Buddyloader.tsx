/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/self-closing-comp */

import * as React from "react";

export type LoaderVariant = "upload" | "submit" | "update" | "default";

export interface IBuddyLoaderProps {
  visible: boolean;
  variant?: LoaderVariant;
  message?: string;
  subMessage?: string;
}

interface HexConfig {
  w: number;
  h: number;
  top: string;
  left?: string;
  right?: string;
  dur: string;
  rev: boolean;
}

const VARIANT_CONFIG: Record<
  LoaderVariant,
  { icon: React.ReactElement; badge: string; defaultMsg: string; defaultSub: string; accentColor: string }
> = {
  upload: {
    accentColor: "#00D264",
    badge: "Uploading",
    defaultMsg: "Uploading your file…",
    defaultSub: "Please wait while we save your content.",
    icon: (
      <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
        <polyline points="16 16 12 12 8 16" />
        <line x1="12" y1="12" x2="12" y2="21" />
        <path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3" />
      </svg>
    ),
  },
  submit: {
    accentColor: "#00D264",
    badge: "Submitting",
    defaultMsg: "Submitting your details…",
    defaultSub: "Hang tight, we're processing your request.",
    icon: (
      <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
        <line x1="22" y1="2" x2="11" y2="13" />
        <polygon points="22 2 15 22 11 13 2 9 22 2" />
      </svg>
    ),
  },
  update: {
    accentColor: "#8CFF8C",
    badge: "Updating",
    defaultMsg: "Saving your changes…",
    defaultSub: "Your updates are being applied.",
    icon: (
      <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
        <polyline points="23 4 23 10 17 10" />
        <polyline points="1 20 1 14 7 14" />
        <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15" />
      </svg>
    ),
  },
  default: {
    accentColor: "#00D264",
    badge: "Processing",
    defaultMsg: "Please wait…",
    defaultSub: "This will only take a moment.",
    icon: (
      <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
        <circle cx="12" cy="12" r="10" />
        <polyline points="12 6 12 12 16 14" />
      </svg>
    ),
  },
};

export const BuddyLoader: React.FC<IBuddyLoaderProps> = ({
  visible,
  variant = "default",
  message,
  subMessage,
}) => {
  if (!visible) return null;

  const cfg = VARIANT_CONFIG[variant];
  const msg = message    ?? cfg.defaultMsg;
  const sub = subMessage ?? cfg.defaultSub;

  return (
    <>
      <style>{`
        @keyframes bldr-backdrop-in { from{opacity:0} to{opacity:1} }
        @keyframes bldr-card-in     { 0%{opacity:0;transform:scale(.78) translateY(28px)} 70%{transform:scale(1.03) translateY(-3px)} 100%{opacity:1;transform:scale(1) translateY(0)} }
        @keyframes bldr-float       { 0%,100%{transform:translateY(0)} 50%{transform:translateY(-10px)} }
        @keyframes bldr-spin-cw     { from{transform:rotate(0deg)}    to{transform:rotate(360deg)} }
        @keyframes bldr-spin-ccw    { from{transform:rotate(0deg)}    to{transform:rotate(-360deg)} }
        @keyframes bldr-ring-out    { 0%{transform:scale(.55);opacity:.75} 100%{transform:scale(2.2);opacity:0} }
        @keyframes bldr-dot-pulse   { 0%,80%,100%{transform:scale(.55);opacity:.35} 40%{transform:scale(1.2);opacity:1} }
        @keyframes bldr-bar         { 0%{width:0%} 40%{width:65%} 70%{width:82%} 90%{width:93%} 100%{width:98%} }
        @keyframes bldr-shimmer     { 0%{transform:translateX(-100%)} 100%{transform:translateX(220%)} }
        @keyframes bldr-spark       { 0%,100%{opacity:0;transform:scale(0) rotate(0deg)} 40%{opacity:1;transform:scale(1) rotate(180deg)} 70%{opacity:.55;transform:scale(.75) rotate(290deg)} }
        @keyframes bldr-hex-spin    { from{transform:rotate(45deg)} to{transform:rotate(405deg)} }
        @keyframes bldr-slide-up    { from{opacity:0;transform:translateY(14px)} to{opacity:1;transform:translateY(0)} }
        @keyframes bldr-orbit1      { from{transform:rotate(0deg)   translateX(48px)} to{transform:rotate(360deg)  translateX(48px)} }
        @keyframes bldr-orbit2      { from{transform:rotate(120deg) translateX(40px)} to{transform:rotate(480deg)  translateX(40px)} }
        @keyframes bldr-orbit3      { from{transform:rotate(240deg) translateX(34px)} to{transform:rotate(600deg)  translateX(34px)} }

        .bldr-hex  { position:absolute; border:1.5px solid rgba(0,210,100,0.13); border-radius:4px; }
        .bldr-spark{ position:absolute; animation:bldr-spark 2.2s ease-in-out infinite; color:#00D264; font-size:12px; pointer-events:none; }
        .bldr-od   { position:absolute; top:50%; left:50%; border-radius:50%; }
        .bldr-ring { position:absolute; border-radius:50%; border:1.5px solid rgba(0,210,100,0.18); animation:bldr-ring-out 2.6s ease-out infinite; }
        .bldr-bar-track { width:100%; height:5px; background:rgba(255,255,255,0.1); border-radius:999px; overflow:hidden; margin-top:1.375rem; }
        .bldr-bar-fill  { height:100%; border-radius:999px; background:linear-gradient(90deg,#00D264,#8CFF8C); animation:bldr-bar 4s cubic-bezier(.4,0,.2,1) forwards; position:relative; overflow:hidden; }
        .bldr-bar-shim  { position:absolute; inset:0; background:linear-gradient(90deg,transparent,rgba(255,255,255,0.35),transparent); transform:translateX(-100%); animation:bldr-shimmer 1.6s 0.5s infinite; }
      `}</style>

      <div
        style={{
          position: "fixed",
          inset: 0,
          background: "rgba(0,0,0,0.62)",
          zIndex: 99999,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          padding: "1.5rem",
          fontFamily: '"Inter",-apple-system,sans-serif',
          animation: "bldr-backdrop-in .25s ease both",
        }}
      >
        <div
          style={{
            background: "#004632",
            borderRadius: 22,
            width: "100%",
            maxWidth: 380,
            padding: "2rem 1.75rem 1.75rem",
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            position: "relative",
            overflow: "hidden",
            animation: "bldr-card-in .5s cubic-bezier(.22,1,.36,1) both",
            boxShadow: "0 32px 80px rgba(0,0,0,0.55)",
          }}
        >
          {([
            { w: 58, h: 40, top: "4%",  left: "4%",   dur: "14s", rev: false },
            { w: 74, h: 52, top: "3%",  right: "5%",  dur: "20s", rev: true  },
            { w: 44, h: 30, top: "55%", left: "2%",   dur: "16s", rev: false },
            { w: 62, h: 44, top: "58%", right: "3%",  dur: "22s", rev: true  },
          ] as HexConfig[]).map((hx, i) => (
            <div
              key={i}
              className="bldr-hex"
              style={{
                width: hx.w,
                height: hx.h,
                top: hx.top,
                left: hx.left,
                right: hx.right,
                animation: `bldr-hex-spin ${hx.dur} linear infinite ${hx.rev ? "reverse" : ""}`,
              }}
            />
          ))}

          {[
            ["11%","14%","0s"],["78%","9%",".35s"],["65%","73%",".65s"],
            ["18%","67%","1s"],["46%","5%","1.25s"],["86%","46%","1.55s"],
          ].map(([t, l, del], i) => (
            <div
              key={i}
              className="bldr-spark"
              style={{ top: t, left: l, animationDelay: del, animationDuration: `${2 + i * 0.18}s` }}
            >✦</div>
          ))}

          {[
            { s: "86px",  d: "0s"   },
            { s: "152px", d: ".85s" },
            { s: "228px", d: "1.7s" },
          ].map((r, i) => (
            <div
              key={i}
              className="bldr-ring"
              style={{
                width: r.s,
                height: r.s,
                top: `calc(30% - ${parseInt(r.s) / 2}px)`,
                left: `calc(50% - ${parseInt(r.s) / 2}px)`,
                animationDelay: r.d,
              }}
            />
          ))}

          <div
            style={{
              position: "relative",
              width: 100,
              height: 100,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              animation: "bldr-float 3s ease-in-out infinite",
              marginBottom: "1.25rem",
              zIndex: 5,
            }}
          >
            <div style={{ position: "absolute", inset: -16, borderRadius: "50%", background: "radial-gradient(circle,rgba(0,210,100,0.22) 0%,transparent 70%)" }} />
            <div style={{ position: "absolute", inset: -5, borderRadius: "50%", border: "2px dashed rgba(0,210,100,0.3)", animation: "bldr-spin-cw 10s linear infinite" }} />
            <div style={{ position: "absolute", inset: -12, borderRadius: "50%", border: "1.5px dashed rgba(140,255,140,0.16)", animation: "bldr-spin-ccw 16s linear infinite" }} />

            {[
              { bg: "#8CFF8C", s: "10px", anim: "bldr-orbit1", dur: "3.2s" },
              { bg: "#00D264", s:  "8px", anim: "bldr-orbit2", dur: "2.6s" },
              { bg: "rgba(255,255,255,.5)", s: "6px", anim: "bldr-orbit3", dur: "4s" },
            ].map((o, i) => (
              <div
                key={i}
                className="bldr-od"
                style={{
                  width: o.s,
                  height: o.s,
                  background: o.bg,
                  marginTop: `-${parseInt(o.s) / 2}px`,
                  marginLeft: `-${parseInt(o.s) / 2}px`,
                  animation: `${o.anim} ${o.dur} linear infinite`,
                }}
              />
            ))}

            <div
              style={{
                width: 92,
                height: 92,
                borderRadius: "50%",
                background: "linear-gradient(135deg,#00D264,#004632)",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                border: "3px solid rgba(255,255,255,0.2)",
                position: "relative",
                zIndex: 2,
                boxShadow: "0 0 0 7px rgba(0,210,100,0.12),0 0 0 16px rgba(0,210,100,0.06)",
              }}
            >
              {cfg.icon}
            </div>
          </div>

          <div
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
              background: "rgba(0,210,100,0.15)",
              border: "1px solid rgba(0,210,100,0.32)",
              borderRadius: 999,
              padding: "3px 13px",
              fontSize: 10,
              fontWeight: 600,
              color: "#8CFF8C",
              letterSpacing: ".07em",
              textTransform: "uppercase",
              marginBottom: "0.625rem",
              animation: "bldr-slide-up .4s .15s both",
            }}
          >
            <div
              style={{
                width: 5,
                height: 5,
                borderRadius: "50%",
                background: "#00D264",
                animation: "bldr-dot-pulse 1.3s infinite",
              }}
            />
            {cfg.badge}
          </div>

          <div
            style={{
              fontSize: 17,
              fontWeight: 700,
              color: "#fff",
              letterSpacing: "-.02em",
              textAlign: "center",
              marginBottom: "0.35rem",
              animation: "bldr-slide-up .4s .25s both",
              zIndex: 5,
            }}
          >
            {msg}
          </div>

          <p
            style={{
              fontSize: 12,
              color: "rgba(255,255,255,0.55)",
              textAlign: "center",
              lineHeight: 1.7,
              maxWidth: 280,
              margin: 0,
              animation: "bldr-slide-up .4s .33s both",
              zIndex: 5,
            }}
          >
            {sub}
          </p>

          <div className="bldr-bar-track" style={{ animation: "bldr-slide-up .4s .42s both" }}>
            <div className="bldr-bar-fill">
              <div className="bldr-bar-shim" />
            </div>
          </div>

          <div
            style={{
              display: "flex",
              gap: 7,
              marginTop: "1rem",
              animation: "bldr-slide-up .4s .5s both",
            }}
          >
            {[0, 1, 2].map(i => (
              <div
                key={i}
                style={{
                  width: 8,
                  height: 8,
                  borderRadius: "50%",
                  background: i === 0 ? "#00D264" : i === 1 ? "#8CFF8C" : "rgba(255,255,255,0.3)",
                  animation: `bldr-dot-pulse 1.3s ${i * 0.2}s infinite`,
                }}
              />
            ))}
          </div>
        </div>
      </div>
    </>
  );
};

export default BuddyLoader;