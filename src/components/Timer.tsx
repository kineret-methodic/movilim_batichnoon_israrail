import React, { useState, useEffect, useRef } from 'react';

interface Props {
  initialSeconds: number;
  stopLabel: string;
  resumeLabel: string;
}

export default function Timer({ initialSeconds, stopLabel, resumeLabel }: Props) {
  const [seconds, setSeconds] = useState(initialSeconds);
  const [running, setRunning] = useState(false);
  const ref = useRef<ReturnType<typeof setInterval> | null>(null);

  useEffect(() => {
    if (running && seconds > 0) {
      ref.current = setInterval(() => {
        setSeconds(s => {
          if (s <= 1) { setRunning(false); return 0; }
          return s - 1;
        });
      }, 1000);
    }
    return () => { if (ref.current) clearInterval(ref.current); };
  }, [running]);

  const reset = () => { setRunning(false); setSeconds(initialSeconds); };
  const toggle = () => { if (seconds === 0) reset(); else setRunning(r => !r); };

  const pct = seconds / initialSeconds;
  const timeColor = seconds === 0 ? '#CC2200' : pct < 0.25 ? '#E07820' : '#1A1C50';
  const mm = String(Math.floor(seconds / 60)).padStart(2, '0');
  const ss = String(seconds % 60).padStart(2, '0');

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
      <span style={{
        fontFamily: 'monospace', fontSize: 26, fontWeight: 700, color: timeColor,
        minWidth: 68, letterSpacing: 2,
      }}>
        {mm}:{ss}
      </span>
      <button onClick={toggle} style={btn('#1E1F56', 'white')}>
        {running ? stopLabel : resumeLabel}
      </button>
      <button onClick={reset} style={btn('#EEF0FF', '#1E1F56')} title="אפס">↺</button>
    </div>
  );
}

function btn(bg: string, color: string): React.CSSProperties {
  return {
    padding: '5px 12px', borderRadius: 6, border: '1px solid #C0C4E8',
    background: bg, color, fontSize: 13, fontWeight: 600,
  };
}
