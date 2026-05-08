import { useEffect, useRef, useState } from 'react';

/**
 * Animates a number toward `value` using requestAnimationFrame.
 * Tuned for fast updates (slider drags) — short duration, stays in sync.
 */
export default function NumberTicker({ value, format, duration = 320, className }) {
  const [display, setDisplay] = useState(value);
  const fromRef = useRef(value);
  const startRef = useRef(null);
  const rafRef = useRef(null);

  useEffect(() => {
    if (typeof value !== 'number' || !isFinite(value)) {
      setDisplay(value);
      return;
    }
    fromRef.current = typeof display === 'number' && isFinite(display) ? display : value;
    startRef.current = null;

    const step = (ts) => {
      if (startRef.current === null) startRef.current = ts;
      const elapsed = ts - startRef.current;
      const t = Math.min(1, elapsed / duration);
      const eased = 1 - Math.pow(1 - t, 3);
      const next = fromRef.current + (value - fromRef.current) * eased;
      setDisplay(next);
      if (t < 1) rafRef.current = requestAnimationFrame(step);
    };

    rafRef.current = requestAnimationFrame(step);
    return () => cancelAnimationFrame(rafRef.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value]);

  return <span className={className}>{format ? format(display) : display}</span>;
}
