import { useCallback, useEffect, useRef, useState } from "react";

export function useMirroredList(onChange?: (next: string[]) => void) {
  const [items, setItems] = useState<string[]>([]);

  // Updates go through state; callbacks follow the committed render.
  const replace = useCallback(
    (next: string[]) => {
      setItems(next);
      onChange?.(next);
    },
    [onChange],
  );

  // A ref that is not a render mirror may be written anywhere.
  const timerRef = useRef<number | null>(null);
  useEffect(() => {
    timerRef.current = window.setTimeout(() => {
      timerRef.current = null;
    }, 0);
  }, []);

  const read = useCallback(() => items, [items]);

  return { items, replace, read };
}

export const useCounter = () => {
  const [count, setCount] = useState(0);

  const increment = () => setCount((current) => current + 1);

  return { count, increment };
};
