import { useCallback, useEffect, useRef, useState } from "react";

export function useMirroredList(onChange?: (next: string[]) => void) {
  const [items, setItems] = useState<string[]>([]);
  const itemsRef = useRef(items);
  itemsRef.current = items;
  const onChangeRef = useRef(onChange);
  onChangeRef.current = onChange;

  // Mutation goes through state; the mirror follows on the next render.
  const replace = useCallback((next: string[]) => {
    setItems(next);
    onChangeRef.current?.(next);
  }, []);

  // A ref that is not a render mirror may be written anywhere.
  const timerRef = useRef<number | null>(null);
  useEffect(() => {
    timerRef.current = window.setTimeout(() => {
      timerRef.current = null;
    }, 0);
  }, []);

  const read = useCallback(() => itemsRef.current, []);

  return { items, replace, read };
}

export const Counter = () => {
  const [count, setCount] = useState(0);
  const countRef = useRef(count);
  countRef.current = count;

  useEffect(() => {
    setCount((current) => current + 1);
  }, []);

  return null;
};
