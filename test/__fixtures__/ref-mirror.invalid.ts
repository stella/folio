import { useCallback, useEffect, useRef, useState } from "react";

export function useMirroredList(onChange?: (next: string[]) => void) {
  const [items, setItems] = useState<string[]>([]);
  const itemsRef = useRef(items);
  itemsRef.current = items;

  const replace = useCallback(
    (next: string[]) => {
      itemsRef.current = next;
      onChange?.(next);
    },
    [onChange],
  );

  return { items, setItems, replace };
}

export const Counter = () => {
  const [count] = useState(0);
  const countRef = useRef(count);
  countRef.current = count;

  useEffect(() => {
    countRef.current = countRef.current + 1;
  }, []);

  return null;
};
