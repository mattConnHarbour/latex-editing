export type EquationEditorPosition = { left: number; top: number };

export type ScrollSnapshot = {
  containers: Array<{ path: number[]; left: number; top: number }>;
  windowX: number;
  windowY: number;
};

const EDITOR_WIDTH = 440;
const EDITOR_ESTIMATED_HEIGHT = 470;

function clamp(value: number, minimum: number, maximum: number) {
  return Math.min(Math.max(value, minimum), Math.max(minimum, maximum));
}

function pathFromHost(host: HTMLElement, element: HTMLElement): number[] | null {
  const path: number[] = [];
  let current: HTMLElement | null = element;
  while (current && current !== host) {
    const parent: HTMLElement | null = current.parentElement;
    if (!parent) return null;
    path.unshift(Array.from(parent.children).indexOf(current));
    current = parent;
  }
  return current === host ? path : null;
}

function elementFromPath(host: HTMLElement, path: number[]): HTMLElement | null {
  let current: Element = host;
  for (const index of path) {
    const next = current.children[index];
    if (!(next instanceof HTMLElement)) return null;
    current = next;
  }
  return current instanceof HTMLElement ? current : null;
}

export function getEquationEditorPosition(bounds: DOMRect): EquationEditorPosition {
  const preferredLeft = bounds.right + 16;
  const fallbackLeft = bounds.left - EDITOR_WIDTH - 16;
  return {
    left: clamp(
      preferredLeft + EDITOR_WIDTH <= window.innerWidth - 12 ? preferredLeft : fallbackLeft,
      12,
      window.innerWidth - EDITOR_WIDTH - 12,
    ),
    top: clamp(bounds.top, 72, window.innerHeight - EDITOR_ESTIMATED_HEIGHT - 12),
  };
}

export function captureScrollSnapshot(host: HTMLElement, target: Element | null): ScrollSnapshot {
  const containers: ScrollSnapshot['containers'] = [];
  let current = target?.parentElement ?? null;
  while (current) {
    if (current.scrollTop !== 0 || current.scrollLeft !== 0) {
      const path = pathFromHost(host, current);
      if (path) containers.push({ path, left: current.scrollLeft, top: current.scrollTop });
    }
    if (current === host) break;
    current = current.parentElement;
  }
  return { containers, windowX: window.scrollX, windowY: window.scrollY };
}

export function restoreScrollSnapshot(host: HTMLElement, snapshot: ScrollSnapshot) {
  window.scrollTo(snapshot.windowX, snapshot.windowY);
  for (const container of snapshot.containers) {
    elementFromPath(host, container.path)?.scrollTo(container.left, container.top);
  }
}
