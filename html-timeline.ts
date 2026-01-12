/**
 * HTML Timeline Generator
 *
 * A TypeScript module that generates HTML table-based timelines using
 * vis-timeline compatible data structures. Outputs pure HTML/CSS that
 * displays well on web and exports cleanly to PowerPoint via pptxgenjs.
 */

// ============================================================================
// TYPES AND INTERFACES
// ============================================================================

/** Item type enumeration - matches vis-timeline types */
export type TimelineItemType = 'box' | 'point' | 'range' | 'background';

/** Time granularity levels */
export type TimeGranularity = 'hour' | 'day' | 'week' | 'month' | 'quarter' | 'year';

/**
 * Timeline item - vis-timeline compatible structure
 */
export interface TimelineItem {
  /** Unique identifier for the item */
  id: string | number;
  /** Display content (text or HTML) */
  content: string;
  /** Start date/time of the item */
  start: Date | string | number;
  /** End date/time (optional for point items) */
  end?: Date | string | number;
  /** Item display type */
  type?: TimelineItemType;
  /** Group ID this item belongs to */
  group?: string | number;
  /** CSS class name(s) for styling */
  className?: string;
  /** Custom inline styles */
  style?: string;
  /** Title/tooltip text */
  title?: string;
}

/**
 * Timeline group - vis-timeline compatible structure
 */
export interface TimelineGroup {
  /** Unique identifier for the group */
  id: string | number;
  /** Display content/label for the group */
  content: string;
  /** Value for custom ordering */
  value?: number;
  /** CSS class name(s) for styling */
  className?: string;
  /** Custom inline styles */
  style?: string;
}

/**
 * Timeline configuration options
 */
export interface TimelineOptions {
  /** Time range start (auto-detected if not set) */
  start?: Date | string;
  /** Time range end (auto-detected if not set) */
  end?: Date | string;
  /** Time granularity ('auto' for automatic selection) */
  granularity?: TimeGranularity | 'auto';
  /** Show group labels column */
  showGroupLabels?: boolean;
  /** Group label column width */
  groupLabelWidth?: string;
  /** Whether to stack overlapping items in same group */
  stackItems?: boolean;
  /** Max stack depth per group (rows within group) */
  maxStackDepth?: number;
  /** CSS class prefix for all generated classes */
  classPrefix?: string;
  /** Locale for date formatting */
  locale?: string;
  /** Week start day (0=Sunday, 1=Monday) */
  weekStartDay?: 0 | 1 | 2 | 3 | 4 | 5 | 6;
  /** Group ordering: 'value' or 'content' */
  groupOrder?: 'value' | 'content' | 'id';
  /**
   * Enable compact stacking to fit more items on the same row.
   * When true, items only avoid overlapping their actual bars/dots,
   * allowing text to flow over other items' text areas.
   * When false (default), items reserve space for their text labels.
   */
  compactStacking?: boolean;
  /**
   * Container width in pixels for accurate text width calculations.
   * When provided, stacking uses actual pixel measurements instead of
   * character-based estimates. This makes compact stacking responsive
   * to different screen sizes.
   */
  containerWidth?: number;
}

// Internal types

interface NormalizedItem extends Omit<TimelineItem, 'start' | 'end'> {
  start: Date;
  end: Date;
  type: TimelineItemType;
  isPoint: boolean;
}

interface NormalizedGroup extends TimelineGroup {
  sortOrder: number;
  items: NormalizedItem[];
}

interface TimeCell {
  start: Date;
  end: Date;
  label: string;
  colIndex: number;
}

interface ItemSpan {
  item: NormalizedItem;
  startCol: number;
  endCol: number;
  startOffset: number;
  endOffset: number;
  stackRow: number;
}

interface ResolvedOptions {
  start: Date;
  end: Date;
  granularity: TimeGranularity;
  showGroupLabels: boolean;
  groupLabelWidth: string;
  groupLabelWidthPx: number;
  stackItems: boolean;
  maxStackDepth: number;
  classPrefix: string;
  locale: string;
  weekStartDay: number;
  groupOrder: 'value' | 'content' | 'id';
  compactStacking: boolean;
  containerWidth: number | null;
}

// ============================================================================
// DATE UTILITIES
// ============================================================================

function normalizeDate(date: Date | string | number): Date {
  if (date instanceof Date) return new Date(date);
  return new Date(date);
}

function daysBetween(start: Date, end: Date): number {
  const MS_PER_DAY = 86400000;
  return (end.getTime() - start.getTime()) / MS_PER_DAY;
}

function alignToGranularity(date: Date, granularity: TimeGranularity, weekStartDay: number = 0): Date {
  const aligned = new Date(date);

  switch (granularity) {
    case 'hour':
      aligned.setMinutes(0, 0, 0);
      break;
    case 'day':
      aligned.setHours(0, 0, 0, 0);
      break;
    case 'week':
      aligned.setHours(0, 0, 0, 0);
      const day = aligned.getDay();
      // Calculate days to subtract to reach the week start day
      const daysToSubtract = (day - weekStartDay + 7) % 7;
      aligned.setDate(aligned.getDate() - daysToSubtract);
      break;
    case 'month':
      aligned.setHours(0, 0, 0, 0);
      aligned.setDate(1);
      break;
    case 'quarter':
      aligned.setHours(0, 0, 0, 0);
      aligned.setDate(1);
      aligned.setMonth(Math.floor(aligned.getMonth() / 3) * 3);
      break;
    case 'year':
      aligned.setHours(0, 0, 0, 0);
      aligned.setMonth(0, 1);
      break;
  }

  return aligned;
}

function advanceByGranularity(date: Date, granularity: TimeGranularity): Date {
  const next = new Date(date);

  switch (granularity) {
    case 'hour':
      next.setHours(next.getHours() + 1);
      break;
    case 'day':
      next.setDate(next.getDate() + 1);
      break;
    case 'week':
      next.setDate(next.getDate() + 7);
      break;
    case 'month':
      next.setMonth(next.getMonth() + 1);
      break;
    case 'quarter':
      next.setMonth(next.getMonth() + 3);
      break;
    case 'year':
      next.setFullYear(next.getFullYear() + 1);
      break;
  }

  return next;
}

function formatDateLabel(date: Date, granularity: TimeGranularity, locale: string): string {
  switch (granularity) {
    case 'hour':
      return date.toLocaleString(locale, { hour: 'numeric', hour12: true });
    case 'day':
      return date.toLocaleDateString(locale, { day: 'numeric', month: 'short' });
    case 'week':
      const weekNum = getWeekNumber(date);
      return `W${weekNum}`;
    case 'month':
      return date.toLocaleDateString(locale, { month: 'short' });
    case 'quarter':
      const q = Math.floor(date.getMonth() / 3) + 1;
      return `Q${q}`;
    case 'year':
      return date.getFullYear().toString();
  }
}

function getWeekNumber(date: Date): number {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
}

// ============================================================================
// GRANULARITY SELECTION
// ============================================================================

function selectGranularity(start: Date, end: Date): TimeGranularity {
  const days = daysBetween(start, end);

  if (days <= 3) return 'hour';
  if (days <= 60) return 'day';
  if (days <= 180) return 'week';
  if (days <= 730) return 'month';
  if (days <= 1825) return 'quarter';
  return 'year';
}

// ============================================================================
// NORMALIZATION
// ============================================================================

function normalizeItem(item: TimelineItem): NormalizedItem {
  const start = normalizeDate(item.start);
  let end: Date;
  let isPoint = false;

  if (item.end) {
    end = normalizeDate(item.end);
  } else if (item.type === 'point') {
    end = new Date(start);
    isPoint = true;
  } else {
    // Default: same day
    end = new Date(start);
    isPoint = !item.end;
  }

  const type: TimelineItemType = item.type || (isPoint ? 'point' : 'range');

  return {
    ...item,
    start,
    end,
    type,
    isPoint,
  };
}

function normalizeGroups(
  groups: TimelineGroup[],
  items: NormalizedItem[],
  groupOrder: 'value' | 'content' | 'id'
): NormalizedGroup[] {
  const itemsByGroup = new Map<string | number, NormalizedItem[]>();

  for (const item of items) {
    const groupId = item.group ?? 0;
    if (!itemsByGroup.has(groupId)) {
      itemsByGroup.set(groupId, []);
    }
    itemsByGroup.get(groupId)!.push(item);
  }

  const normalized: NormalizedGroup[] = groups.map((group, index) => ({
    ...group,
    sortOrder: groupOrder === 'value' ? (group.value ?? index) : index,
    items: itemsByGroup.get(group.id) || [],
  }));

  // Sort groups
  normalized.sort((a, b) => {
    if (groupOrder === 'value') {
      return (a.value ?? 0) - (b.value ?? 0);
    } else if (groupOrder === 'content') {
      return a.content.localeCompare(b.content);
    }
    return a.sortOrder - b.sortOrder;
  });

  return normalized;
}

// ============================================================================
// TIME AXIS GENERATION
// ============================================================================

function generateTimeCells(
  start: Date,
  end: Date,
  granularity: TimeGranularity,
  locale: string,
  weekStartDay: number = 0
): TimeCell[] {
  const cells: TimeCell[] = [];
  let current = alignToGranularity(start, granularity, weekStartDay);
  let colIndex = 0;

  while (current < end) {
    const cellEnd = advanceByGranularity(current, granularity);

    cells.push({
      start: new Date(current),
      end: new Date(Math.min(cellEnd.getTime(), end.getTime())),
      label: formatDateLabel(current, granularity, locale),
      colIndex,
    });

    current = cellEnd;
    colIndex++;
  }

  return cells;
}

// ============================================================================
// ITEM-TO-CELL MAPPING
// ============================================================================

function findCellIndex(date: Date, timeCells: TimeCell[], inclusive: boolean = false): number {
  for (let i = 0; i < timeCells.length; i++) {
    const cell = timeCells[i];
    if (inclusive) {
      if (date >= cell.start && date <= cell.end) return i;
    } else {
      if (date >= cell.start && date < cell.end) return i;
    }
  }

  if (timeCells.length > 0) {
    if (date < timeCells[0].start) return 0;
    if (date >= timeCells[timeCells.length - 1].end) return timeCells.length - 1;
  }

  return -1;
}

function calculateOffsetPercent(date: Date, cellStart: Date, cellEnd: Date): number {
  const cellDuration = cellEnd.getTime() - cellStart.getTime();
  if (cellDuration === 0) return 0;

  const offset = date.getTime() - cellStart.getTime();
  const percent = (offset / cellDuration) * 100;

  return Math.max(0, Math.min(100, percent));
}

function calculateItemSpan(item: NormalizedItem, timeCells: TimeCell[]): ItemSpan | null {
  const startCellIndex = findCellIndex(item.start, timeCells);
  if (startCellIndex === -1) return null;

  let endCellIndex = findCellIndex(item.end, timeCells, true);
  if (endCellIndex === -1) return null;

  // For point items or when end date equals start date, ensure endCol >= startCol
  if (item.isPoint || endCellIndex < startCellIndex) {
    endCellIndex = startCellIndex;
  }

  const startCell = timeCells[startCellIndex];
  const endCell = timeCells[endCellIndex];

  const startOffset = calculateOffsetPercent(item.start, startCell.start, startCell.end);
  const endOffset = calculateOffsetPercent(item.end, endCell.start, endCell.end);

  return {
    item,
    startCol: startCellIndex,
    endCol: endCellIndex,
    startOffset,
    endOffset,
    stackRow: 0,
  };
}

// ============================================================================
// ITEM STACKING
// ============================================================================

// Average character width in pixels for 11px font
const AVG_CHAR_WIDTH_PX = 6;
// Default total timeline width assumption when containerWidth not provided
const DEFAULT_TIMELINE_WIDTH_PX = 800;

function estimateTextWidthPx(content: string): number {
  // Estimate text width in pixels based on character count
  // Using average character width for 11px sans-serif font
  return content.length * AVG_CHAR_WIDTH_PX;
}

/**
 * Calculate text extent as percentage of total timeline width.
 * This makes stacking consistent across different granularities.
 */
function estimateTextPercent(content: string, timelineWidthPx: number, compact: boolean = false): number {
  const textWidthPx = estimateTextWidthPx(content);

  if (compact) {
    // Compact mode: use actual text width with minimal gap
    const gapPx = 12;
    return ((textWidthPx + gapPx) / timelineWidthPx) * 100;
  }

  // Normal mode: text width plus generous padding (30px extra)
  const paddingPx = 30;
  return ((textWidthPx + paddingPx) / timelineWidthPx) * 100;
}

/**
 * Check if two item spans overlap, considering text that extends beyond bars.
 * Uses percentage-based positions for consistency across granularities.
 */
function spansOverlap(
  a: ItemSpan,
  b: ItemSpan,
  compactStacking: boolean = false,
  timelineWidthPx: number = DEFAULT_TIMELINE_WIDTH_PX,
  numCells: number = 10
): boolean {
  // Convert cell positions to percentage of timeline (0-100)
  const toPercent = (col: number, offset: number) => ((col + offset / 100) / numCells) * 100;

  const aStart = toPercent(a.startCol, a.startOffset);
  const bStart = toPercent(b.startCol, b.startOffset);

  // Ensure a starts before or at the same position as b
  if (aStart > bStart) {
    [a, b] = [b, a];
  }

  // Recalculate after potential swap
  const aStartPercent = toPercent(a.startCol, a.startOffset);
  const aEndPercent = toPercent(a.endCol, a.endOffset);
  const bStartPercent = toPercent(b.startCol, b.startOffset);

  // Calculate text extent as percentage of timeline
  const aTextPercent = estimateTextPercent(a.item.content, timelineWidthPx, compactStacking);

  // For box/range items: text is inside the bar or flows outside to the right
  // For point items: text flows to the right of the dot
  let aEffectiveEnd: number;
  if (a.item.isPoint) {
    // Point: dot position + text extends to the right
    aEffectiveEnd = aStartPercent + aTextPercent;
  } else {
    // Box/range item: check if text fits inside the bar
    const barWidthPercent = aEndPercent - aStartPercent;
    const textFitsInBar = barWidthPercent >= aTextPercent;

    if (compactStacking && textFitsInBar) {
      // Compact mode with text inside bar: just use bar end
      aEffectiveEnd = aEndPercent;
    } else {
      // Text flows outside the bar (either narrow bar or normal mode)
      // Account for text extending from the bar's end (or start if very narrow)
      const textStartPercent = textFitsInBar ? aEndPercent : aStartPercent;
      aEffectiveEnd = textStartPercent + aTextPercent;
    }
  }

  // Minimum gap between items for visual clarity
  const minGapPercent = (8 / timelineWidthPx) * 100; // 8px minimum gap

  // Items overlap if a's effective end reaches b's start
  return aEffectiveEnd + minGapPercent >= bStartPercent;
}

function stackItems(
  spans: ItemSpan[],
  maxDepth: number,
  compactStacking: boolean = false,
  timelineWidthPx: number = DEFAULT_TIMELINE_WIDTH_PX,
  numCells: number = 10
): ItemSpan[][] {
  if (spans.length === 0) return [];

  // Sort by start position (column + offset), then by duration (longer items first)
  const sorted = [...spans].sort((a, b) => {
    if (a.startCol !== b.startCol) return a.startCol - b.startCol;
    if (a.startOffset !== b.startOffset) return a.startOffset - b.startOffset;
    // Longer items first
    const aDuration = (a.endCol - a.startCol) * 100 + a.endOffset - a.startOffset;
    const bDuration = (b.endCol - b.startCol) * 100 + b.endOffset - b.startOffset;
    return bDuration - aDuration;
  });

  const stacks: ItemSpan[][] = [];

  // Normal mode: each item gets its own row (no horizontal stacking)
  // Compact mode: items can share rows if they don't overlap
  if (!compactStacking) {
    for (const span of sorted) {
      if (stacks.length < maxDepth) {
        span.stackRow = stacks.length;
        stacks.push([span]);
      }
    }
    return stacks;
  }

  // Compact stacking: try to fit items on existing rows
  for (const span of sorted) {
    let placed = false;

    for (let row = 0; row < Math.min(stacks.length, maxDepth); row++) {
      const canPlace = stacks[row].every(existing => !spansOverlap(span, existing, compactStacking, timelineWidthPx, numCells));

      if (canPlace) {
        span.stackRow = row;
        stacks[row].push(span);
        placed = true;
        break;
      }
    }

    if (!placed && stacks.length < maxDepth) {
      span.stackRow = stacks.length;
      stacks.push([span]);
    }
  }

  return stacks;
}

// ============================================================================
// HTML RENDERING
// ============================================================================

function cls(prefix: string, ...parts: string[]): string {
  return parts.map(p => `${prefix}-${p}`).join(' ');
}

function escapeHtml(text: string): string {
  const div = typeof document !== 'undefined'
    ? document.createElement('div')
    : null;
  if (div) {
    div.textContent = text;
    return div.innerHTML;
  }
  // Fallback for Node.js
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function escapeAttr(value: string | number): string {
  // Escape for use in HTML attributes (handles quotes and special chars)
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function renderItemBar(
  span: ItemSpan,
  timeCells: TimeCell[],
  stackDepth: number,
  prefix: string
): string {
  const { item } = span;
  const classes = [cls(prefix, 'item'), cls(prefix, `item--${item.type}`)];

  if (item.className) {
    classes.push(item.className);
  }

  // Calculate position and width
  const totalCols = span.endCol - span.startCol + 1;
  let left: number;
  let width: number;

  if (item.isPoint) {
    // Point items position at their actual date within the cell
    left = span.startOffset;
    width = 100 - span.startOffset; // Extend to end of cell (text flows beyond)
  } else {
    // Range items: calculate based on span
    if (totalCols === 1) {
      left = span.startOffset;
      width = Math.max(span.endOffset - span.startOffset, 5);
    } else {
      // Multi-cell span: start offset in first cell
      left = span.startOffset;
      // Width extends across cells
      width = (100 - span.startOffset) + ((totalCols - 2) * 100) + span.endOffset;
    }
  }

  // Calculate vertical position based on stack row
  // Each stack row gets a fixed height slot
  const itemHeight = 24; // Fixed pixel height per item slot
  const top = span.stackRow * (itemHeight + 4) + 4; // 4px gap between items

  const styles: string[] = [
    `left: ${left.toFixed(1)}%`,
    `width: ${width.toFixed(1)}%`,
    `top: ${top}px`,
  ];

  // Point items use fixed size, range items use calculated height
  if (!item.isPoint) {
    styles.push(`height: ${itemHeight}px`);
  }

  if (item.style) {
    // Escape quotes to prevent attribute breakout
    styles.push(escapeAttr(item.style));
  }

  // Add title attribute - for point items, always show title since content is hidden
  const titleText = item.title || (item.isPoint ? item.content : '');
  const titleAttr = titleText ? ` title="${escapeAttr(titleText)}"` : '';

  // Escape class names and item ID to prevent XSS
  const safeClasses = classes.map(c => escapeAttr(c)).join(' ');
  const safeId = escapeAttr(item.id);

  return `<div class="${safeClasses}" style="${styles.join('; ')}" data-item-id="${safeId}"${titleAttr}>
    <span class="${cls(prefix, 'item-content')}">${escapeHtml(item.content)}</span>
  </div>`;
}

function renderCell(
  spans: ItemSpan[],
  colIndex: number,
  timeCells: TimeCell[],
  stackDepth: number,
  prefix: string
): string {
  // Filter spans that start in this column
  const cellSpans = spans.filter(s => s.startCol === colIndex);

  let content = '';
  for (const span of cellSpans) {
    content += renderItemBar(span, timeCells, stackDepth, prefix);
  }

  return `<td class="${cls(prefix, 'cell')}">
    <div class="${cls(prefix, 'cell-content')}">${content}</div>
  </td>`;
}

function renderGroupRow(
  group: NormalizedGroup,
  spans: ItemSpan[],
  timeCells: TimeCell[],
  options: ResolvedOptions,
  rowIndex: number,
  timelineWidthPx: number
): string {
  const { classPrefix: prefix, showGroupLabels, stackItems: doStack, maxStackDepth, compactStacking } = options;
  const numCells = timeCells.length;

  const rowClass = `${cls(prefix, 'row')} ${cls(prefix, rowIndex % 2 === 0 ? 'row--even' : 'row--odd')}`;
  const groupClass = group.className ? ` ${escapeAttr(group.className)}` : '';

  // Stack items for this group
  const groupSpans = spans.filter(s => s.item.group === group.id);
  const stacks = doStack ? stackItems(groupSpans, maxStackDepth, compactStacking, timelineWidthPx, numCells) : [groupSpans];
  const stackDepth = stacks.length || 1;

  // Flatten stacks back to array with updated stackRow
  const allSpans = stacks.flat();

  let cells = '';

  // Group label cell
  if (showGroupLabels) {
    const labelStyle = group.style ? ` style="${escapeAttr(group.style)}"` : '';
    cells += `<td class="${cls(prefix, 'group-label')}${groupClass}"${labelStyle}>${escapeHtml(group.content)}</td>`;
  }

  // Data cells
  for (let i = 0; i < timeCells.length; i++) {
    cells += renderCell(allSpans, i, timeCells, stackDepth, prefix);
  }

  // Calculate row height: 28px per stack row (24px item + 4px gap) + 8px padding
  const rowHeight = stackDepth * 28 + 8;
  const rowStyle = ` style="height: ${rowHeight}px"`;

  return `<tr class="${rowClass}" data-group-id="${escapeAttr(group.id)}"${rowStyle}>${cells}</tr>`;
}

function renderHeader(
  timeCells: TimeCell[],
  options: ResolvedOptions
): string {
  const { classPrefix: prefix, showGroupLabels, groupLabelWidth } = options;

  let cells = '';

  if (showGroupLabels) {
    const widthStyle = groupLabelWidth ? ` style="width: ${groupLabelWidth}"` : '';
    cells += `<th class="${cls(prefix, 'group-label', 'header-label')}"${widthStyle}></th>`;
  }

  for (const cell of timeCells) {
    cells += `<th class="${cls(prefix, 'header-cell')}">${escapeHtml(cell.label)}</th>`;
  }

  return `<tr class="${cls(prefix, 'header')}">${cells}</tr>`;
}

function renderTable(
  groups: NormalizedGroup[],
  allSpans: ItemSpan[],
  timeCells: TimeCell[],
  options: ResolvedOptions
): string {
  const { classPrefix: prefix, containerWidth, groupLabelWidthPx, showGroupLabels } = options;

  // Calculate timeline width in pixels (excluding label column)
  // This is used for consistent stacking calculations across granularities
  let timelineWidthPx = DEFAULT_TIMELINE_WIDTH_PX;
  if (containerWidth && timeCells.length > 0) {
    const labelWidth = showGroupLabels ? groupLabelWidthPx : 0;
    timelineWidthPx = containerWidth - labelWidth;
  }

  const header = renderHeader(timeCells, options);

  let rows = '';
  for (let i = 0; i < groups.length; i++) {
    const groupSpans = allSpans.filter(s => s.item.group === groups[i].id);
    rows += renderGroupRow(groups[i], groupSpans, timeCells, options, i, timelineWidthPx);
  }

  return `<table class="${cls(prefix, 'table')}">
  <thead>${header}</thead>
  <tbody>${rows}</tbody>
</table>`;
}

// ============================================================================
// MAIN RENDERING FUNCTIONS
// ============================================================================

function resolveOptions(
  items: NormalizedItem[],
  userOptions?: TimelineOptions
): ResolvedOptions {
  // Auto-detect date range from items
  let start: Date;
  let end: Date;

  if (userOptions?.start) {
    start = normalizeDate(userOptions.start);
  } else {
    start = items.reduce((min, item) =>
      item.start < min ? item.start : min, items[0]?.start || new Date());
  }

  if (userOptions?.end) {
    end = normalizeDate(userOptions.end);
  } else {
    end = items.reduce((max, item) =>
      item.end > max ? item.end : max, items[0]?.end || new Date());
  }

  // Select granularity first (needed for proper alignment)
  let granularity: TimeGranularity;
  if (!userOptions?.granularity || userOptions.granularity === 'auto') {
    granularity = selectGranularity(start, end);
  } else {
    granularity = userOptions.granularity;
  }

  // Align dates to granularity boundaries, then add padding only at the end
  // This prevents empty leading columns
  const weekStartDay = userOptions?.weekStartDay ?? 0;
  start = alignToGranularity(start, granularity, weekStartDay);

  // For end date, align and then advance to include the last item fully
  end = alignToGranularity(end, granularity, weekStartDay);
  end = advanceByGranularity(end, granularity);

  // Parse group label width to pixels (assume px if number, parse if string)
  const groupLabelWidth = userOptions?.groupLabelWidth ?? '180px';
  let groupLabelWidthPx = 180;
  if (typeof groupLabelWidth === 'string') {
    const match = groupLabelWidth.match(/^(\d+)/);
    if (match) {
      groupLabelWidthPx = parseInt(match[1], 10);
    }
  }

  return {
    start,
    end,
    granularity,
    showGroupLabels: userOptions?.showGroupLabels ?? true,
    groupLabelWidth,
    groupLabelWidthPx,
    stackItems: userOptions?.stackItems ?? true,
    maxStackDepth: userOptions?.maxStackDepth ?? 20,
    classPrefix: userOptions?.classPrefix ?? 'tl',
    locale: userOptions?.locale ?? 'en-US',
    weekStartDay: userOptions?.weekStartDay ?? 0,
    groupOrder: userOptions?.groupOrder ?? 'value',
    compactStacking: userOptions?.compactStacking ?? false,
    containerWidth: userOptions?.containerWidth ?? null,
  };
}

/**
 * Render timeline to HTML string
 */
export function renderTimelineToString(
  items: TimelineItem[],
  groups: TimelineGroup[],
  options?: TimelineOptions
): string {
  // Normalize items
  const normalizedItems = items.map(normalizeItem);

  // Resolve options
  const resolved = resolveOptions(normalizedItems, options);

  // Normalize groups
  const normalizedGroups = normalizeGroups(groups, normalizedItems, resolved.groupOrder);

  // Generate time cells
  const timeCells = generateTimeCells(
    resolved.start,
    resolved.end,
    resolved.granularity,
    resolved.locale,
    resolved.weekStartDay
  );

  // Calculate item spans
  const allSpans: ItemSpan[] = [];
  for (const item of normalizedItems) {
    const span = calculateItemSpan(item, timeCells);
    if (span) {
      allSpans.push(span);
    }
  }

  // Render table
  return renderTable(normalizedGroups, allSpans, timeCells, resolved);
}

/**
 * Render timeline to DOM element (browser only)
 */
export function renderTimeline(
  items: TimelineItem[],
  groups: TimelineGroup[],
  options?: TimelineOptions
): HTMLTableElement {
  const html = renderTimelineToString(items, groups, options);

  const container = document.createElement('div');
  container.innerHTML = html;

  const table = container.querySelector('table');
  if (!table) {
    throw new Error('Failed to render timeline table');
  }
  return table;
}

/**
 * Fix overflowing text in box/range items (browser only)
 * Call this after rendering to move text outside boxes that are too small.
 * @param container - The container element holding the timeline, or document if not specified
 * @param prefix - CSS class prefix (default: 'tl')
 */
export function fixOverflowingText(container?: Element, prefix: string = 'tl'): void {
  const root = container || document;
  const boxItems = root.querySelectorAll(`.${prefix}-item--range, .${prefix}-item--box`);

  boxItems.forEach(item => {
    const content = item.querySelector(`.${prefix}-item-content`) as HTMLElement;
    if (content) {
      // First, ensure text is positioned inside to measure correctly
      item.classList.remove(`${prefix}-item--text-outside`);

      // Force synchronous layout recalculation by reading offsetWidth.
      // This ensures the browser computes the new layout before we measure scrollWidth.
      // The void operator discards the value to satisfy linters.
      void content.offsetWidth;

      // Check if text is truncated (content wider than box allows)
      const isTruncated = content.scrollWidth > content.clientWidth + 1;

      if (isTruncated) {
        item.classList.add(`${prefix}-item--text-outside`);
      }
    }
  });
}

/**
 * Get default CSS styles for the timeline
 */
export function getTimelineStyles(prefix: string = 'tl'): string {
  return `
/* HTML Timeline Generator - Default Styles */

.${prefix}-table {
  border-collapse: collapse;
  width: 100%;
  table-layout: fixed;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
  font-size: 12px;
  border: 1px solid #d0d7de;
}

/* Header */
.${prefix}-header {
  background-color: #f6f8fa;
}

.${prefix}-header-cell {
  padding: 10px 4px;
  text-align: center;
  font-weight: 600;
  border: 1px solid #d0d7de;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  color: #24292f;
  font-size: 11px;
}

.${prefix}-header-label {
  text-align: left;
  padding-left: 12px;
  font-weight: 600;
  color: #57606a;
}

/* Group Label */
.${prefix}-group-label {
  padding: 8px 12px;
  font-weight: 500;
  text-align: left;
  border: 1px solid #d0d7de;
  background-color: #f6f8fa;
  vertical-align: middle;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  color: #24292f;
}

/* Rows */
.${prefix}-row {
  height: 40px;
}

.${prefix}-row--even {
  background-color: #ffffff;
}

.${prefix}-row--odd {
  background-color: #f6f8fa;
}

/* Cells */
.${prefix}-cell {
  position: relative;
  padding: 0;
  border: 1px solid #eaeef2;
  vertical-align: top;
  overflow: visible;
}

.${prefix}-cell-content {
  position: relative;
  width: 100%;
  height: 100%;
  min-height: 36px;
}

/* Items */
.${prefix}-item {
  position: absolute;
  border-radius: 4px;
  overflow: visible;
  cursor: pointer;
  transition: box-shadow 0.15s ease, transform 0.1s ease;
  z-index: 1;
  display: flex;
  align-items: center;
}

.${prefix}-item:hover {
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
  z-index: 10;
  transform: translateY(-1px);
}

/* Item Types */
.${prefix}-item--range,
.${prefix}-item--box {
  background: linear-gradient(135deg, #0969da 0%, #0550ae 100%);
  border: 1px solid #0550ae;
}

.${prefix}-item--point {
  height: 20px !important;
  background: transparent !important;
  border: none !important;
  display: flex;
  align-items: center;
  gap: 4px;
  overflow: visible !important;
  white-space: nowrap;
}

.${prefix}-item--point::before {
  content: '';
  width: 10px;
  height: 10px;
  min-width: 10px;
  border-radius: 50%;
  background-color: #cf222e;
  border: 2px solid #a40e26;
  flex-shrink: 0;
  box-sizing: border-box;
}

.${prefix}-item--point:hover {
  transform: none;
}

.${prefix}-item--point:hover::before {
  transform: scale(1.2);
}

.${prefix}-item--background {
  background-color: rgba(9, 105, 218, 0.1);
  border: none;
  border-radius: 0;
  z-index: 0;
}

/* Item Content */
.${prefix}-item-content {
  display: block;
  padding: 2px 8px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  color: #ffffff;
  font-size: 11px;
  font-weight: 500;
  line-height: 1.4;
}

/* Box/Range items where text doesn't fit: move text outside to the right */
.${prefix}-item--text-outside .${prefix}-item-content {
  position: absolute;
  left: 100%;
  top: 50%;
  transform: translateY(-50%);
  padding: 2px 6px;
  color: #24292f;
  overflow: visible;
}

.${prefix}-item--point .${prefix}-item-content {
  display: block;
  padding: 0;
  color: #24292f;
  font-size: 11px;
  font-weight: 500;
  white-space: nowrap;
  overflow: visible;
}

.${prefix}-item--background .${prefix}-item-content {
  color: #0969da;
}

/* ==========================================================================
   COLOR VARIATIONS
   These are sensible defaults that can be overridden in your stylesheet.
   Apply these classes via the className property on items.
   ========================================================================== */

/* Success - Green */
.${prefix}-item.success {
  background: linear-gradient(135deg, #1a7f37 0%, #116329 100%);
  border-color: #116329;
}
.${prefix}-item--point.success::before {
  background-color: #1a7f37;
  border-color: #116329;
}

/* Warning - Amber/Yellow */
.${prefix}-item.warning {
  background: linear-gradient(135deg, #bf8700 0%, #9a6700 100%);
  border-color: #9a6700;
}
.${prefix}-item--point.warning::before {
  background-color: #bf8700;
  border-color: #9a6700;
}

/* Danger - Red */
.${prefix}-item.danger {
  background: linear-gradient(135deg, #cf222e 0%, #a40e26 100%);
  border-color: #a40e26;
}
.${prefix}-item--point.danger::before {
  background-color: #cf222e;
  border-color: #a40e26;
}

/* Info - Blue (lighter than default) */
.${prefix}-item.info {
  background: linear-gradient(135deg, #0550ae 0%, #033d8b 100%);
  border-color: #033d8b;
}
.${prefix}-item--point.info::before {
  background-color: #0550ae;
  border-color: #033d8b;
}

/* Purple */
.${prefix}-item.purple {
  background: linear-gradient(135deg, #8250df 0%, #6639ba 100%);
  border-color: #6639ba;
}
.${prefix}-item--point.purple::before {
  background-color: #8250df;
  border-color: #6639ba;
}

/* Pink */
.${prefix}-item.pink {
  background: linear-gradient(135deg, #bf3989 0%, #99286e 100%);
  border-color: #99286e;
}
.${prefix}-item--point.pink::before {
  background-color: #bf3989;
  border-color: #99286e;
}
`;
}

// ============================================================================
// BROWSER HELPERS
// ============================================================================

/**
 * Options for mounting a timeline
 */
export interface MountOptions extends TimelineOptions {
  /** Debounce delay for resize handler in ms (default: 150) */
  resizeDebounce?: number;
  /** Disable automatic resize handling */
  disableResize?: boolean;
  /** Callback after each render */
  onRender?: () => void;
}

/**
 * Mounted timeline instance with cleanup
 */
export interface MountedTimeline {
  /** Re-render the timeline with current options */
  render: () => void;
  /** Update options and re-render */
  update: (newOptions: Partial<MountOptions>) => void;
  /** Clean up resize listeners */
  destroy: () => void;
  /** Get current options */
  getOptions: () => MountOptions;
}

/**
 * Inject timeline CSS styles into the document head
 * @param prefix - CSS class prefix (default: 'tl')
 * @param styleId - ID for the style element (default: 'timeline-styles')
 * @returns The style element
 */
export function injectStyles(prefix: string = 'tl', styleId: string = 'timeline-styles'): HTMLStyleElement {
  let styleEl = document.getElementById(styleId) as HTMLStyleElement;

  if (!styleEl) {
    styleEl = document.createElement('style');
    styleEl.id = styleId;
    document.head.appendChild(styleEl);
  }

  styleEl.textContent = getTimelineStyles(prefix);
  return styleEl;
}

/**
 * Create a debounced resize handler
 * @param callback - Function to call on resize
 * @param delay - Debounce delay in ms (default: 150)
 * @returns Object with attach/detach methods
 */
export function createResizeHandler(
  callback: () => void,
  delay: number = 150
): { attach: () => void; detach: () => void } {
  let timeout: ReturnType<typeof setTimeout> | null = null;

  const handler = () => {
    if (timeout) clearTimeout(timeout);
    timeout = setTimeout(callback, delay);
  };

  return {
    attach: () => window.addEventListener('resize', handler),
    detach: () => {
      if (timeout) clearTimeout(timeout);
      window.removeEventListener('resize', handler);
    },
  };
}

/**
 * Mount a timeline to a container element with full setup
 * Handles: style injection, rendering, text overflow fix, and resize handling
 *
 * @param container - Container element or selector
 * @param items - Timeline items
 * @param groups - Timeline groups
 * @param options - Mount options
 * @returns Mounted timeline instance with render/update/destroy methods
 *
 * @example
 * ```typescript
 * const timeline = mountTimeline('#timeline', items, groups, {
 *   granularity: 'month',
 *   compactStacking: true,
 * });
 *
 * // Update options later
 * timeline.update({ compactStacking: false });
 *
 * // Clean up when done
 * timeline.destroy();
 * ```
 */
export function mountTimeline(
  container: HTMLElement | string,
  items: TimelineItem[],
  groups: TimelineGroup[],
  options: MountOptions = {}
): MountedTimeline {
  // Resolve container
  const containerEl = typeof container === 'string'
    ? document.querySelector<HTMLElement>(container)
    : container;

  if (!containerEl) {
    throw new Error(`Timeline container not found: ${container}`);
  }

  // Inject styles
  const prefix = options.classPrefix ?? 'tl';
  injectStyles(prefix);

  // Current options
  let currentOptions: MountOptions = { ...options };

  // Render function
  const render = () => {
    // Auto-detect container width if not specified
    const containerWidth = currentOptions.containerWidth ?? containerEl.clientWidth;

    const html = renderTimelineToString(items, groups, {
      ...currentOptions,
      containerWidth,
    });

    containerEl.innerHTML = html;
    fixOverflowingText(containerEl, prefix);

    if (currentOptions.onRender) {
      currentOptions.onRender();
    }
  };

  // Set up resize handler
  const resizeHandler = createResizeHandler(render, currentOptions.resizeDebounce ?? 150);
  if (!currentOptions.disableResize) {
    resizeHandler.attach();
  }

  // Initial render
  render();

  // Return mounted instance
  return {
    render,
    update: (newOptions: Partial<MountOptions>) => {
      currentOptions = { ...currentOptions, ...newOptions };
      render();
    },
    destroy: () => {
      resizeHandler.detach();
      containerEl.innerHTML = '';
    },
    getOptions: () => ({ ...currentOptions }),
  };
}

// ============================================================================
// CONVENIENCE EXPORTS
// ============================================================================

export default {
  renderTimeline,
  renderTimelineToString,
  getTimelineStyles,
  fixOverflowingText,
  injectStyles,
  createResizeHandler,
  mountTimeline,
};
