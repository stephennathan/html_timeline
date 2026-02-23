/**
 * HTML Timeline Generator
 *
 * A TypeScript module that generates HTML table-based timelines using
 * vis-timeline compatible data structures. Outputs pure HTML/CSS that
 * displays well on web and exports cleanly to PowerPoint via pptxgenjs.
 */
// ============================================================================
// DATE UTILITIES
// ============================================================================
function normalizeDate(date) {
  if (date instanceof Date) return new Date(date);
  return new Date(date);
}
function daysBetween(start, end) {
  const MS_PER_DAY = 86400000;
  return (end.getTime() - start.getTime()) / MS_PER_DAY;
}
function alignToGranularity(date, granularity, weekStartDay = 0) {
  const aligned = new Date(date);
  switch (granularity) {
    case 'hour': {
      aligned.setMinutes(0, 0, 0);
      break;
    }
    case 'day': {
      aligned.setHours(0, 0, 0, 0);
      break;
    }
    case 'week': {
      aligned.setHours(0, 0, 0, 0);
      const day = aligned.getDay();
      // Calculate days to subtract to reach the week start day
      const daysToSubtract = (day - weekStartDay + 7) % 7;
      aligned.setDate(aligned.getDate() - daysToSubtract);
      break;
    }
    case 'month': {
      aligned.setHours(0, 0, 0, 0);
      aligned.setDate(1);
      break;
    }
    case 'quarter': {
      aligned.setHours(0, 0, 0, 0);
      aligned.setDate(1);
      aligned.setMonth(Math.floor(aligned.getMonth() / 3) * 3);
      break;
    }
    case 'year': {
      aligned.setHours(0, 0, 0, 0);
      aligned.setMonth(0, 1);
      break;
    }
  }
  return aligned;
}
function advanceByGranularity(date, granularity) {
  const next = new Date(date);
  switch (granularity) {
    case 'hour': {
      next.setHours(next.getHours() + 1);
      break;
    }
    case 'day': {
      next.setDate(next.getDate() + 1);
      break;
    }
    case 'week': {
      next.setDate(next.getDate() + 7);
      break;
    }
    case 'month': {
      next.setMonth(next.getMonth() + 1);
      break;
    }
    case 'quarter': {
      next.setMonth(next.getMonth() + 3);
      break;
    }
    case 'year': {
      next.setFullYear(next.getFullYear() + 1);
      break;
    }
  }
  return next;
}
function formatDateLabel(date, granularity, locale) {
  switch (granularity) {
    case 'hour': {
      return date.toLocaleString(locale, { hour: 'numeric', hour12: true });
    }
    case 'day': {
      return date.toLocaleDateString(locale, { day: 'numeric', month: 'short' });
    }
    case 'week': {
      // Format as w/c d/mm (week commencing)
      const day = date.getDate();
      const month = date.getMonth() + 1;
      return `w/c ${day}/${month}`;
    }
    case 'month': {
      return date.toLocaleDateString(locale, { month: 'short' });
    }
    case 'quarter': {
      const q = Math.floor(date.getMonth() / 3) + 1;
      return `Q${q}`;
    }
    case 'year': {
      return date.getFullYear().toString();
    }
  }
}
// ============================================================================
// GRANULARITY SELECTION
// ============================================================================
function selectGranularity(start, end) {
  const days = daysBetween(start, end);
  if (days <= 3) return 'hour';
  if (days <= 14) return 'day'; // Up to 2 weeks: daily
  if (days <= 90) return 'week'; // Up to 3 months: weekly
  if (days <= 730) return 'month'; // Up to 2 years: monthly
  if (days <= 1825) return 'quarter';
  return 'year';
}
// ============================================================================
// NORMALIZATION
// ============================================================================
function normalizeItem(item) {
  const start = normalizeDate(item.start);
  let end;
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
  const type = item.type || (isPoint ? 'point' : 'range');
  return {
    ...item,
    start,
    end,
    type,
    isPoint,
  };
}
function normalizeGroups(groups, items, groupOrder) {
  const itemsByGroup = new Map();
  for (const item of items) {
    const groupId = item.group ?? 0;
    if (!itemsByGroup.has(groupId)) {
      itemsByGroup.set(groupId, []);
    }
    itemsByGroup.get(groupId).push(item);
  }
  const normalized = groups.map((group, index) => ({
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
function generateTimeCells(start, end, granularity, locale, weekStartDay = 0) {
  const cells = [];
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
/**
 * Generate header groups for two-row headers.
 * - Month/Quarter mode: groups by year
 * - Day mode: groups by week commencing
 */
function generateHeaderGroups(timeCells, granularity, weekStartDay) {
  const groups = [];
  if (granularity === 'month' || granularity === 'quarter') {
    // Group by year
    let currentYear = null;
    let currentGroup = null;
    for (const [index, cell] of timeCells.entries()) {
      const year = cell.start.getFullYear();
      if (year !== currentYear) {
        if (currentGroup) groups.push(currentGroup);
        currentGroup = { label: year.toString(), startIndex: index, span: 1 };
        currentYear = year;
      } else if (currentGroup) {
        currentGroup.span++;
      }
    }
    if (currentGroup) groups.push(currentGroup);
  } else if (granularity === 'day') {
    // Group by week commencing
    let currentWeekStart = null;
    let currentGroup = null;
    for (const [index, cell] of timeCells.entries()) {
      const weekStart = alignToGranularity(cell.start, 'week', weekStartDay);
      const weekStartTime = weekStart.getTime();
      if (weekStartTime !== currentWeekStart) {
        if (currentGroup) groups.push(currentGroup);
        // Format as w/c d/mm
        const day = weekStart.getDate();
        const month = weekStart.getMonth() + 1;
        currentGroup = { label: `w/c ${day}/${month}`, startIndex: index, span: 1 };
        currentWeekStart = weekStartTime;
      } else if (currentGroup) {
        currentGroup.span++;
      }
    }
    if (currentGroup) groups.push(currentGroup);
  }
  return groups;
}
// ============================================================================
// ITEM-TO-CELL MAPPING
// ============================================================================
function findCellIndex(date, timeCells, inclusive = false) {
  for (const [index, cell] of timeCells.entries()) {
    if (!cell) continue;
    if (inclusive) {
      if (date >= cell.start && date <= cell.end) return index;
    } else {
      if (date >= cell.start && date < cell.end) return index;
    }
  }
  if (timeCells.length > 0) {
    const firstCell = timeCells[0];
    const lastCell = timeCells.at(-1);
    if (firstCell && date < firstCell.start) return 0;
    if (lastCell && date >= lastCell.end) return timeCells.length - 1;
  }
  return -1;
}
function calculateOffsetPercent(date, cellStart, cellEnd) {
  const cellDuration = cellEnd.getTime() - cellStart.getTime();
  if (cellDuration === 0) return 0;
  const offset = date.getTime() - cellStart.getTime();
  const percent = (offset / cellDuration) * 100;
  return Math.max(0, Math.min(100, percent));
}
function calculateItemSpan(item, timeCells) {
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
  if (!startCell || !endCell) return null;
  const startOffset = calculateOffsetPercent(item.start, startCell.start, startCell.end);
  const endOffset = calculateOffsetPercent(item.end, endCell.start, endCell.end);
  return {
    item,
    startCol: startCellIndex,
    endCol: endCellIndex,
    startOffset,
    endOffset,
    stackRow: 0,
    labelLeft: false, // Will be determined later based on position
  };
}
// ============================================================================
// ITEM STACKING
// ============================================================================
// Average character width in pixels for 11px font
const AVG_CHAR_WIDTH_PX = 6;
// Default total timeline width assumption when containerWidth not provided
const DEFAULT_TIMELINE_WIDTH_PX = 800;
function estimateTextWidthPx(content) {
  // Estimate text width in pixels based on character count
  // Using average character width for 11px sans-serif font
  return content.length * AVG_CHAR_WIDTH_PX;
}
/**
 * Calculate text extent as percentage of total timeline width.
 * This makes stacking consistent across different granularities.
 */
function estimateTextPercent(content, timelineWidthPx, compact = false) {
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
 * Determine if a label should be positioned to the left of the bar/point.
 * Only applies to items that END at the last column of the timeline (i.e., at the edge).
 * This is a conservative check - we only move labels left for items truly at the end.
 */
function shouldLabelBeLeft(span, numberCells, timelineWidthPx, compactStacking, labelPosition) {
  // If labelPosition is 'right', always use right positioning (original behavior)
  if (labelPosition === 'right') {
    return false;
  }
  // Only consider items that end in the last column of the timeline
  const isAtTimelineEnd = span.endCol === numberCells - 1;
  if (!isAtTimelineEnd) {
    return false;
  }
  // For items at the end, check if the label would overflow
  const toPercent = (col, offset) => ((col + offset / 100) / numberCells) * 100;
  const startPercent = toPercent(span.startCol, span.startOffset);
  const endPercent = toPercent(span.endCol, span.endOffset);
  // Calculate text extent
  const textPercent = estimateTextPercent(span.item.content, timelineWidthPx, compactStacking);
  // Determine where the label would start (for right-positioned labels)
  let labelStartPercent;
  if (span.item.isPoint) {
    // Point: label starts just after the dot (roughly 1.5% for the dot width)
    labelStartPercent = startPercent + 1.5;
  } else {
    // Box/range: check if text fits inside the bar
    const barWidthPercent = endPercent - startPercent;
    const textFitsInBar = barWidthPercent >= textPercent;
    // If text fits inside the bar, it stays inside - no need for label-left
    if (textFitsInBar) {
      return false;
    }
    // Text doesn't fit inside, so it would be positioned outside starting at bar's start
    labelStartPercent = startPercent;
  }
  // Check if label would extend past 100% (end of timeline)
  const labelEndPercent = labelStartPercent + textPercent;
  return labelEndPercent > 100;
}
/**
 * Calculate the effective bounds of an item span including its label.
 * Returns [effectiveStart, effectiveEnd] as percentages.
 */
function getEffectiveBounds(span, numberCells, timelineWidthPx, compactStacking) {
  const toPercent = (col, offset) => ((col + offset / 100) / numberCells) * 100;
  const startPercent = toPercent(span.startCol, span.startOffset);
  const endPercent = toPercent(span.endCol, span.endOffset);
  const textPercent = estimateTextPercent(span.item.content, timelineWidthPx, compactStacking);
  if (span.labelLeft) {
    // Label is on the left - extends leftward from the start
    // Both point and box/range items extend left from the start position
    const effectiveStart = startPercent - textPercent;
    return [effectiveStart, endPercent];
  } else {
    // Label is on the right (default) - extends rightward
    let effectiveEnd;
    if (span.item.isPoint) {
      // Point: dot position + text extends to the right
      effectiveEnd = startPercent + textPercent;
    } else {
      // Box/range item: check if text fits inside the bar
      const barWidthPercent = endPercent - startPercent;
      const textFitsInBar = barWidthPercent >= textPercent;
      if (compactStacking && textFitsInBar) {
        // Compact mode with text inside bar: just use bar end
        effectiveEnd = endPercent;
      } else {
        // Text flows outside the bar
        const textStartPercent = textFitsInBar ? endPercent : startPercent;
        effectiveEnd = textStartPercent + textPercent;
      }
    }
    return [startPercent, effectiveEnd];
  }
}
/**
 * Check if two item spans overlap, considering text that extends beyond bars.
 * Uses percentage-based positions for consistency across granularities.
 */
function spansOverlap(
  a,
  b,
  compactStacking = false,
  timelineWidthPx = DEFAULT_TIMELINE_WIDTH_PX,
  numberCells = 10
) {
  const [aStart, aEnd] = getEffectiveBounds(a, numberCells, timelineWidthPx, compactStacking);
  const [bStart, bEnd] = getEffectiveBounds(b, numberCells, timelineWidthPx, compactStacking);
  // Minimum gap between items for visual clarity
  const minGapPercent = (8 / timelineWidthPx) * 100; // 8px minimum gap
  // Check for overlap: items overlap if their effective ranges intersect
  // Add minimum gap to prevent items from being too close
  return !(aEnd + minGapPercent <= bStart || bEnd + minGapPercent <= aStart);
}
function stackItems(
  spans,
  maxDepth,
  compactStacking = false,
  timelineWidthPx = DEFAULT_TIMELINE_WIDTH_PX,
  numberCells = 10,
  labelPosition = 'auto'
) {
  if (spans.length === 0) return [];
  // First, determine labelLeft for each span based on position
  for (const span of spans) {
    span.labelLeft = shouldLabelBeLeft(
      span,
      numberCells,
      timelineWidthPx,
      compactStacking,
      labelPosition
    );
  }
  // Sort by order field first (lower order = earlier in stack), then by start position, then by duration
  const sorted = [...spans].sort((a, b) => {
    // First sort by order field (lower order = earlier in stack)
    const orderA = a.item.order ?? Infinity;
    const orderB = b.item.order ?? Infinity;
    if (orderA !== orderB) return orderA - orderB;
    // Then by start position
    if (a.startCol !== b.startCol) return a.startCol - b.startCol;
    if (a.startOffset !== b.startOffset) return a.startOffset - b.startOffset;
    // Longer items first
    const aDuration = (a.endCol - a.startCol) * 100 + a.endOffset - a.startOffset;
    const bDuration = (b.endCol - b.startCol) * 100 + b.endOffset - b.startOffset;
    return bDuration - aDuration;
  });
  const stacks = [];
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
      const stackRow = stacks[row];
      if (!stackRow) continue;
      const canPlace = stackRow.every(
        (existing) => !spansOverlap(span, existing, compactStacking, timelineWidthPx, numberCells)
      );
      if (canPlace) {
        span.stackRow = row;
        stackRow.push(span);
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
function cls(prefix, ...parts) {
  return parts.map((p) => `${prefix}-${p}`).join(' ');
}
function escapeHtml(text) {
  const div = typeof document === 'undefined' ? null : document.createElement('div');
  if (div) {
    div.textContent = text;
    return div.innerHTML;
  }
  // Fallback for Node.js
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}
function escapeAttribute(value) {
  // Escape for use in HTML attributes (handles quotes and special chars)
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}
function renderItemBar(span, timeCells, stackDepth, prefix) {
  const { item } = span;
  const classes = [cls(prefix, 'item'), cls(prefix, `item--${item.type}`)];
  if (item.className) {
    classes.push(item.className);
  }
  // Add label-left class if label should be positioned to the left
  if (span.labelLeft) {
    classes.push(cls(prefix, 'item--label-left'));
  }
  // Calculate position and width
  const totalCols = span.endCol - span.startCol + 1;
  let left;
  let width;
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
      width = 100 - span.startOffset + (totalCols - 2) * 100 + span.endOffset;
    }
  }
  // Calculate vertical position based on stack row
  // Each stack row gets a fixed height slot
  const itemHeight = 24; // Fixed pixel height per item slot
  const top = span.stackRow * (itemHeight + 4) + 4; // 4px gap between items
  const styles = [`left: ${left.toFixed(1)}%`, `width: ${width.toFixed(1)}%`, `top: ${top}px`];
  // Point items use fixed size, range items use calculated height
  if (!item.isPoint) {
    styles.push(`height: ${itemHeight}px`);
  }
  if (item.style) {
    // Escape quotes to prevent attribute breakout
    styles.push(escapeAttribute(item.style));
  }
  // Add title attribute - for point items, always show title since content is hidden
  const titleText = item.title || (item.isPoint ? item.content : '');
  const titleAttribute = titleText ? ` title="${escapeAttribute(titleText)}"` : '';
  // Escape class names and item ID to prevent XSS
  const safeClasses = classes.map((c) => escapeAttribute(c)).join(' ');
  const safeId = escapeAttribute(item.id);
  return `<div class="${safeClasses}" style="${styles.join('; ')}" data-item-id="${safeId}"${titleAttribute}>
    <span class="${cls(prefix, 'item-content')}">${escapeHtml(item.content)}</span>
  </div>`;
}
function renderCell(spans, colIndex, timeCells, stackDepth, prefix) {
  // Filter spans that start in this column
  const cellSpans = spans.filter((s) => s.startCol === colIndex);
  let content = '';
  for (const span of cellSpans) {
    content += renderItemBar(span, timeCells, stackDepth, prefix);
  }
  return `<td class="${cls(prefix, 'cell')}">
    <div class="${cls(prefix, 'cell-content')}">${content}</div>
  </td>`;
}
function renderGroupRow(group, spans, timeCells, options, rowIndex, timelineWidthPx) {
  const {
    classPrefix: prefix,
    showGroupLabels,
    stackItems: doStack,
    maxStackDepth,
    compactStacking,
    labelPosition,
  } = options;
  const numberCells = timeCells.length;
  const rowClass = `${cls(prefix, 'row')} ${cls(prefix, rowIndex % 2 === 0 ? 'row--even' : 'row--odd')}`;
  const groupClass = group.className ? ` ${escapeAttribute(group.className)}` : '';
  // Stack items for this group
  const groupSpans = spans.filter((s) => s.item.group === group.id);
  const stacks = doStack
    ? stackItems(
        groupSpans,
        maxStackDepth,
        compactStacking,
        timelineWidthPx,
        numberCells,
        labelPosition
      )
    : [groupSpans];
  const stackDepth = stacks.length || 1;
  // Flatten stacks back to array with updated stackRow
  const allSpans = stacks.flat();
  let cells = '';
  // Group label cell
  if (showGroupLabels) {
    const labelStyle = group.style ? ` style="${escapeAttribute(group.style)}"` : '';
    cells += `<td class="${cls(prefix, 'group-label')}${groupClass}"${labelStyle}>${escapeHtml(group.content)}</td>`;
  }
  // Data cells
  for (let index = 0; index < timeCells.length; index++) {
    cells += renderCell(allSpans, index, timeCells, stackDepth, prefix);
  }
  // Calculate row height: 28px per stack row (24px item + 4px gap) + 8px padding
  const rowHeight = stackDepth * 28 + 8;
  const rowStyle = ` style="height: ${rowHeight}px"`;
  return `<tr class="${rowClass}" data-group-id="${escapeAttribute(group.id)}"${rowStyle}>${cells}</tr>`;
}
function renderHeader(timeCells, options) {
  const {
    classPrefix: prefix,
    showGroupLabels,
    groupLabelWidth,
    granularity,
    weekStartDay,
  } = options;
  // Check if we need a grouped header row (year for month/quarter, w/c for day)
  const needsGroupRow =
    granularity === 'month' || granularity === 'quarter' || granularity === 'day';
  const groups = needsGroupRow ? generateHeaderGroups(timeCells, granularity, weekStartDay) : [];
  let headerHtml = '';
  // Render group row (years for month/quarter, w/c for day)
  if (needsGroupRow && groups.length > 0) {
    let groupCells = '';
    if (showGroupLabels) {
      const widthStyle = groupLabelWidth ? ` style="width: ${groupLabelWidth}"` : '';
      groupCells += `<th class="${cls(prefix, 'group-label', 'header-label', 'header-group')}"${widthStyle} rowspan="2"></th>`;
    }
    for (const group of groups) {
      groupCells += `<th class="${cls(prefix, 'header-cell', 'header-group-cell')}" colspan="${group.span}">${escapeHtml(group.label)}</th>`;
    }
    headerHtml += `<tr class="${cls(prefix, 'header', 'header-group-row')}">${groupCells}</tr>`;
  }
  // Render main header row
  let cells = '';
  // Only add group label cell if we don't have a group row (it uses rowspan="2")
  if (showGroupLabels && !needsGroupRow) {
    const widthStyle = groupLabelWidth ? ` style="width: ${groupLabelWidth}"` : '';
    cells += `<th class="${cls(prefix, 'group-label', 'header-label')}"${widthStyle}></th>`;
  }
  for (const cell of timeCells) {
    cells += `<th class="${cls(prefix, 'header-cell')}">${escapeHtml(cell.label)}</th>`;
  }
  headerHtml += `<tr class="${cls(prefix, 'header')}">${cells}</tr>`;
  return headerHtml;
}
function renderTable(groups, allSpans, timeCells, options) {
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
  for (const [index, group] of groups.entries()) {
    if (!group) continue;
    const groupSpans = allSpans.filter((s) => s.item.group === group.id);
    rows += renderGroupRow(group, groupSpans, timeCells, options, index, timelineWidthPx);
  }
  return `<table class="${cls(prefix, 'table')}">
  <thead>${header}</thead>
  <tbody>${rows}</tbody>
</table>`;
}
// ============================================================================
// MAIN RENDERING FUNCTIONS
// ============================================================================
function resolveOptions(items, userOptions) {
  // Auto-detect date range from items
  let start;
  let end;
  start = userOptions?.start
    ? normalizeDate(userOptions.start)
    : new Date(
        items.reduce(
          (min, item) => Math.min(item.start.getTime(), min),
          items[0]?.start.getTime() ?? Date.now()
        )
      );
  end = userOptions?.end
    ? normalizeDate(userOptions.end)
    : new Date(
        items.reduce(
          (max, item) => Math.max(item.end.getTime(), max),
          items[0]?.end.getTime() ?? Date.now()
        )
      );
  // Select granularity first (needed for proper alignment)
  const granularity =
    !userOptions?.granularity || userOptions.granularity === 'auto'
      ? selectGranularity(start, end)
      : userOptions.granularity;
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
    const match = /^(\d+)/.exec(groupLabelWidth);
    if (match?.[1]) {
      groupLabelWidthPx = Number.parseInt(match[1], 10);
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
    labelPosition: userOptions?.labelPosition ?? 'auto',
  };
}
/**
 * Render timeline to HTML string
 */
export function renderTimelineToString(items, groups, options) {
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
  const allSpans = [];
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
export function renderTimeline(items, groups, options) {
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
 * For items with label-left class, text is already positioned outside to the left.
 * @param container - The container element holding the timeline, or document if not specified
 * @param prefix - CSS class prefix (default: 'tl')
 */
export function fixOverflowingText(container, prefix = 'tl') {
  const root = container || document;
  const boxItems = root.querySelectorAll(`.${prefix}-item--range, .${prefix}-item--box`);
  for (const item of boxItems) {
    // Skip items with label-left - their labels are already positioned outside
    if (item.classList.contains(`${prefix}-item--label-left`)) {
      continue;
    }
    const content = item.querySelector(`.${prefix}-item-content`);
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
  }
}
/**
 * Get default CSS styles for the timeline
 */
export function getTimelineStyles(prefix = 'tl') {
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

/* Header group row (years for month/quarter, week commencing for day) */
.${prefix}-header-group-row {
  background-color: #eaeef2;
}

.${prefix}-header-group-cell {
  padding: 6px 4px;
  text-align: center;
  font-weight: 700;
  border: 1px solid #d0d7de;
  color: #24292f;
  font-size: 11px;
}

.${prefix}-header-group {
  vertical-align: middle;
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

/* Label positioned to the left (for items near end of timeline) */
.${prefix}-item--label-left .${prefix}-item-content {
  position: absolute;
  right: 100%;
  left: auto;
  top: 50%;
  transform: translateY(-50%);
  padding: 2px 6px;
  color: #24292f;
  overflow: visible;
  text-align: right;
}

/* Point items with label on left: reverse flex direction */
.${prefix}-item--point.${prefix}-item--label-left {
  flex-direction: row-reverse;
}

.${prefix}-item--point.${prefix}-item--label-left .${prefix}-item-content {
  position: relative;
  right: auto;
  top: auto;
  transform: none;
  text-align: right;
}

/* Box/Range items with label-left that also has text-outside: prioritize label-left */
.${prefix}-item--text-outside.${prefix}-item--label-left .${prefix}-item-content {
  left: auto;
  right: 100%;
  text-align: right;
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
/**
 * Inject timeline CSS styles into the document head
 * @param prefix - CSS class prefix (default: 'tl')
 * @param styleId - ID for the style element (default: 'timeline-styles')
 * @returns The style element
 */
export function injectStyles(prefix = 'tl', styleId = 'timeline-styles') {
  let styleElement = document.getElementById(styleId);
  if (!styleElement) {
    styleElement = document.createElement('style');
    styleElement.id = styleId;
    document.head.append(styleElement);
  }
  styleElement.textContent = getTimelineStyles(prefix);
  return styleElement;
}
/**
 * Create a debounced resize handler
 * @param callback - Function to call on resize
 * @param delay - Debounce delay in ms (default: 150)
 * @returns Object with attach/detach methods
 */
export function createResizeHandler(callback, delay = 150) {
  let timeout = null;
  const handler = () => {
    if (timeout) clearTimeout(timeout);
    timeout = setTimeout(callback, delay);
  };
  return {
    attach: () => {
      window.addEventListener('resize', handler);
    },
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
export function mountTimeline(container, items, groups, options = {}) {
  // Resolve container
  const containerElement =
    typeof container === 'string' ? document.querySelector(container) : container;
  if (!containerElement) {
    throw new Error(`Timeline container not found: ${container}`);
  }
  // Inject styles
  const prefix = options.classPrefix ?? 'tl';
  injectStyles(prefix);
  // Current options
  let currentOptions = { ...options };
  // Render function
  const render = () => {
    // Auto-detect container width if not specified
    const containerWidth = currentOptions.containerWidth ?? containerElement.clientWidth;
    const html = renderTimelineToString(items, groups, {
      ...currentOptions,
      containerWidth,
    });
    containerElement.innerHTML = html;
    fixOverflowingText(containerElement, prefix);
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
    update: (newOptions) => {
      currentOptions = { ...currentOptions, ...newOptions };
      render();
    },
    destroy: () => {
      resizeHandler.detach();
      containerElement.innerHTML = '';
    },
    getOptions: () => ({ ...currentOptions }),
  };
}
// ============================================================================
// PPTX SHARED UTILITIES
// ============================================================================
// Color map for item classes
const pptxColorMap = {
  default: { fill: '0969da', border: '0550ae' },
  success: { fill: '1a7f37', border: '116329' },
  warning: { fill: 'bf8700', border: '9a6700' },
  danger: { fill: 'cf222e', border: 'a40e26' },
  info: { fill: '0550ae', border: '033d8b' },
  purple: { fill: '8250df', border: '6639ba' },
  pink: { fill: 'bf3989', border: '99286e' },
  orange: { fill: 'ea580c', border: 'c2410c' },
};
const DEFAULT_ITEM_COLOR = { fill: '0969da', border: '0550ae' };
function getItemColor(classList) {
  for (const colorName of Object.keys(pptxColorMap)) {
    if (colorName !== 'default' && classList.contains(colorName)) {
      const color = pptxColorMap[colorName];
      if (color) return color;
    }
  }
  return pptxColorMap.default ?? DEFAULT_ITEM_COLOR;
}
/**
 * Extract timeline shapes from the DOM as JSON data for backend PPTX rendering.
 * Mirrors the logic of exportTimelineToPptx but returns structured data instead of
 * calling pptxgenjs methods directly.
 *
 * @param container - The container element or selector containing the rendered timeline
 * @param options - Export options
 * @returns TimelineSlideData with all shapes and texts for rendering
 */
export function extractTimelineShapes(container, options = {}) {
  const containerElement =
    typeof container === 'string' ? document.querySelector(container) : container;
  if (!containerElement) {
    throw new Error(`Timeline container not found: ${container}`);
  }
  const table = containerElement.querySelector('table');
  if (!table) {
    throw new Error('No timeline table found in container');
  }
  const {
    title = 'Timeline',
    slideWidth = 13.33,
    slideHeight = 7.5,
    margin = 0.3,
    titleHeight = 0.6,
    prefix = 'tl',
  } = options;
  const shapes = [];
  const texts = [];
  // Title text
  texts.push({
    content: title,
    x: margin,
    y: margin,
    w: slideWidth - 2 * margin,
    h: titleHeight,
    fontSize: 16,
    bold: true,
    color: '24292f',
  });
  // Get table dimensions from DOM
  const tableRect = table.getBoundingClientRect();
  const tableW = tableRect.width;
  const tableH = tableRect.height;
  // Calculate independent X and Y scales to fill available area
  const availW = slideWidth - 2 * margin;
  const availH = slideHeight - titleHeight - 2 * margin;
  const scaleX = availW / tableW;
  const scaleY = availH / tableH;
  const pptxTableW = tableW * scaleX;
  const pptxTableH = tableH * scaleY;
  const tableX = margin;
  const tableY = margin + titleHeight + 0.2;
  // Helper: convert px to inches on slide
  const toInchX = (px) => tableX + (px / tableW) * pptxTableW;
  const toInchY = (py) => tableY + (py / tableH) * pptxTableH;
  const toInchW = (pw) => (pw / tableW) * pptxTableW;
  const toInchH = (ph) => (ph / tableH) * pptxTableH;
  // Process header rows
  const headerRows = table.querySelectorAll('thead tr');
  for (const headerRow of headerRows) {
    const headerRect = headerRow.getBoundingClientRect();
    const headerY = headerRect.top - tableRect.top;
    const headerH = headerRect.height;
    const isGroupRow = headerRow.classList.contains(`${prefix}-header-group-row`);
    const bgColor = isGroupRow ? 'eaeef2' : 'f6f8fa';
    shapes.push({
      type: 'rect',
      x: tableX,
      y: toInchY(headerY),
      w: pptxTableW,
      h: toInchH(headerH),
      fill: bgColor,
      border: 'd0d7de',
      borderPt: 0.5,
    });
    for (const th of headerRow.querySelectorAll('th')) {
      const thRect = th.getBoundingClientRect();
      const cellX = thRect.left - tableRect.left;
      const cellW = thRect.width;
      const cellH = thRect.height;
      const cellY = thRect.top - tableRect.top;
      texts.push({
        content: th.textContent || '',
        x: toInchX(cellX),
        y: toInchY(cellY),
        w: toInchW(cellW),
        h: toInchH(cellH),
        fontSize: isGroupRow ? 9 : 8,
        bold: true,
        color: '24292f',
        align: 'center',
        valign: 'middle',
      });
    }
  }
  // Process body rows
  const bodyRows = table.querySelectorAll('tbody tr');
  for (const [rowIndex, row] of bodyRows.entries()) {
    const rowRect = row.getBoundingClientRect();
    const rowY = rowRect.top - tableRect.top;
    const rowH = rowRect.height;
    const bgColor = rowIndex % 2 === 0 ? 'FFFFFF' : 'f6f8fa';
    shapes.push({
      type: 'rect',
      x: tableX,
      y: toInchY(rowY),
      w: pptxTableW,
      h: toInchH(rowH),
      fill: bgColor,
      border: 'eaeef2',
      borderPt: 0.25,
    });
    // Group label
    const groupLabel = row.querySelector(`.${prefix}-group-label`);
    if (groupLabel) {
      const labelRect = groupLabel.getBoundingClientRect();
      const labelX = labelRect.left - tableRect.left;
      const labelW = labelRect.width;
      shapes.push({
        type: 'rect',
        x: toInchX(labelX),
        y: toInchY(rowY),
        w: toInchW(labelW),
        h: toInchH(rowH),
        fill: 'f6f8fa',
        border: 'd0d7de',
        borderPt: 0.5,
      });
      texts.push({
        content: groupLabel.textContent || '',
        x: toInchX(labelX) + 0.05,
        y: toInchY(rowY),
        w: toInchW(labelW) - 0.1,
        h: toInchH(rowH),
        fontSize: 8,
        bold: true,
        color: '24292f',
        valign: 'middle',
      });
    }
  }
  // Draw vertical date dividers
  const firstBodyRow = bodyRows[0];
  const bodyRowsArray = [...bodyRows];
  const lastBodyRow = bodyRowsArray.at(-1);
  const headerRowsArray = [...headerRows];
  const mainHeaderRow = headerRowsArray.at(-1);
  if (firstBodyRow && lastBodyRow && mainHeaderRow) {
    const firstRowRect = firstBodyRow.getBoundingClientRect();
    const lastRowRect = lastBodyRow.getBoundingClientRect();
    const gridStartY = firstRowRect.top - tableRect.top;
    const gridEndY = lastRowRect.bottom - tableRect.top;
    mainHeaderRow.querySelectorAll('th').forEach((th, index) => {
      if (index === 0) return;
      const thRect = th.getBoundingClientRect();
      const cellX = thRect.left - tableRect.left;
      shapes.push({
        type: 'line',
        x: toInchX(cellX),
        y: toInchY(gridStartY),
        w: 0,
        h: toInchH(gridEndY - gridStartY),
        border: 'eaeef2',
        borderPt: 0.5,
      });
    });
    // Right edge of last column
    const lastTh = mainHeaderRow.querySelector('th:last-child');
    if (lastTh) {
      const lastThRect = lastTh.getBoundingClientRect();
      const rightEdge = lastThRect.right - tableRect.left;
      shapes.push({
        type: 'line',
        x: toInchX(rightEdge),
        y: toInchY(gridStartY),
        w: 0,
        h: toInchH(gridEndY - gridStartY),
        border: 'eaeef2',
        borderPt: 0.5,
      });
    }
  }
  // Draw timeline items
  for (const row of bodyRows) {
    for (const item of row.querySelectorAll(`.${prefix}-item`)) {
      const itemRect = item.getBoundingClientRect();
      const itemX = itemRect.left - tableRect.left;
      const itemY = itemRect.top - tableRect.top;
      const itemW = itemRect.width;
      const itemH = itemRect.height;
      const isPoint = item.classList.contains(`${prefix}-item--point`);
      const isLabelLeft = item.classList.contains(`${prefix}-item--label-left`);
      const colors = getItemColor(item.classList);
      const contentElement = item.querySelector(`.${prefix}-item-content`);
      const content = contentElement?.textContent || '';
      if (isPoint) {
        const dotSize = 0.12;
        const labelGap = 0.04;
        const dotX = toInchX(itemX);
        const dotCenterY = toInchY(itemY) + toInchH(itemH) / 2;
        shapes.push({
          type: 'ellipse',
          x: dotX,
          y: dotCenterY - dotSize / 2,
          w: dotSize,
          h: dotSize,
          fill: colors.fill,
          border: colors.border,
          borderPt: 1.5,
        });
        // Use row height for text box so valign: middle aligns with dot center
        const textH = toInchH(itemH);
        const textY = toInchY(itemY);
        const textW = 4;
        if (isLabelLeft) {
          texts.push({
            content,
            x: dotX - textW - labelGap,
            y: textY,
            w: textW,
            h: textH,
            fontSize: 8,
            color: '24292f',
            valign: 'middle',
            align: 'right',
          });
        } else {
          const textX = dotX + dotSize + labelGap;
          texts.push({
            content,
            x: textX,
            y: textY,
            w: textW,
            h: textH,
            fontSize: 8,
            color: '24292f',
            valign: 'middle',
          });
        }
      } else {
        // Box/range item
        const labelGap = 0.04;
        shapes.push({
          type: 'rect',
          x: toInchX(itemX),
          y: toInchY(itemY),
          w: Math.max(toInchW(itemW), 0.1),
          h: toInchH(itemH),
          fill: colors.fill,
          border: colors.border,
          borderPt: 0.75,
          rectRadius: 0.03,
        });
        const textOutside = item.classList.contains(`${prefix}-item--text-outside`);
        if (textOutside || isLabelLeft) {
          const boxTextW = 4;
          if (isLabelLeft) {
            texts.push({
              content,
              x: toInchX(itemX) - boxTextW - labelGap,
              y: toInchY(itemY),
              w: boxTextW,
              h: toInchH(itemH),
              fontSize: 8,
              color: '24292f',
              valign: 'middle',
              align: 'right',
            });
          } else {
            texts.push({
              content,
              x: toInchX(itemX) + toInchW(itemW) + labelGap,
              y: toInchY(itemY),
              w: boxTextW,
              h: toInchH(itemH),
              fontSize: 8,
              color: '24292f',
              valign: 'middle',
            });
          }
        } else {
          texts.push({
            content,
            x: toInchX(itemX),
            y: toInchY(itemY),
            w: Math.max(toInchW(itemW) - 0.06, 0.1),
            h: toInchH(itemH),
            fontSize: 8,
            color: 'FFFFFF',
            valign: 'middle',
          });
        }
      }
    }
  }
  return { title, slideWidth, slideHeight, shapes, texts };
}
// ============================================================================
// PPTX EXPORT (legacy client-side via pptxgenjs)
// ============================================================================
/**
 * Export a rendered timeline to PowerPoint using pptxgenjs.
 * Requires pptxgenjs to be loaded (e.g., via CDN or npm).
 *
 * @param container - The container element or selector containing the rendered timeline
 * @param pptx - A PptxGenJS instance
 * @param options - Export options
 * @returns The slide that was created
 */
export function exportTimelineToPptx(container, pptx, options = {}) {
  const containerElement =
    typeof container === 'string' ? document.querySelector(container) : container;
  if (!containerElement) {
    throw new Error(`Timeline container not found: ${container}`);
  }
  const table = containerElement.querySelector('table');
  if (!table) {
    throw new Error('No timeline table found in container');
  }
  const {
    title = 'Timeline',
    slideWidth = 13.33,
    slideHeight = 7.5,
    margin = 0.3,
    titleHeight = 0.6,
    prefix = 'tl',
  } = options;
  const slide = pptx.addSlide();
  // Add title
  slide.addText(title, {
    x: margin,
    y: margin,
    w: slideWidth - 2 * margin,
    h: titleHeight,
    fontSize: 16,
    bold: true,
    color: '24292f',
  });
  // Get table dimensions from DOM
  const tableRect = table.getBoundingClientRect();
  const tableW = tableRect.width;
  const tableH = tableRect.height;
  // Calculate independent X and Y scales to fill available area
  const availW = slideWidth - 2 * margin;
  const availH = slideHeight - titleHeight - 2 * margin;
  const scaleX = availW / tableW;
  const scaleY = availH / tableH;
  const pptxTableW = tableW * scaleX;
  const pptxTableH = tableH * scaleY;
  const tableX = margin;
  const tableY = margin + titleHeight + 0.2;
  // Helper: convert px to inches on slide
  const toInchX = (px) => tableX + (px / tableW) * pptxTableW;
  const toInchY = (py) => tableY + (py / tableH) * pptxTableH;
  const toInchW = (pw) => (pw / tableW) * pptxTableW;
  const toInchH = (ph) => (ph / tableH) * pptxTableH;
  // Process header rows (may have 1 or 2 rows for grouped headers)
  const headerRows = table.querySelectorAll('thead tr');
  for (const headerRow of headerRows) {
    const headerRect = headerRow.getBoundingClientRect();
    const headerY = headerRect.top - tableRect.top;
    const headerH = headerRect.height;
    // Header background - use darker color for group row
    const isGroupRow = headerRow.classList.contains(`${prefix}-header-group-row`);
    const bgColor = isGroupRow ? 'eaeef2' : 'f6f8fa';
    slide.addShape(pptx.ShapeType.rect, {
      x: tableX,
      y: toInchY(headerY),
      w: pptxTableW,
      h: toInchH(headerH),
      fill: { color: bgColor },
      line: { color: 'd0d7de', pt: 0.5 },
    });
    // Header cells text
    for (const th of headerRow.querySelectorAll('th')) {
      const thRect = th.getBoundingClientRect();
      const cellX = thRect.left - tableRect.left;
      const cellW = thRect.width;
      const cellH = thRect.height;
      const cellY = thRect.top - tableRect.top;
      slide.addText(th.textContent || '', {
        x: toInchX(cellX),
        y: toInchY(cellY),
        w: toInchW(cellW),
        h: toInchH(cellH),
        fontSize: isGroupRow ? 9 : 8,
        bold: true,
        color: '24292f',
        align: 'center',
        valign: 'middle',
      });
    }
  }
  // Process body rows
  const bodyRows = table.querySelectorAll('tbody tr');
  for (const [rowIndex, row] of bodyRows.entries()) {
    const rowRect = row.getBoundingClientRect();
    const rowY = rowRect.top - tableRect.top;
    const rowH = rowRect.height;
    // Row background
    const bgColor = rowIndex % 2 === 0 ? 'FFFFFF' : 'f6f8fa';
    slide.addShape(pptx.ShapeType.rect, {
      x: tableX,
      y: toInchY(rowY),
      w: pptxTableW,
      h: toInchH(rowH),
      fill: { color: bgColor },
      line: { color: 'eaeef2', pt: 0.25 },
    });
    // Group label
    const groupLabel = row.querySelector(`.${prefix}-group-label`);
    if (groupLabel) {
      const labelRect = groupLabel.getBoundingClientRect();
      const labelX = labelRect.left - tableRect.left;
      const labelW = labelRect.width;
      slide.addShape(pptx.ShapeType.rect, {
        x: toInchX(labelX),
        y: toInchY(rowY),
        w: toInchW(labelW),
        h: toInchH(rowH),
        fill: { color: 'f6f8fa' },
        line: { color: 'd0d7de', pt: 0.5 },
      });
      slide.addText(groupLabel.textContent || '', {
        x: toInchX(labelX) + 0.05,
        y: toInchY(rowY),
        w: toInchW(labelW) - 0.1,
        h: toInchH(rowH),
        fontSize: 8,
        bold: true,
        color: '24292f',
        valign: 'middle',
      });
    }
  }
  // Draw vertical date dividers
  const firstBodyRow = bodyRows[0];
  const bodyRowsArray = [...bodyRows];
  const lastBodyRow = bodyRowsArray.at(-1);
  // Use the last header row (main date labels, not group row) for column positions
  const headerRowsArray = [...headerRows];
  const mainHeaderRow = headerRowsArray.at(-1);
  if (firstBodyRow && lastBodyRow && mainHeaderRow) {
    const firstRowRect = firstBodyRow.getBoundingClientRect();
    const lastRowRect = lastBodyRow.getBoundingClientRect();
    const gridStartY = firstRowRect.top - tableRect.top;
    const gridEndY = lastRowRect.bottom - tableRect.top;
    mainHeaderRow.querySelectorAll('th').forEach((th, index) => {
      if (index === 0) return; // Skip group label column
      const thRect = th.getBoundingClientRect();
      const cellX = thRect.left - tableRect.left;
      slide.addShape(pptx.ShapeType.line, {
        x: toInchX(cellX),
        y: toInchY(gridStartY),
        w: 0,
        h: toInchH(gridEndY - gridStartY),
        line: { color: 'eaeef2', pt: 0.5 },
      });
    });
    // Right edge of last column
    const lastTh = mainHeaderRow.querySelector('th:last-child');
    if (lastTh) {
      const lastThRect = lastTh.getBoundingClientRect();
      const rightEdge = lastThRect.right - tableRect.left;
      slide.addShape(pptx.ShapeType.line, {
        x: toInchX(rightEdge),
        y: toInchY(gridStartY),
        w: 0,
        h: toInchH(gridEndY - gridStartY),
        line: { color: 'eaeef2', pt: 0.5 },
      });
    }
  }
  // Draw timeline items
  for (const row of bodyRows) {
    for (const item of row.querySelectorAll(`.${prefix}-item`)) {
      const itemRect = item.getBoundingClientRect();
      const itemX = itemRect.left - tableRect.left;
      const itemY = itemRect.top - tableRect.top;
      const itemW = itemRect.width;
      const itemH = itemRect.height;
      const isPoint = item.classList.contains(`${prefix}-item--point`);
      const isLabelLeft = item.classList.contains(`${prefix}-item--label-left`);
      const colors = getItemColor(item.classList);
      const contentElement = item.querySelector(`.${prefix}-item-content`);
      const content = contentElement?.textContent || '';
      if (isPoint) {
        // Point item: draw circle
        const dotSize = 0.12;
        const labelGap = 0.04;
        const dotX = toInchX(itemX);
        const dotCenterY = toInchY(itemY) + toInchH(itemH) / 2;
        slide.addShape(pptx.ShapeType.ellipse, {
          x: dotX,
          y: dotCenterY - dotSize / 2,
          w: dotSize,
          h: dotSize,
          fill: { color: colors.fill },
          line: { color: colors.border, pt: 1.5 },
        });
        // Use row height for text box so valign: middle aligns with dot center
        const textH = toInchH(itemH);
        const textY = toInchY(itemY);
        const textW = 4;
        if (isLabelLeft) {
          // Text to the left of the dot, right-aligned
          slide.addText(content, {
            x: dotX - textW - labelGap,
            y: textY,
            w: textW,
            h: textH,
            fontSize: 8,
            color: '24292f',
            valign: 'middle',
            align: 'right',
          });
        } else {
          // Text to the right of the dot (default)
          const textX = dotX + dotSize + labelGap;
          slide.addText(content, {
            x: textX,
            y: textY,
            w: textW,
            h: textH,
            fontSize: 8,
            color: '24292f',
            valign: 'middle',
          });
        }
      } else {
        // Box/range item
        const labelGap = 0.04;
        slide.addShape(pptx.ShapeType.rect, {
          x: toInchX(itemX),
          y: toInchY(itemY),
          w: Math.max(toInchW(itemW), 0.1),
          h: toInchH(itemH),
          fill: { color: colors.fill },
          line: { color: colors.border, pt: 0.75 },
          rectRadius: 0.03,
        });
        const textOutside = item.classList.contains(`${prefix}-item--text-outside`);
        if (textOutside || isLabelLeft) {
          // Text outside the bar
          const boxTextW = 4;
          if (isLabelLeft) {
            // Text to the left of the bar, right-aligned
            slide.addText(content, {
              x: toInchX(itemX) - boxTextW - labelGap,
              y: toInchY(itemY),
              w: boxTextW,
              h: toInchH(itemH),
              fontSize: 8,
              color: '24292f',
              valign: 'middle',
              align: 'right',
            });
          } else {
            // Text to the right of the bar (default text-outside behavior)
            slide.addText(content, {
              x: toInchX(itemX) + toInchW(itemW) + labelGap,
              y: toInchY(itemY),
              w: boxTextW,
              h: toInchH(itemH),
              fontSize: 8,
              color: '24292f',
              valign: 'middle',
            });
          }
        } else {
          // Text inside the bar
          slide.addText(content, {
            x: toInchX(itemX),
            y: toInchY(itemY),
            w: Math.max(toInchW(itemW) - 0.06, 0.1),
            h: toInchH(itemH),
            fontSize: 8,
            color: 'FFFFFF',
            valign: 'middle',
          });
        }
      }
    }
  }
  return slide;
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
  exportTimelineToPptx,
  extractTimelineShapes,
};
