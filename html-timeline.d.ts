/**
 * HTML Timeline Generator
 *
 * A TypeScript module that generates HTML table-based timelines using
 * vis-timeline compatible data structures. Outputs pure HTML/CSS that
 * displays well on web and exports cleanly to PowerPoint via pptxgenjs.
 */
/** Item type enumeration - matches vis-timeline types */
export type TimelineItemType = 'box' | 'point' | 'range' | 'background';
/** Time granularity levels */
export type TimeGranularity = 'hour' | 'day' | 'week' | 'month' | 'quarter' | 'year';
/** Group ordering options */
export type GroupOrderOption = 'value' | 'content' | 'id';
/** Timeline date value - accepts Date object, ISO string, or timestamp */
export type TimelineDate = Date | string | number;
/** Timeline ID value - accepts string or numeric identifier */
export type TimelineId = string | number;
/**
 * Timeline item - vis-timeline compatible structure
 */
export interface TimelineItem {
  /** Unique identifier for the item */
  id: TimelineId;
  /** Display content (text or HTML) */
  content: string;
  /** Start date/time of the item */
  start: TimelineDate;
  /** End date/time (optional for point items) */
  end?: TimelineDate;
  /** Item display type */
  type?: TimelineItemType;
  /** Group ID this item belongs to */
  group?: TimelineId;
  /** CSS class name(s) for styling */
  className?: string;
  /** Custom inline styles */
  style?: string;
  /** Title/tooltip text */
  title?: string;
  /** Stacking order within group (lower = higher priority for earlier rows) */
  order?: number;
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
/** Label position behavior for items extending past timeline end */
export type LabelPosition = 'auto' | 'right';
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
  groupOrder?: GroupOrderOption;
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
  /**
   * Label position behavior for items near the end of the timeline.
   * - 'auto' (default): Labels that would extend past the timeline end are
   *   positioned to the left of the bar/point, right-aligned with the same gap spacing.
   * - 'right': Labels always positioned to the right (original behavior).
   */
  labelPosition?: LabelPosition;
}
/**
 * Minimal interface for PptxGenJS slide object.
 * Only includes methods used by exportTimelineToPptx.
 */
interface PptxSlide {
  addText(text: string, options?: Record<string, unknown>): void;
  addShape(shape: string, options?: Record<string, unknown>): void;
}
/**
 * Minimal interface for PptxGenJS instance.
 * Only includes methods used by exportTimelineToPptx.
 */
interface PptxGenJS {
  addSlide(): PptxSlide;
  ShapeType: {
    rect: string;
    line: string;
    ellipse: string;
  };
}
/**
 * Render timeline to HTML string
 */
export declare function renderTimelineToString(
  items: TimelineItem[],
  groups: TimelineGroup[],
  options?: TimelineOptions
): string;
/**
 * Render timeline to DOM element (browser only)
 */
export declare function renderTimeline(
  items: TimelineItem[],
  groups: TimelineGroup[],
  options?: TimelineOptions
): HTMLTableElement;
/**
 * Fix overflowing text in box/range items (browser only)
 * Call this after rendering to move text outside boxes that are too small.
 * For items with label-left class, text is already positioned outside to the left.
 * @param container - The container element holding the timeline, or document if not specified
 * @param prefix - CSS class prefix (default: 'tl')
 */
export declare function fixOverflowingText(container?: Element, prefix?: string): void;
/**
 * Get default CSS styles for the timeline
 */
export declare function getTimelineStyles(prefix?: string): string;
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
export declare function injectStyles(prefix?: string, styleId?: string): HTMLStyleElement;
/**
 * Create a debounced resize handler
 * @param callback - Function to call on resize
 * @param delay - Debounce delay in ms (default: 150)
 * @returns Object with attach/detach methods
 */
export declare function createResizeHandler(
  callback: () => void,
  delay?: number
): {
  attach: () => void;
  detach: () => void;
};
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
export declare function mountTimeline(
  container: HTMLElement | string,
  items: TimelineItem[],
  groups: TimelineGroup[],
  options?: MountOptions
): MountedTimeline;
export interface PptxExportOptions {
  title?: string;
  slideWidth?: number;
  slideHeight?: number;
  margin?: number;
  titleHeight?: number;
  prefix?: string;
}
/** A shape (rectangle, ellipse, line) in the extracted timeline data. */
export interface TimelineShape {
  type: 'rect' | 'ellipse' | 'line';
  x: number;
  y: number;
  w: number;
  h: number;
  fill?: string;
  border?: string;
  borderPt?: number;
  rectRadius?: number;
}
/** A text element in the extracted timeline data. */
export interface TimelineText {
  content: string;
  x: number;
  y: number;
  w: number;
  h: number;
  fontSize: number;
  bold?: boolean;
  color: string;
  align?: 'left' | 'center' | 'right';
  valign?: 'top' | 'middle' | 'bottom';
}
/** Full slide data extracted from the DOM for backend PPTX rendering. */
export interface TimelineSlideData {
  title: string;
  slideWidth: number;
  slideHeight: number;
  shapes: TimelineShape[];
  texts: TimelineText[];
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
export declare function extractTimelineShapes(
  container: HTMLElement | string,
  options?: PptxExportOptions
): TimelineSlideData;
/**
 * Export a rendered timeline to PowerPoint using pptxgenjs.
 * Requires pptxgenjs to be loaded (e.g., via CDN or npm).
 *
 * @param container - The container element or selector containing the rendered timeline
 * @param pptx - A PptxGenJS instance
 * @param options - Export options
 * @returns The slide that was created
 */
export declare function exportTimelineToPptx(
  container: HTMLElement | string,
  pptx: PptxGenJS,
  options?: PptxExportOptions
): PptxSlide;
declare const _default: {
  renderTimeline: typeof renderTimeline;
  renderTimelineToString: typeof renderTimelineToString;
  getTimelineStyles: typeof getTimelineStyles;
  fixOverflowingText: typeof fixOverflowingText;
  injectStyles: typeof injectStyles;
  createResizeHandler: typeof createResizeHandler;
  mountTimeline: typeof mountTimeline;
  exportTimelineToPptx: typeof exportTimelineToPptx;
  extractTimelineShapes: typeof extractTimelineShapes;
};
export default _default;
