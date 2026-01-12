# HTML Timeline Generator

A TypeScript module that generates HTML table-based timelines using vis-timeline compatible data structures. Outputs pure HTML/CSS that displays well on web and exports cleanly to PowerPoint via pptxgenjs's `tableToSlides()`.

## Features

- **vis-timeline compatible** - Uses the same data structures as vis-timeline
- **Pure HTML/CSS output** - No JavaScript dependencies for rendering
- **Multiple granularities** - Hour, day, week, month, quarter, year (or auto-detect)
- **Grouped items** - Organize items into labeled groups/swimlanes
- **Item types** - Box (range), point (milestone), and background items
- **Smart stacking** - Each item on its own row by default
- **Compact mode** - Items share rows when they don't overlap
- **Responsive** - Adjusts to container width
- **Customizable colors** - Built-in color classes with easy customization
- **PowerPoint export** - Compatible with pptxgenjs tableToSlides()

## Installation

```bash
# Copy html-timeline.js to your project
# Or use the TypeScript source: html-timeline.ts
```

## Usage

### Simple (Recommended)

```javascript
import { mountTimeline } from './html-timeline.js';

const items = [
  { id: 1, content: "Project Kickoff", start: "2024-01-15", type: "point", group: 1 },
  { id: 2, content: "Development Phase", start: "2024-02-01", end: "2024-04-30", group: 1 },
  { id: 3, content: "Launch", start: "2024-05-01", type: "point", group: 1, className: "success" }
];

const groups = [
  { id: 1, content: "Project Alpha", value: 1 }
];

// Mount timeline with one call - handles styles, rendering, and resize
const timeline = mountTimeline('#timeline', items, groups, {
  granularity: 'month',
  compactStacking: true,
});

// Update options later
timeline.update({ compactStacking: false });

// Clean up when done
timeline.destroy();
```

### Manual Setup

```javascript
import { renderTimelineToString, getTimelineStyles, fixOverflowingText } from './html-timeline.js';

// Inject CSS styles
document.getElementById('timeline-styles').textContent = getTimelineStyles();

// Render timeline
const html = renderTimelineToString(items, groups, {
  granularity: 'month',
  compactStacking: true,
  containerWidth: document.querySelector('.container').clientWidth
});

document.getElementById('timeline').innerHTML = html;

// Fix text that overflows box items
fixOverflowingText();
```

## Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `start` | Date/string | auto | Timeline start date |
| `end` | Date/string | auto | Timeline end date |
| `granularity` | string | 'auto' | Time granularity: 'hour', 'day', 'week', 'month', 'quarter', 'year', 'auto' |
| `showGroupLabels` | boolean | true | Show group name column |
| `groupLabelWidth` | string | '180px' | Width of group label column |
| `stackItems` | boolean | true | Stack overlapping items |
| `maxStackDepth` | number | 20 | Maximum rows per group |
| `compactStacking` | boolean | false | Allow items to share rows when they don't overlap |
| `containerWidth` | number | null | Container width for responsive stacking |
| `classPrefix` | string | 'tl' | CSS class prefix |
| `locale` | string | 'en-US' | Date formatting locale |
| `groupOrder` | string | 'value' | Group ordering: 'value', 'content', 'id' |
| `resizeDebounce` | number | 150 | Debounce delay for resize handler (mountTimeline only) |
| `disableResize` | boolean | false | Disable automatic resize handling (mountTimeline only) |
| `onRender` | function | null | Callback after each render (mountTimeline only) |

## Item Types

- **box/range** - Colored bar spanning start to end date
- **point** - Milestone marker (dot) at start date
- **background** - Subtle background highlighting

## Color Classes

Built-in classes that can be applied via `className`:

- `success` - Green
- `warning` - Amber/Yellow
- `danger` - Red
- `info` - Blue
- `purple` - Purple
- `pink` - Pink

### Custom Colors

Add custom colors in your CSS:

```css
/* Custom orange color */
.tl-item.orange {
  background: linear-gradient(135deg, #ea580c 0%, #c2410c 100%);
  border-color: #c2410c;
}
.tl-item--point.orange::before {
  background-color: #ea580c;
  border-color: #c2410c;
}
```

## Stacking Modes

**Normal mode** (`compactStacking: false`, default): Each item gets its own row. This provides a clean, predictable layout where every item is clearly visible on a separate line.

**Compact mode** (`compactStacking: true`): Items can share rows when their text doesn't overlap. This produces a denser timeline with fewer rows.

For best results with compact stacking:
- Pass `containerWidth` for accurate text width calculations
- Stacking is consistent across all granularities (week, month, etc.)
- Keep labels concise when possible

## API

### `mountTimeline(container, items, groups, options?): MountedTimeline`
Mounts a timeline with full setup (styles, rendering, resize handling).
Returns an object with `render()`, `update(options)`, `destroy()`, and `getOptions()` methods.

### `renderTimelineToString(items, groups, options?): string`
Renders timeline to HTML string.

### `renderTimeline(items, groups, options?): HTMLTableElement`
Renders timeline to DOM element (browser only).

### `getTimelineStyles(prefix?): string`
Returns default CSS styles.

### `injectStyles(prefix?, styleId?): HTMLStyleElement`
Injects timeline CSS into the document head.

### `fixOverflowingText(container?, prefix?): void`
Moves text outside box items when it doesn't fit inside.

### `createResizeHandler(callback, delay?): { attach, detach }`
Creates a debounced resize handler with attach/detach methods.

## License

MIT
(c) Steve Nathan 2026