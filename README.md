# HTML Timeline Generator

A TypeScript module that generates HTML table-based timelines using vis-timeline compatible data structures. Outputs pure HTML/CSS that displays well on web and exports cleanly to PowerPoint via pptxgenjs.

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
- **PowerPoint export** - Export to PPTX with `exportTimelineToPptx()`

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

## Styling

### CSS Injection

The timeline CSS is automatically injected when using `mountTimeline()`. For manual setup, call `injectStyles()`:

```javascript
import { injectStyles } from './html-timeline.js';

// Inject core timeline styles into document head
injectStyles();
```

### Built-in Color Classes

Apply these via the `className` property on items:

| Class | Color |
|-------|-------|
| `success` | Green |
| `warning` | Amber/Yellow |
| `danger` | Red |
| `info` | Blue |
| `purple` | Purple |
| `pink` | Pink |

Example:
```javascript
{ id: 1, content: "Launch", start: "2024-05-01", type: "point", className: "success" }
```

### Custom Colors

Add custom color classes in your own CSS. The pattern differs for box/range items vs point items:

```css
/* Custom orange color for box/range items */
.tl-item.orange {
  background: linear-gradient(135deg, #ea580c 0%, #c2410c 100%);
  border-color: #c2410c;
}

/* Custom orange color for point items (milestones) */
.tl-item--point.orange::before {
  background-color: #ea580c;
  border-color: #c2410c;
}
```

Then use it like the built-in classes:
```javascript
{ id: 1, content: "Urgent", start: "2024-03-15", end: "2024-03-20", className: "orange" }
```

### Overriding Default Styles

To customize the default appearance, add CSS rules after the timeline styles are injected:

```css
/* Make all items have rounded corners */
.tl-item {
  border-radius: 8px;
}

/* Change default blue color */
.tl-item--range,
.tl-item--box {
  background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%);
  border-color: #4f46e5;
}

/* Customize group labels */
.tl-group-label {
  font-size: 13px;
  font-weight: 600;
}
```

### CSS Class Reference

Core classes (prefixed with `tl-` by default):

| Class | Description |
|-------|-------------|
| `.tl-table` | Main table element |
| `.tl-header` | Header row |
| `.tl-header-cell` | Date column headers |
| `.tl-row` | Body row |
| `.tl-row--even`, `.tl-row--odd` | Alternating row styles |
| `.tl-group-label` | Group name cell |
| `.tl-cell` | Data cell |
| `.tl-item` | All timeline items |
| `.tl-item--box`, `.tl-item--range` | Box/range items |
| `.tl-item--point` | Point/milestone items |
| `.tl-item--background` | Background items |
| `.tl-item--text-outside` | Items where text overflows outside |
| `.tl-item-content` | Text content inside items |

## Stacking Modes

**Normal mode** (`compactStacking: false`, default): Each item gets its own row. This provides a clean, predictable layout where every item is clearly visible on a separate line.

**Compact mode** (`compactStacking: true`): Items can share rows when their text doesn't overlap. This produces a denser timeline with fewer rows.

For best results with compact stacking:
- Pass `containerWidth` for accurate text width calculations
- Stacking is consistent across all granularities (week, month, etc.)
- Keep labels concise when possible

## API

`mountTimeline(container, items, groups, options?): MountedTimeline`
Mounts a timeline with full setup (styles, rendering, resize handling).
Returns an object with `render()`, `update(options)`, `destroy()`, and `getOptions()` methods.

`renderTimelineToString(items, groups, options?): string`
Renders timeline to HTML string.

`renderTimeline(items, groups, options?): HTMLTableElement`
Renders timeline to DOM element (browser only).

`getTimelineStyles(prefix?): string`
Returns default CSS styles.

`injectStyles(prefix?, styleId?): HTMLStyleElement`
Injects timeline CSS into the document head.

 `fixOverflowingText(container?, prefix?): void`
Moves text outside box items when it doesn't fit inside.

 `createResizeHandler(callback, delay?): { attach, detach }`
Creates a debounced resize handler with attach/detach methods.

 `exportTimelineToPptx(container, pptx, options?): slide`
Exports a rendered timeline to PowerPoint. Requires pptxgenjs.

## PowerPoint Export

Export timelines to PowerPoint using pptxgenjs:

```html
<script src="https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js"></script>
```

```javascript
import { mountTimeline, exportTimelineToPptx } from './html-timeline.js';

// First, render the timeline
const timeline = mountTimeline('#timeline', items, groups, options);

// Then export to PPTX
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE'; // 13.33" x 7.5"

exportTimelineToPptx('#timeline', pptx, {
  title: 'My Project Timeline',
});

await pptx.writeFile({ fileName: 'timeline.pptx' });
```

### Export Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `title` | string | 'Timeline' | Slide title |
| `slideWidth` | number | 13.33 | Slide width in inches |
| `slideHeight` | number | 7.5 | Slide height in inches |
| `margin` | number | 0.3 | Page margin in inches |
| `titleHeight` | number | 0.6 | Title area height in inches |
| `prefix` | string | 'tl' | CSS class prefix |

## License

MIT
(c) Steve Nathan 2026