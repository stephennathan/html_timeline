# Claude Code Context

## Project Overview
HTML Timeline Generator - A TypeScript module that generates HTML table-based timelines using vis-timeline compatible data structures. Outputs pure HTML/CSS that displays well on web and exports cleanly to PowerPoint via pptxgenjs.

## Key Files
- `html-timeline.ts` - Main TypeScript source module
- `html-timeline.js` - Compiled ES module (DO NOT edit directly - compile from .ts)
- `html-timeline.d.ts` - TypeScript type declarations
- `test-html-timeline.html` - Test/demo page with sample data

## Build
To compile TypeScript to JavaScript:
```bash
npx tsc html-timeline.ts --outDir . --declaration --module ES2020 --target ES2020 --moduleResolution node --esModuleInterop --strict
```
**Important:** Never edit `html-timeline.js` directly. Always edit `html-timeline.ts` and compile.

## Architecture

### Data Structures (vis-timeline compatible)
```typescript
interface TimelineItem {
  id: string | number;
  content: string;           // Display text
  start: Date | string;      // Start date
  end?: Date | string;       // End date (optional for points)
  type?: 'box' | 'point' | 'range' | 'background';
  group?: string | number;   // Group ID
  className?: string;        // CSS class for styling (e.g., 'success', 'purple')
  order?: number;            // Stacking order (lower = higher priority for earlier rows)
}

interface TimelineGroup {
  id: string | number;
  content: string;           // Group label
  value?: number;            // Sort order
}
```

### Key Options
- `granularity`: 'auto' | 'hour' | 'day' | 'week' | 'month' | 'quarter' | 'year'
- `compactStacking`: boolean - Allow items to share rows when they don't overlap
- `containerWidth`: number - Pass actual container width for accurate text width estimation

### Stacking Algorithm
The stacking algorithm in `spansOverlap()` and `stackItems()` determines row placement:

1. **Normal mode** (`compactStacking: false`): Each item gets its own row - no horizontal stacking
2. **Compact mode** (`compactStacking: true`): Items can share rows if they don't overlap
   - Uses percentage-based overlap detection (consistent across granularities)
   - `estimateTextPercent()` calculates text width as percentage of timeline
   - Box items account for text flowing outside when bar is narrow

### CSS Classes
Built-in color classes: `success`, `warning`, `danger`, `info`, `purple`, `pink`

Custom colors can be added in the HTML:
```css
.tl-item.custom-color {
  background: linear-gradient(135deg, #color1 0%, #color2 100%);
  border-color: #color2;
}
.tl-item--point.custom-color::before {
  background-color: #color1;
  border-color: #color2;
}
```

## Browser Helpers

### `mountTimeline(container, items, groups, options)`
One-call setup that handles everything:
- Injects CSS styles
- Renders timeline HTML
- Calls fixOverflowingText
- Sets up debounced resize handler
- Auto-detects container width

Returns a `MountedTimeline` object:
```typescript
{
  render(): void;           // Re-render with current options
  update(options): void;    // Update options and re-render
  destroy(): void;          // Clean up listeners
  getOptions(): MountOptions;
}
```

### `injectStyles(prefix?, styleId?)`
Injects timeline CSS into document head. Safe to call multiple times.

### `createResizeHandler(callback, delay?)`
Creates a debounced resize handler. Returns `{ attach(), detach() }`.

## Common Tasks

### Modifying stacking behavior
Edit `stackItems()` and `spansOverlap()` in `html-timeline.ts`, then compile. Key functions:
- `stackItems()` - Controls row assignment (normal vs compact mode), sorts by `order` field first
- `spansOverlap()` - Determines if two items overlap using percentage-based positions
- `estimateTextPercent()` - Calculates text width as percentage of timeline

### Adding new color classes
Add to `getTimelineStyles()` function following the existing pattern.

### Changing text width estimation
Modify `estimateTextPercent()` - currently uses 6px per character average.

## Testing
1. Run `npx serve` in the project directory
2. Open `test-html-timeline.html` in browser
3. Toggle "Compact stacking" checkbox to test stacking modes
4. Change granularity dropdown to test different time scales

The test file demonstrates using `mountTimeline()` with `timeline.update()` for control changes:
```javascript
const timeline = mountTimeline('#timeline', items, groups, {
  granularity: 'auto',
  compactStacking: false,
});

// Update on control change
document.getElementById('granularity').addEventListener('change', (e) => {
  timeline.update({ granularity: e.target.value });
});
```
