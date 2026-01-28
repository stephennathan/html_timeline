/**
 * Test Data Set 6: Order Field for Stack Priority
 *
 * Challenge: Controlling visual stacking order explicitly.
 * Tests the order field's ability to:
 * - Place high-priority items on earlier rows regardless of start date
 * - Mix ordered and unordered items correctly
 * - Handle items with the same order value
 * - Work with both compact and normal stacking modes
 */

export const items = [
  // Group 1: Explicit ordering - lower order = higher row
  // Despite starting later, Critical Path should appear on row 1
  {
    id: 1,
    content: 'Critical Path (order: 1)',
    group: 1,
    start: '2024-02-01',
    end: '2024-02-28',
    type: 'box',
    className: 'danger',
    order: 1,
  },
  {
    id: 2,
    content: 'Medium Priority (order: 3)',
    group: 1,
    start: '2024-01-01',
    end: '2024-01-31',
    type: 'box',
    className: 'info',
    order: 3,
  },
  {
    id: 4,
    content: 'Low Priority (order: 4)',
    group: 1,
    start: '2024-01-08',
    end: '2024-02-08',
    type: 'box',
    className: 'purple',
    order: 4,
  },
  {
    id: 3,
    content: 'High Priority (order: 2)',
    group: 1,
    start: '2024-01-15',
    end: '2024-02-15',
    type: 'box',
    className: 'warning',
    order: 2,
  },

  // Group 2: Mixed ordered and unordered items
  // Ordered items should appear first, then unordered by date
  {
    id: 5,
    content: 'Must Do First (order: 1)',
    group: 2,
    start: '2024-02-15',
    end: '2024-03-01',
    type: 'box',
    className: 'danger',
    order: 1,
  },
  {
    id: 6,
    content: 'Then This (order: 2)',
    group: 2,
    start: '2024-02-01',
    end: '2024-02-20',
    type: 'box',
    className: 'warning',
    order: 2,
  },
  {
    id: 7,
    content: 'No order - starts Jan 1',
    group: 2,
    start: '2024-01-01',
    end: '2024-01-20',
    type: 'box',
    className: 'info',
  },
  {
    id: 8,
    content: 'No order - starts Jan 10',
    group: 2,
    start: '2024-01-10',
    end: '2024-01-30',
    type: 'box',
    className: 'purple',
  },
  {
    id: 9,
    content: 'No order - starts Feb 1',
    group: 2,
    start: '2024-02-01',
    end: '2024-02-15',
    type: 'box',
    className: 'pink',
  },

  // Group 3: Same order values - should fall back to date sorting
  {
    id: 10,
    content: 'Same order=1, starts Feb',
    group: 3,
    start: '2024-02-01',
    end: '2024-02-15',
    type: 'box',
    className: 'danger',
    order: 1,
  },
  {
    id: 11,
    content: 'Same order=1, starts Jan',
    group: 3,
    start: '2024-01-15',
    end: '2024-01-31',
    type: 'box',
    className: 'warning',
    order: 1,
  },
  {
    id: 12,
    content: 'Same order=2, longer',
    group: 3,
    start: '2024-01-01',
    end: '2024-02-28',
    type: 'box',
    className: 'info',
    order: 2,
  },
  {
    id: 13,
    content: 'Same order=2, shorter',
    group: 3,
    start: '2024-01-01',
    end: '2024-01-15',
    type: 'box',
    className: 'purple',
    order: 2,
  },

  // Group 4: Point items with order
  {
    id: 14,
    content: 'Top Milestone',
    group: 4,
    start: '2024-02-15',
    type: 'point',
    className: 'danger',
    order: 1,
  },
  {
    id: 15,
    content: 'Second Milestone',
    group: 4,
    start: '2024-01-15',
    type: 'point',
    className: 'warning',
    order: 2,
  },
  {
    id: 16,
    content: 'Third Milestone',
    group: 4,
    start: '2024-01-01',
    type: 'point',
    className: 'success',
    order: 3,
  },
  {
    id: 17,
    content: 'Unordered Milestone 1',
    group: 4,
    start: '2024-02-01',
    type: 'point',
    className: 'info',
  },
  {
    id: 18,
    content: 'Unordered Milestone 2',
    group: 4,
    start: '2024-02-08',
    type: 'point',
    className: 'purple',
  },

  // Group 5: Overlapping items with explicit order for compact mode testing
  {
    id: 19,
    content: 'Row 1: Alpha',
    group: 5,
    start: '2024-01-01',
    end: '2024-01-15',
    type: 'box',
    className: 'danger',
    order: 1,
  },
  {
    id: 20,
    content: 'Row 1: Beta',
    group: 5,
    start: '2024-01-20',
    end: '2024-02-05',
    type: 'box',
    className: 'danger',
    order: 1,
  },
  {
    id: 21,
    content: 'Row 2: Gamma',
    group: 5,
    start: '2024-01-01',
    end: '2024-01-10',
    type: 'box',
    className: 'warning',
    order: 2,
  },
  {
    id: 22,
    content: 'Row 2: Delta',
    group: 5,
    start: '2024-01-15',
    end: '2024-02-01',
    type: 'box',
    className: 'warning',
    order: 2,
  },
  {
    id: 23,
    content: 'Row 3: Epsilon',
    group: 5,
    start: '2024-01-05',
    end: '2024-01-25',
    type: 'box',
    className: 'info',
    order: 3,
  },
  {
    id: 24,
    content: 'Row 3: Zeta',
    group: 5,
    start: '2024-01-28',
    end: '2024-02-10',
    type: 'box',
    className: 'info',
    order: 3,
  },
];

export const groups = [
  { id: 1, content: 'Explicit Priority Order', value: 1 },
  { id: 2, content: 'Mixed Ordered/Unordered', value: 2 },
  { id: 3, content: 'Same Order Values', value: 3 },
  { id: 4, content: 'Point Items with Order', value: 4 },
  { id: 5, content: 'Compact Mode Test', value: 5 },
];

export const description = 'Order field controls stacking priority - lower order = earlier rows';
