/**
 * Test Data Set 5: Edge Cases and Boundary Conditions
 *
 * Challenge: Various edge cases that might break the timeline.
 * Tests:
 * - Single day items
 * - Items at exact boundaries
 * - Empty groups
 * - All item types together
 * - Background items
 * - Numeric vs string IDs
 * - Items without groups
 * - Same start and end date
 */

export const items = [
  // Single day range items (start === end conceptually)
  {
    id: 1,
    content: 'One Day Sprint',
    group: 1,
    start: '2024-05-01',
    end: '2024-05-01',
    type: 'box'
  },
  {
    id: 2,
    content: 'Single Day Task',
    group: 1,
    start: '2024-05-03',
    end: '2024-05-03',
    type: 'box',
    className: 'info'
  },

  // Point items clustering
  { id: 3, content: 'Morning Standup', group: 1, start: '2024-05-06', type: 'point' },
  { id: 4, content: 'Team Sync', group: 1, start: '2024-05-06', type: 'point', className: 'info' },
  {
    id: 5,
    content: 'Code Review',
    group: 1,
    start: '2024-05-06',
    type: 'point',
    className: 'purple'
  },

  // Background items
  {
    id: 6,
    content: 'Conference Week',
    group: 2,
    start: '2024-05-13',
    end: '2024-05-17',
    type: 'background'
  },
  {
    id: 7,
    content: 'Holiday Period',
    group: 2,
    start: '2024-05-27',
    end: '2024-05-31',
    type: 'background',
    className: 'warning'
  },

  // Mix of types in same group
  { id: 8, content: 'Planning', group: 2, start: '2024-05-01', end: '2024-05-03', type: 'box' },
  { id: 9, content: 'Kickoff', group: 2, start: '2024-05-06', type: 'point', className: 'success' },
  {
    id: 10,
    content: 'Development',
    group: 2,
    start: '2024-05-06',
    end: '2024-05-24',
    type: 'box',
    className: 'info'
  },
  { id: 11, content: 'Demo', group: 2, start: '2024-05-24', type: 'point', className: 'success' },

  // Very short items next to very long items
  {
    id: 12,
    content: 'Quick Fix',
    group: 3,
    start: '2024-05-01',
    end: '2024-05-02',
    type: 'box',
    className: 'danger'
  },
  {
    id: 13,
    content: 'Extended Maintenance Window',
    group: 3,
    start: '2024-05-01',
    end: '2024-05-31',
    type: 'box',
    className: 'warning'
  },
  {
    id: 14,
    content: 'Hotfix Deploy',
    group: 3,
    start: '2024-05-15',
    type: 'point',
    className: 'danger'
  },

  // Items at exact month boundaries
  { id: 15, content: 'Month Start', group: 4, start: '2024-05-01', type: 'point' },
  { id: 16, content: 'Month End', group: 4, start: '2024-05-31', type: 'point' },
  {
    id: 17,
    content: 'Full Month',
    group: 4,
    start: '2024-05-01',
    end: '2024-05-31',
    type: 'box',
    className: 'purple'
  },

  // Overlapping background with regular items
  {
    id: 18,
    content: 'Sprint Background',
    group: 5,
    start: '2024-05-06',
    end: '2024-05-17',
    type: 'background',
    className: 'info'
  },
  { id: 19, content: 'Feature A', group: 5, start: '2024-05-06', end: '2024-05-10', type: 'box' },
  {
    id: 20,
    content: 'Feature B',
    group: 5,
    start: '2024-05-13',
    end: '2024-05-17',
    type: 'box',
    className: 'success'
  },
  { id: 21, content: 'Review', group: 5, start: '2024-05-17', type: 'point', className: 'purple' },

  // Numeric IDs (vs string IDs above)
  {
    id: 100,
    content: 'Numeric ID 100',
    group: 6,
    start: '2024-05-08',
    end: '2024-05-15',
    type: 'box'
  },
  {
    id: 101,
    content: 'Numeric ID 101',
    group: 6,
    start: '2024-05-20',
    type: 'point',
    className: 'info'
  },

  // Items with all color classes
  {
    id: 22,
    content: 'Default Color',
    group: 7,
    start: '2024-05-01',
    end: '2024-05-04',
    type: 'box'
  },
  {
    id: 23,
    content: 'Success',
    group: 7,
    start: '2024-05-06',
    end: '2024-05-09',
    type: 'box',
    className: 'success'
  },
  {
    id: 24,
    content: 'Warning',
    group: 7,
    start: '2024-05-11',
    end: '2024-05-14',
    type: 'box',
    className: 'warning'
  },
  {
    id: 25,
    content: 'Danger',
    group: 7,
    start: '2024-05-16',
    end: '2024-05-19',
    type: 'box',
    className: 'danger'
  },
  {
    id: 26,
    content: 'Info',
    group: 7,
    start: '2024-05-21',
    end: '2024-05-24',
    type: 'box',
    className: 'info'
  },
  {
    id: 27,
    content: 'Purple',
    group: 7,
    start: '2024-05-26',
    end: '2024-05-29',
    type: 'box',
    className: 'purple'
  },
  {
    id: 28,
    content: 'Pink',
    group: 7,
    start: '2024-05-31',
    end: '2024-06-03',
    type: 'box',
    className: 'pink'
  },

  // Adjacent items (end date = next start date)
  { id: 29, content: 'Phase 1', group: 8, start: '2024-05-01', end: '2024-05-10', type: 'box' },
  {
    id: 30,
    content: 'Phase 2',
    group: 8,
    start: '2024-05-10',
    end: '2024-05-20',
    type: 'box',
    className: 'info'
  },
  {
    id: 31,
    content: 'Phase 3',
    group: 8,
    start: '2024-05-20',
    end: '2024-05-31',
    type: 'box',
    className: 'purple'
  },

  // Empty content (should still render)
  {
    id: 32,
    content: '',
    group: 9,
    start: '2024-05-10',
    end: '2024-05-15',
    type: 'box',
    className: 'warning'
  },
  { id: 33, content: '', group: 9, start: '2024-05-20', type: 'point', className: 'danger' },

  // Special characters in content
  {
    id: 34,
    content: "Test <script>alert('xss')</script>",
    group: 10,
    start: '2024-05-05',
    end: '2024-05-10',
    type: 'box'
  },
  {
    id: 35,
    content: 'Quotes "double" and \'single\'',
    group: 10,
    start: '2024-05-12',
    end: '2024-05-17',
    type: 'box',
    className: 'info'
  },
  {
    id: 36,
    content: 'Ampersand & More',
    group: 10,
    start: '2024-05-19',
    end: '2024-05-24',
    type: 'box',
    className: 'purple'
  },
  {
    id: 37,
    content: 'Unicode: 日本語 🚀 émojis',
    group: 10,
    start: '2024-05-26',
    type: 'point',
    className: 'success'
  }
];

export const groups = [
  { id: 1, content: 'Single Day Items', value: 1 },
  { id: 2, content: 'Mixed Types + Background', value: 2 },
  { id: 3, content: 'Short vs Long', value: 3 },
  { id: 4, content: 'Boundary Items', value: 4 },
  { id: 5, content: 'Background Overlay', value: 5 },
  { id: 6, content: 'Numeric IDs', value: 6 },
  { id: 7, content: 'All Colors', value: 7 },
  { id: 8, content: 'Adjacent Phases', value: 8 },
  { id: 9, content: 'Empty Content', value: 9 },
  { id: 10, content: 'Special Characters', value: 10 },
  { id: 11, content: 'Empty Group (No Items)', value: 11 } // Intentionally no items
];

export const description =
  'Edge cases - tests boundaries, special characters, empty states, all types';
