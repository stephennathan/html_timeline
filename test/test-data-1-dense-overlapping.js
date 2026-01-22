/**
 * Test Data Set 1: Dense Overlapping Items
 *
 * Challenge: Many items in the same group with significant overlap.
 * Tests the stacking algorithm's ability to handle:
 * - Multiple items starting on the same date
 * - Items with varying durations that overlap
 * - Compact stacking efficiency with crowded timelines
 */

export const items = [
  // Sprint 1 - Multiple items starting same day
  {
    id: 1,
    content: 'Backend API Development',
    group: 1,
    start: '2024-01-08',
    end: '2024-01-26',
    type: 'box'
  },
  {
    id: 2,
    content: 'Frontend Setup',
    group: 1,
    start: '2024-01-08',
    end: '2024-01-19',
    type: 'box',
    className: 'info'
  },
  {
    id: 3,
    content: 'Database Schema Design',
    group: 1,
    start: '2024-01-08',
    end: '2024-01-12',
    type: 'box',
    className: 'purple'
  },
  {
    id: 4,
    content: 'CI/CD Pipeline',
    group: 1,
    start: '2024-01-08',
    end: '2024-01-15',
    type: 'box',
    className: 'warning'
  },
  {
    id: 5,
    content: 'Sprint 1 Kickoff',
    group: 1,
    start: '2024-01-08',
    type: 'point',
    className: 'success'
  },

  // Sprint 1 - Overlapping mid-sprint items
  {
    id: 6,
    content: 'Authentication Module',
    group: 1,
    start: '2024-01-15',
    end: '2024-01-26',
    type: 'box'
  },
  {
    id: 7,
    content: 'User Management UI',
    group: 1,
    start: '2024-01-15',
    end: '2024-01-22',
    type: 'box',
    className: 'info'
  },
  {
    id: 8,
    content: 'API Documentation',
    group: 1,
    start: '2024-01-18',
    end: '2024-01-26',
    type: 'box',
    className: 'pink'
  },
  {
    id: 9,
    content: 'Sprint 1 Demo',
    group: 1,
    start: '2024-01-26',
    type: 'point',
    className: 'success'
  },

  // Sprint 2 - Dense overlapping
  {
    id: 10,
    content: 'Payment Integration',
    group: 2,
    start: '2024-01-29',
    end: '2024-02-16',
    type: 'box',
    className: 'danger'
  },
  {
    id: 11,
    content: 'Order Processing',
    group: 2,
    start: '2024-01-29',
    end: '2024-02-09',
    type: 'box'
  },
  {
    id: 12,
    content: 'Inventory System',
    group: 2,
    start: '2024-01-29',
    end: '2024-02-09',
    type: 'box',
    className: 'purple'
  },
  {
    id: 13,
    content: 'Shipping Module',
    group: 2,
    start: '2024-02-05',
    end: '2024-02-16',
    type: 'box',
    className: 'info'
  },
  {
    id: 14,
    content: 'Tax Calculation',
    group: 2,
    start: '2024-02-05',
    end: '2024-02-12',
    type: 'box',
    className: 'warning'
  },
  {
    id: 15,
    content: 'Discount Engine',
    group: 2,
    start: '2024-02-08',
    end: '2024-02-16',
    type: 'box',
    className: 'pink'
  },
  {
    id: 16,
    content: 'Sprint 2 Kickoff',
    group: 2,
    start: '2024-01-29',
    type: 'point',
    className: 'success'
  },
  { id: 17, content: 'Mid-Sprint Review', group: 2, start: '2024-02-07', type: 'point' },
  {
    id: 18,
    content: 'Sprint 2 Demo',
    group: 2,
    start: '2024-02-16',
    type: 'point',
    className: 'success'
  },

  // Sprint 3 - Even denser
  {
    id: 19,
    content: 'Performance Optimization',
    group: 3,
    start: '2024-02-19',
    end: '2024-03-01',
    type: 'box'
  },
  {
    id: 20,
    content: 'Security Audit',
    group: 3,
    start: '2024-02-19',
    end: '2024-03-01',
    type: 'box',
    className: 'danger'
  },
  {
    id: 21,
    content: 'Load Testing',
    group: 3,
    start: '2024-02-19',
    end: '2024-02-26',
    type: 'box',
    className: 'warning'
  },
  {
    id: 22,
    content: 'Bug Fixes',
    group: 3,
    start: '2024-02-19',
    end: '2024-03-01',
    type: 'box',
    className: 'pink'
  },
  {
    id: 23,
    content: 'Code Review',
    group: 3,
    start: '2024-02-22',
    end: '2024-03-01',
    type: 'box',
    className: 'purple'
  },
  {
    id: 24,
    content: 'Documentation Update',
    group: 3,
    start: '2024-02-22',
    end: '2024-02-28',
    type: 'box',
    className: 'info'
  },
  {
    id: 25,
    content: 'Release Prep',
    group: 3,
    start: '2024-02-26',
    end: '2024-03-01',
    type: 'box',
    className: 'success'
  },
  {
    id: 26,
    content: 'v1.0 Release',
    group: 3,
    start: '2024-03-01',
    type: 'point',
    className: 'success'
  }
];

export const groups = [
  { id: 1, content: 'Sprint 1 (Jan 8-26)', value: 1 },
  { id: 2, content: 'Sprint 2 (Jan 29 - Feb 16)', value: 2 },
  { id: 3, content: 'Sprint 3 (Feb 19 - Mar 1)', value: 3 }
];

export const description = 'Dense overlapping items - tests stacking with many concurrent tasks';
