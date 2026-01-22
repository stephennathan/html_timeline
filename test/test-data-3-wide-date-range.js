/**
 * Test Data Set 3: Wide Date Range (Multi-Year)
 *
 * Challenge: Timeline spanning multiple years with varying granularities.
 * Tests:
 * - Auto granularity selection
 * - Year/quarter/month transitions
 * - Very long duration items
 * - Sparse vs dense periods
 */

export const items = [
  // 2023 - Sparse beginning
  {
    id: 1,
    content: 'Market Research',
    group: 1,
    start: '2023-01-15',
    end: '2023-06-30',
    type: 'box',
    className: 'info'
  },
  { id: 2, content: 'Initial Concept', group: 1, start: '2023-03-01', type: 'point' },
  {
    id: 3,
    content: 'Feasibility Study',
    group: 1,
    start: '2023-04-01',
    end: '2023-07-31',
    type: 'box',
    className: 'purple'
  },
  {
    id: 4,
    content: 'Board Approval',
    group: 1,
    start: '2023-08-15',
    type: 'point',
    className: 'success'
  },

  // 2023-2024 - Long spanning items
  {
    id: 5,
    content: 'Platform Development',
    group: 2,
    start: '2023-09-01',
    end: '2024-08-31',
    type: 'box'
  },
  {
    id: 6,
    content: 'Infrastructure Setup',
    group: 2,
    start: '2023-09-01',
    end: '2023-12-31',
    type: 'box',
    className: 'warning'
  },
  {
    id: 7,
    content: 'Core Services',
    group: 2,
    start: '2024-01-01',
    end: '2024-06-30',
    type: 'box',
    className: 'info'
  },
  {
    id: 8,
    content: 'Beta Testing',
    group: 2,
    start: '2024-03-01',
    end: '2024-07-31',
    type: 'box',
    className: 'pink'
  },
  { id: 9, content: 'Alpha Release', group: 2, start: '2024-01-15', type: 'point' },
  {
    id: 10,
    content: 'Beta Release',
    group: 2,
    start: '2024-05-01',
    type: 'point',
    className: 'success'
  },

  // 2024 - Dense activity
  {
    id: 11,
    content: 'Partner Onboarding Phase 1',
    group: 3,
    start: '2024-06-01',
    end: '2024-08-31',
    type: 'box',
    className: 'purple'
  },
  {
    id: 12,
    content: 'Partner Onboarding Phase 2',
    group: 3,
    start: '2024-09-01',
    end: '2024-11-30',
    type: 'box',
    className: 'purple'
  },
  {
    id: 13,
    content: 'Marketing Campaign',
    group: 3,
    start: '2024-07-01',
    end: '2024-10-31',
    type: 'box',
    className: 'pink'
  },
  {
    id: 14,
    content: 'Public Launch',
    group: 3,
    start: '2024-09-15',
    type: 'point',
    className: 'success'
  },
  {
    id: 15,
    content: 'First 1000 Users',
    group: 3,
    start: '2024-10-20',
    type: 'point',
    className: 'success'
  },

  // 2024-2025 - Continued growth
  {
    id: 16,
    content: 'International Expansion',
    group: 4,
    start: '2024-11-01',
    end: '2025-06-30',
    type: 'box',
    className: 'info'
  },
  { id: 17, content: 'Europe Launch', group: 4, start: '2025-01-15', type: 'point' },
  { id: 18, content: 'Asia Pacific Launch', group: 4, start: '2025-04-01', type: 'point' },
  {
    id: 19,
    content: 'Regulatory Compliance',
    group: 4,
    start: '2024-10-01',
    end: '2025-03-31',
    type: 'box',
    className: 'danger'
  },
  {
    id: 20,
    content: 'Series B Funding',
    group: 4,
    start: '2025-02-28',
    type: 'point',
    className: 'success'
  },

  // 2025 - Future planning
  {
    id: 21,
    content: 'Platform v2.0 Development',
    group: 5,
    start: '2025-03-01',
    end: '2025-12-31',
    type: 'box'
  },
  {
    id: 22,
    content: 'AI Features Integration',
    group: 5,
    start: '2025-04-01',
    end: '2025-09-30',
    type: 'box',
    className: 'purple'
  },
  {
    id: 23,
    content: 'Mobile App Launch',
    group: 5,
    start: '2025-06-01',
    end: '2025-08-31',
    type: 'box',
    className: 'info'
  },
  {
    id: 24,
    content: 'Enterprise Edition',
    group: 5,
    start: '2025-07-01',
    end: '2025-11-30',
    type: 'box',
    className: 'warning'
  },
  {
    id: 25,
    content: 'IPO Preparation',
    group: 5,
    start: '2025-09-01',
    end: '2025-12-31',
    type: 'box',
    className: 'success'
  },
  {
    id: 26,
    content: 'Target: IPO',
    group: 5,
    start: '2026-01-15',
    type: 'point',
    className: 'success'
  }
];

export const groups = [
  { id: 1, content: 'Discovery (2023)', value: 1 },
  { id: 2, content: 'Development (2023-24)', value: 2 },
  { id: 3, content: 'Launch (2024)', value: 3 },
  { id: 4, content: 'Expansion (2024-25)', value: 4 },
  { id: 5, content: 'Scale (2025-26)', value: 5 }
];

export const description =
  'Wide date range (3+ years) - tests granularity auto-detection and long spans';
