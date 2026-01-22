/**
 * Test Data Set 4: Many Groups (Swimlanes)
 *
 * Challenge: Large number of groups/swimlanes with varying item counts.
 * Tests:
 * - Group ordering (by value)
 * - Group label rendering with many rows
 * - Mixed density across groups
 * - Performance with many groups
 */

export const items = [
  // Executive Team (sparse)
  {
    id: 1,
    content: 'Q1 Strategy Review',
    group: 'exec',
    start: '2024-01-15',
    type: 'point',
    className: 'purple'
  },
  {
    id: 2,
    content: 'Board Meeting',
    group: 'exec',
    start: '2024-02-20',
    type: 'point',
    className: 'danger'
  },
  {
    id: 3,
    content: 'Investor Update',
    group: 'exec',
    start: '2024-03-15',
    type: 'point',
    className: 'success'
  },

  // Product Team
  {
    id: 4,
    content: 'Roadmap Planning',
    group: 'product',
    start: '2024-01-08',
    end: '2024-01-19',
    type: 'box'
  },
  {
    id: 5,
    content: 'Feature Specs',
    group: 'product',
    start: '2024-01-22',
    end: '2024-02-02',
    type: 'box',
    className: 'info'
  },
  {
    id: 6,
    content: 'User Research',
    group: 'product',
    start: '2024-02-05',
    end: '2024-02-23',
    type: 'box',
    className: 'pink'
  },
  { id: 7, content: 'Design Review', group: 'product', start: '2024-02-26', type: 'point' },

  // Engineering - Backend
  {
    id: 8,
    content: 'API v2 Development',
    group: 'eng-be',
    start: '2024-01-15',
    end: '2024-03-15',
    type: 'box'
  },
  {
    id: 9,
    content: 'Database Migration',
    group: 'eng-be',
    start: '2024-02-01',
    end: '2024-02-28',
    type: 'box',
    className: 'danger'
  },
  {
    id: 10,
    content: 'Performance Tuning',
    group: 'eng-be',
    start: '2024-03-01',
    end: '2024-03-22',
    type: 'box',
    className: 'warning'
  },

  // Engineering - Frontend
  {
    id: 11,
    content: 'UI Component Library',
    group: 'eng-fe',
    start: '2024-01-08',
    end: '2024-02-16',
    type: 'box',
    className: 'info'
  },
  {
    id: 12,
    content: 'Dashboard Redesign',
    group: 'eng-fe',
    start: '2024-02-19',
    end: '2024-03-15',
    type: 'box',
    className: 'purple'
  },
  {
    id: 13,
    content: 'Mobile Responsive',
    group: 'eng-fe',
    start: '2024-03-04',
    end: '2024-03-29',
    type: 'box',
    className: 'pink'
  },

  // Engineering - Platform
  {
    id: 14,
    content: 'CI/CD Improvements',
    group: 'eng-platform',
    start: '2024-01-15',
    end: '2024-02-09',
    type: 'box',
    className: 'warning'
  },
  {
    id: 15,
    content: 'Kubernetes Migration',
    group: 'eng-platform',
    start: '2024-02-12',
    end: '2024-03-22',
    type: 'box'
  },
  {
    id: 16,
    content: 'Monitoring Setup',
    group: 'eng-platform',
    start: '2024-02-26',
    end: '2024-03-15',
    type: 'box',
    className: 'info'
  },

  // QA Team
  {
    id: 17,
    content: 'Test Automation Framework',
    group: 'qa',
    start: '2024-01-22',
    end: '2024-02-23',
    type: 'box'
  },
  {
    id: 18,
    content: 'Integration Testing',
    group: 'qa',
    start: '2024-02-26',
    end: '2024-03-15',
    type: 'box',
    className: 'warning'
  },
  {
    id: 19,
    content: 'UAT Support',
    group: 'qa',
    start: '2024-03-18',
    end: '2024-03-29',
    type: 'box',
    className: 'success'
  },

  // Security Team
  {
    id: 20,
    content: 'Security Audit',
    group: 'security',
    start: '2024-02-05',
    end: '2024-02-23',
    type: 'box',
    className: 'danger'
  },
  {
    id: 21,
    content: 'Pen Testing',
    group: 'security',
    start: '2024-03-04',
    end: '2024-03-15',
    type: 'box',
    className: 'danger'
  },
  {
    id: 22,
    content: 'Compliance Review',
    group: 'security',
    start: '2024-03-18',
    end: '2024-03-29',
    type: 'box',
    className: 'warning'
  },

  // DevOps
  {
    id: 23,
    content: 'Infrastructure as Code',
    group: 'devops',
    start: '2024-01-08',
    end: '2024-02-02',
    type: 'box'
  },
  {
    id: 24,
    content: 'Disaster Recovery',
    group: 'devops',
    start: '2024-02-05',
    end: '2024-02-23',
    type: 'box',
    className: 'danger'
  },
  {
    id: 25,
    content: 'Cost Optimization',
    group: 'devops',
    start: '2024-02-26',
    end: '2024-03-15',
    type: 'box',
    className: 'success'
  },

  // Data Team
  {
    id: 26,
    content: 'Data Warehouse Setup',
    group: 'data',
    start: '2024-01-15',
    end: '2024-02-16',
    type: 'box',
    className: 'purple'
  },
  {
    id: 27,
    content: 'ETL Pipelines',
    group: 'data',
    start: '2024-02-19',
    end: '2024-03-15',
    type: 'box'
  },
  {
    id: 28,
    content: 'Analytics Dashboard',
    group: 'data',
    start: '2024-03-04',
    end: '2024-03-29',
    type: 'box',
    className: 'info'
  },

  // Design Team
  {
    id: 29,
    content: 'Brand Refresh',
    group: 'design',
    start: '2024-01-08',
    end: '2024-01-26',
    type: 'box',
    className: 'pink'
  },
  {
    id: 30,
    content: 'UI Kit Update',
    group: 'design',
    start: '2024-01-29',
    end: '2024-02-16',
    type: 'box',
    className: 'purple'
  },
  {
    id: 31,
    content: 'Marketing Assets',
    group: 'design',
    start: '2024-02-19',
    end: '2024-03-08',
    type: 'box',
    className: 'info'
  },

  // Marketing Team
  {
    id: 32,
    content: 'Campaign Planning',
    group: 'marketing',
    start: '2024-01-15',
    end: '2024-02-02',
    type: 'box'
  },
  {
    id: 33,
    content: 'Content Creation',
    group: 'marketing',
    start: '2024-02-05',
    end: '2024-03-01',
    type: 'box',
    className: 'pink'
  },
  {
    id: 34,
    content: 'Launch Campaign',
    group: 'marketing',
    start: '2024-03-04',
    end: '2024-03-29',
    type: 'box',
    className: 'success'
  },

  // Sales Team
  {
    id: 35,
    content: 'Sales Training',
    group: 'sales',
    start: '2024-02-19',
    end: '2024-03-01',
    type: 'box',
    className: 'info'
  },
  {
    id: 36,
    content: 'Partner Outreach',
    group: 'sales',
    start: '2024-03-04',
    end: '2024-03-22',
    type: 'box'
  },
  {
    id: 37,
    content: 'Launch Prep',
    group: 'sales',
    start: '2024-03-25',
    type: 'point',
    className: 'success'
  },

  // Customer Success
  {
    id: 38,
    content: 'Documentation',
    group: 'cs',
    start: '2024-02-12',
    end: '2024-03-08',
    type: 'box',
    className: 'info'
  },
  {
    id: 39,
    content: 'Training Materials',
    group: 'cs',
    start: '2024-03-11',
    end: '2024-03-29',
    type: 'box',
    className: 'purple'
  },

  // Legal
  {
    id: 40,
    content: 'Contract Reviews',
    group: 'legal',
    start: '2024-01-08',
    end: '2024-01-26',
    type: 'box',
    className: 'warning'
  },
  {
    id: 41,
    content: 'Terms Update',
    group: 'legal',
    start: '2024-02-05',
    end: '2024-02-16',
    type: 'box'
  },
  {
    id: 42,
    content: 'Compliance Certification',
    group: 'legal',
    start: '2024-03-04',
    type: 'point',
    className: 'success'
  },

  // HR
  {
    id: 43,
    content: 'Hiring Sprint',
    group: 'hr',
    start: '2024-01-15',
    end: '2024-02-16',
    type: 'box'
  },
  {
    id: 44,
    content: 'Onboarding Program',
    group: 'hr',
    start: '2024-02-19',
    end: '2024-03-15',
    type: 'box',
    className: 'info'
  },

  // Key Milestones (across all)
  {
    id: 45,
    content: 'Alpha Release',
    group: 'exec',
    start: '2024-02-02',
    type: 'point',
    className: 'success'
  },
  {
    id: 46,
    content: 'Beta Release',
    group: 'exec',
    start: '2024-03-01',
    type: 'point',
    className: 'success'
  },
  {
    id: 47,
    content: 'GA Release',
    group: 'exec',
    start: '2024-03-29',
    type: 'point',
    className: 'success'
  }
];

export const groups = [
  { id: 'exec', content: 'Executive', value: 1 },
  { id: 'product', content: 'Product', value: 2 },
  { id: 'eng-be', content: 'Engineering - Backend', value: 3 },
  { id: 'eng-fe', content: 'Engineering - Frontend', value: 4 },
  { id: 'eng-platform', content: 'Engineering - Platform', value: 5 },
  { id: 'qa', content: 'QA', value: 6 },
  { id: 'security', content: 'Security', value: 7 },
  { id: 'devops', content: 'DevOps', value: 8 },
  { id: 'data', content: 'Data Engineering', value: 9 },
  { id: 'design', content: 'Design', value: 10 },
  { id: 'marketing', content: 'Marketing', value: 11 },
  { id: 'sales', content: 'Sales', value: 12 },
  { id: 'cs', content: 'Customer Success', value: 13 },
  { id: 'legal', content: 'Legal', value: 14 },
  { id: 'hr', content: 'Human Resources', value: 15 }
];

export const description =
  'Many groups (15 swimlanes) - tests group ordering and rendering performance';
