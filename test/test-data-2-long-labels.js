/**
 * Test Data Set 2: Long Labels and Text Overflow
 *
 * Challenge: Items with very long content text.
 * Tests:
 * - Text overflow handling for box items
 * - fixOverflowingText() function effectiveness
 * - Label positioning with narrow date ranges
 * - Compact stacking text width estimation
 */

export const items = [
  // Very long labels on short duration items
  {
    id: 1,
    content: 'Initial Project Kickoff Meeting with All Stakeholders and Executive Leadership',
    group: 1,
    start: '2024-03-01',
    end: '2024-03-03',
    type: 'box'
  },
  {
    id: 2,
    content: 'Requirements Gathering and Documentation Phase Including User Research',
    group: 1,
    start: '2024-03-04',
    end: '2024-03-08',
    type: 'box',
    className: 'info'
  },
  {
    id: 3,
    content: 'Architecture Review Board Presentation and Technical Design Approval',
    group: 1,
    start: '2024-03-11',
    end: '2024-03-13',
    type: 'box',
    className: 'purple'
  },

  // Long labels on point items
  {
    id: 4,
    content: 'Critical Milestone: Complete System Architecture Documentation Review',
    group: 1,
    start: '2024-03-15',
    type: 'point',
    className: 'success'
  },
  {
    id: 5,
    content: 'Stakeholder Sign-off on Technical Requirements Document v2.0',
    group: 1,
    start: '2024-03-18',
    type: 'point'
  },

  // Mix of long and short labels
  { id: 6, content: 'Development', group: 2, start: '2024-03-18', end: '2024-04-15', type: 'box' },
  {
    id: 7,
    content: 'Comprehensive Integration Testing Phase with External Partner Systems',
    group: 2,
    start: '2024-03-25',
    end: '2024-04-05',
    type: 'box',
    className: 'warning'
  },
  {
    id: 8,
    content: 'QA',
    group: 2,
    start: '2024-04-01',
    end: '2024-04-12',
    type: 'box',
    className: 'info'
  },
  {
    id: 9,
    content: 'User Acceptance Testing with Production-Like Environment Configuration',
    group: 2,
    start: '2024-04-08',
    end: '2024-04-15',
    type: 'box',
    className: 'pink'
  },

  // Extremely long labels
  {
    id: 10,
    content:
      'Cross-Functional Team Alignment Meeting to Discuss Q2 Strategic Priorities and Resource Allocation for Mission-Critical Initiatives',
    group: 3,
    start: '2024-04-01',
    end: '2024-04-02',
    type: 'box',
    className: 'purple'
  },
  {
    id: 11,
    content:
      'Enterprise Architecture Review: Microservices Migration Strategy Assessment and Cloud Infrastructure Modernization Roadmap Planning Session',
    group: 3,
    start: '2024-04-03',
    end: '2024-04-05',
    type: 'box',
    className: 'danger'
  },

  // Long labels with overlapping dates (stacking challenge)
  {
    id: 12,
    content: 'Security Compliance Audit and Penetration Testing Report Generation',
    group: 4,
    start: '2024-04-15',
    end: '2024-04-22',
    type: 'box',
    className: 'danger'
  },
  {
    id: 13,
    content: 'Performance Benchmark Analysis and Optimization Recommendations',
    group: 4,
    start: '2024-04-15',
    end: '2024-04-19',
    type: 'box',
    className: 'warning'
  },
  {
    id: 14,
    content: 'Infrastructure Scaling Assessment for Expected Holiday Traffic Surge',
    group: 4,
    start: '2024-04-18',
    end: '2024-04-25',
    type: 'box',
    className: 'info'
  },
  {
    id: 15,
    content: 'Final Production Deployment with Zero-Downtime Migration Strategy',
    group: 4,
    start: '2024-04-25',
    end: '2024-04-26',
    type: 'box',
    className: 'success'
  },
  {
    id: 16,
    content: 'Post-Deployment Monitoring and Incident Response Readiness Verification',
    group: 4,
    start: '2024-04-26',
    type: 'point',
    className: 'success'
  },

  // Additional items for documentation and training
  {
    id: 17,
    content: 'Technical Documentation Update for New API Endpoints and Integration Guidelines',
    group: 5,
    start: '2024-04-29',
    end: '2024-05-03',
    type: 'box',
    className: 'info'
  },
  {
    id: 18,
    content: 'End User Training Material Development and Learning Management System Upload',
    group: 5,
    start: '2024-04-29',
    end: '2024-05-06',
    type: 'box',
    className: 'purple'
  },
  {
    id: 19,
    content: 'Administrator Guide Creation with Troubleshooting Procedures and Escalation Paths',
    group: 5,
    start: '2024-05-01',
    end: '2024-05-08',
    type: 'box'
  },
  {
    id: 20,
    content: 'Training Session Delivery: Wave 1 Regional Teams and Support Staff Certification',
    group: 5,
    start: '2024-05-06',
    end: '2024-05-10',
    type: 'box',
    className: 'success'
  },
  {
    id: 21,
    content: 'Documentation Review Complete',
    group: 5,
    start: '2024-05-10',
    type: 'point',
    className: 'success'
  },

  // Release and communication items
  {
    id: 22,
    content: 'Executive Stakeholder Communication and Change Management Notification Distribution',
    group: 6,
    start: '2024-05-13',
    end: '2024-05-14',
    type: 'box',
    className: 'warning'
  },
  {
    id: 23,
    content:
      'Customer Communication Campaign Launch with Feature Announcement and Migration Timeline',
    group: 6,
    start: '2024-05-15',
    end: '2024-05-17',
    type: 'box',
    className: 'info'
  },
  {
    id: 24,
    content: 'Partner Integration Support Window and Technical Assistance Hotline Activation',
    group: 6,
    start: '2024-05-15',
    end: '2024-05-22',
    type: 'box',
    className: 'purple'
  },
  {
    id: 25,
    content: 'Feedback Collection and Customer Satisfaction Survey Distribution Phase One',
    group: 6,
    start: '2024-05-20',
    end: '2024-05-24',
    type: 'box',
    className: 'pink'
  },
  {
    id: 26,
    content:
      'Release Retrospective Meeting with Cross-Functional Teams and Lessons Learned Workshop',
    group: 6,
    start: '2024-05-27',
    end: '2024-05-28',
    type: 'box'
  },

  // Post-launch monitoring
  {
    id: 27,
    content: '30-Day Post-Launch Performance Review and Key Metrics Analysis Presentation',
    group: 6,
    start: '2024-05-28',
    type: 'point',
    className: 'success'
  },
  {
    id: 28,
    content: 'Bug Triage and Priority Assessment for Version 1.1 Patch Release Planning',
    group: 4,
    start: '2024-05-01',
    end: '2024-05-08',
    type: 'box',
    className: 'danger'
  },
  {
    id: 29,
    content: 'Database Optimization and Index Tuning Based on Production Query Analysis',
    group: 4,
    start: '2024-05-06',
    end: '2024-05-13',
    type: 'box',
    className: 'warning'
  },
  {
    id: 30,
    content: 'Capacity Planning Review for Q3 Traffic Projections and Infrastructure Budgeting',
    group: 4,
    start: '2024-05-13',
    end: '2024-05-17',
    type: 'box',
    className: 'info'
  }
];

export const groups = [
  { id: 1, content: 'Planning Phase', value: 1 },
  { id: 2, content: 'Development & Testing', value: 2 },
  { id: 3, content: 'Strategic Reviews', value: 3 },
  { id: 4, content: 'Security & Performance', value: 4 },
  { id: 5, content: 'Documentation & Training', value: 5 },
  { id: 6, content: 'Release & Communication', value: 6 }
];

export const description =
  'Long labels and text overflow - tests text handling and width estimation';
