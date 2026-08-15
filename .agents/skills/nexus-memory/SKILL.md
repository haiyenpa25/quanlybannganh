---
name: nexus-memory-graph
description: Nexus Memory Knowledge Graph skill for persistent project memory, ERD schema tracking, and component dependency maps across sessions.
---

# Nexus Memory Knowledge Graph Skill

This skill maintains project persistent memory across conversation sessions using **NEXUS_GRAPH.md**:

## Memory Architecture

1. **Entity Relationship Diagram (ERD)**: Tracks relationships between `Members`, `Groups`, `Weekly_Schedules`, `Attendance_Records`, `Visitations`, `Quarterly_Themes`, `Accounts`, `Config`.
2. **Schema Index**: Defines Google Sheet column names, types, and default fallbacks.
3. **Feature Component Map**: Maps client HTML tab IDs (`#tab-checkin`, `#tab-visitations`, `#tab-members`), JS render functions, and GAS backend endpoints.
4. **State Transition Sync**: Keeps memory updated whenever schema changes or new features are introduced.
