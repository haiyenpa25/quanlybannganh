# Workspace Agent Rules - ManagerTeams (Superpowers + Nexus + DESIGN.md)

This workspace integrates three core frameworks to empower AI pair-programming:

## 1. Superpowers Agentic Methodology (obra/superpowers)
- **Structured Planning**: All architectural additions require research, detailed implementation plans, and user sign-off.
- **Red-Green TDD Verification**: Always run AST syntax checks (`scratch/check_syntax.js`) and build/deploy steps (`clasp push`) before claiming task completion.
- **DRY & YAGNI**: Re-use project icon generators (`getSvgIcon`), date formatters (`formatDisplayDate`), and escape utilities (`escapeHtml`).

## 2. Nexus Memory Knowledge Graph (NEXUS_GRAPH.md)
- Maintain persistent memory of entities: `Members`, `Groups`, `Weekly_Schedules`, `Attendance_Records`, `Visitations`, `Quarterly_Themes`, `Accounts`, `Config`.
- Map frontend components to backend GAS endpoints in `NEXUS_GRAPH.md`.

## 3. Google Labs DESIGN.md Specification (google-labs-code/design.md)
- Machine-readable YAML frontmatter tokens in `DESIGN.md` govern visual styles in `Styles.html`.
- Enforce iOS 18 Deep Space Glassmorphism, Floating Glass Pill Navigation Bar (`66px` height, `28px` radius, `blur(32px)`), Master-Detail expandable rows (`.master-row` + `.child-row`), and WCAG 4.5:1 contrast.
