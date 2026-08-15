---
name: google-design-md
description: Google Labs DESIGN.md specification skill for machine-readable design token parsing, YAML frontmatter validation, and aesthetic UI enforcement.
---

# Google Labs DESIGN.md Specification Skill

This skill enforces visual identity consistency adhering to **google-labs-code/design.md**:

## Design Token Rules

1. **YAML Frontmatter Structure**: `DESIGN.md` MUST specify `name`, `colors` (primary, secondary, tertiary, surface-*, text-*), `typography` (h1, h2, body-md, label-caps), `rounded`, and `spacing`.
2. **Design Tokens Over Ad-Hoc Utilities**: All CSS tokens in `Styles.html` MUST align with tokens declared in `DESIGN.md`.
3. **Contrast Verification**: WCAG 4.5:1 text-to-background ratio on dark surfaces.
4. **Touch Target Enforcement**: 44px min touch target size for mobile thumb navigation.
