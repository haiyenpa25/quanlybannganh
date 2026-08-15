---
name: ManagerTeams-iOS18-Dark
colors:
  primary: "#0a84ff"
  secondary: "#38bdf8"
  tertiary: "#a855f7"
  surface-base: "#090d16"
  surface-elevated: "#0f172a"
  surface-card: "rgba(22, 30, 49, 0.75)"
  text-primary: "#f1f5f9"
  text-secondary: "#94a3b8"
  success: "#10b981"
  warning: "#f59e0b"
  danger: "#ef4444"
typography:
  h1:
    fontFamily: "-apple-system, 'SF Pro Display', sans-serif"
    fontSize: "1.5rem"
    fontWeight: "800"
  h2:
    fontFamily: "-apple-system, 'SF Pro Display', sans-serif"
    fontSize: "1.2rem"
    fontWeight: "700"
  body-md:
    fontFamily: "-apple-system, 'SF Pro Display', sans-serif"
    fontSize: "0.875rem"
  label-caps:
    fontFamily: "-apple-system, 'SF Pro Display', sans-serif"
    fontSize: "0.75rem"
    fontWeight: "700"
rounded:
  sm: "10px"
  md: "14px"
  lg: "18px"
  xl: "24px"
  pill: "28px"
spacing:
  xs: "4px"
  sm: "8px"
  md: "14px"
  lg: "20px"
  xl: "28px"
---

# DESIGN.md - ManagerTeams Design System Specification (Google Labs Spec)

Official DESIGN.md specification for **ManagerTeams** adhering strictly to **@google/design.md** standards & **UI/UX Pro Max 2026**.

---

## Overview

Ultra-Clean iOS 18 Deep Space Glassmorphism meets Church Pastoral Management. The UI evokes a premium native iOS 18 experience — high contrast dark slate backdrop, glowing cyan/purple accent tokens, floating glass pill tab bar, and 44px+ touch targets.

## Color Rationale

- **Primary Accent (`#0a84ff` / `#38bdf8`)**: iOS Sky Blue for active tab badges, primary buttons, and link anchors.
- **Tertiary Glow (`#a855f7`)**: Royal Purple for special anniversaries, birthdays, and ambient background light.
- **Surface Base (`#090d16`)**: Deep space dark foundation, eliminating eye strain during evening leadership meetings.
- **Surface Card (`rgba(22, 30, 49, 0.75)`)**: Glassmorphic container with `backdrop-filter: blur(32px)` and subtle linear border.

## Component Specifications

### Floating Glass Pill Navigation Bar (`.mobile-bottom-nav`)
- Height: `66px`
- Border Radius: `28px`
- Backdrop Filter: `blur(32px) saturate(210%)`
- Border: `1px solid rgba(255, 255, 255, 0.15)`
- Shadow: `0 12px 36px rgba(0,0,0,0.6)`

### Master-Detail Expandable Drawer (`.master-row` & `.child-row`)
- Animation: `@keyframes expandSlideDown 0.2s cubic-bezier(0.4, 0, 0.2, 1)`
- Expand Button: `28px x 28px` circular toggle with smooth chevron flip.

### Accordion Mobile Cards (`.mob-card`)
- Border Radius: `20px`
- Status Indicator: `4px` left accent stripe with status-tailored color.
