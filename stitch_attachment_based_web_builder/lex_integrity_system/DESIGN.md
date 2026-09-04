---
name: Lex Integrity System
colors:
  surface: '#f9f9ff'
  surface-dim: '#d9d9e0'
  surface-bright: '#f9f9ff'
  surface-container-lowest: '#ffffff'
  surface-container-low: '#f3f3fa'
  surface-container: '#ededf4'
  surface-container-high: '#e8e8ee'
  surface-container-highest: '#e2e2e8'
  on-surface: '#1a1c20'
  on-surface-variant: '#424751'
  inverse-surface: '#2e3035'
  inverse-on-surface: '#f0f0f7'
  outline: '#737782'
  outline-variant: '#c3c6d2'
  surface-tint: '#2e5ea5'
  primary: '#003874'
  on-primary: '#ffffff'
  primary-container: '#1a4f95'
  on-primary-container: '#a3c3ff'
  inverse-primary: '#aac7ff'
  secondary: '#006d42'
  on-secondary: '#ffffff'
  secondary-container: '#9af6be'
  on-secondary-container: '#087347'
  tertiary: '#5e2b00'
  on-tertiary: '#ffffff'
  tertiary-container: '#813d00'
  on-tertiary-container: '#ffb17e'
  error: '#ba1a1a'
  on-error: '#ffffff'
  error-container: '#ffdad6'
  on-error-container: '#93000a'
  primary-fixed: '#d6e3ff'
  primary-fixed-dim: '#aac7ff'
  on-primary-fixed: '#001b3e'
  on-primary-fixed-variant: '#08458b'
  secondary-fixed: '#9af6be'
  secondary-fixed-dim: '#7ed9a3'
  on-secondary-fixed: '#002110'
  on-secondary-fixed-variant: '#005230'
  tertiary-fixed: '#ffdbc7'
  tertiary-fixed-dim: '#ffb688'
  on-tertiary-fixed: '#311300'
  on-tertiary-fixed-variant: '#733600'
  background: '#f9f9ff'
  on-background: '#1a1c20'
  surface-variant: '#e2e2e8'
  status-notice: '#007AFF'
  status-enactment: '#10B981'
  background-subtle: '#F8FAFC'
  border-neutral: '#E2E8F0'
typography:
  headline-lg:
    fontFamily: Hanken Grotesk
    fontSize: 32px
    fontWeight: '700'
    lineHeight: 40px
    letterSpacing: -0.02em
  headline-md:
    fontFamily: Hanken Grotesk
    fontSize: 24px
    fontWeight: '600'
    lineHeight: 32px
    letterSpacing: -0.01em
  headline-sm:
    fontFamily: Hanken Grotesk
    fontSize: 20px
    fontWeight: '600'
    lineHeight: 28px
  body-lg:
    fontFamily: Inter
    fontSize: 18px
    fontWeight: '400'
    lineHeight: 28px
  body-md:
    fontFamily: Inter
    fontSize: 16px
    fontWeight: '400'
    lineHeight: 24px
  body-sm:
    fontFamily: Inter
    fontSize: 14px
    fontWeight: '400'
    lineHeight: 20px
  label-md:
    fontFamily: Inter
    fontSize: 13px
    fontWeight: '600'
    lineHeight: 16px
  code-sm:
    fontFamily: JetBrains Mono
    fontSize: 12px
    fontWeight: '400'
    lineHeight: 16px
rounded:
  sm: 0.125rem
  DEFAULT: 0.25rem
  md: 0.375rem
  lg: 0.5rem
  xl: 0.75rem
  full: 9999px
spacing:
  base: 4px
  xs: 4px
  sm: 8px
  md: 16px
  lg: 24px
  xl: 40px
  sidebar-width: 280px
  container-max: 1200px
---

## Brand & Style
The design system is engineered for a legal and public data service, where **trust, authority, and clarity** are paramount. The target audience includes legal professionals, environmental officers, and corporate compliance teams who require high-density information presented without distraction.

The chosen style is **Corporate / Modern** with a focus on **Information Density**. It leverages a structured grid, generous whitespace to separate complex data points, and a professional aesthetic that mirrors the seriousness of legislative data. The interface should feel like a reliable tool: objective, precise, and highly functional.

## Colors
This design system utilizes a palette rooted in institutional reliability and environmental awareness.
- **Primary (Blue):** A deep, authoritative blue representing the law and official governance. Use this for primary actions, navigation headers, and brand elements.
- **Secondary (Green):** A balanced forest green representing the environmental focus of the data. Use this for category badges related to environmental law and positive action states.
- **Neutral:** A range of cool grays (Slate) to provide a clean, "paper-like" digital environment that minimizes eye strain during long reading sessions.
- **Semantic Colors:** Blue is used for "Legislative Notice" (입법예고) and Green for "Enactment/Revision" (개정) to provide instant visual categorization in the list view.

## Typography
The typography strategy prioritizes legibility for complex Korean legal text.
- **Hanken Grotesk** is used for headlines to provide a modern, sharp, and professional character.
- **Inter** is the workhorse for body text, chosen for its exceptional readability in data-heavy environments and its neutral tone.
- **JetBrains Mono** is used sparingly for metadata (IDs, dates, and status codes) to give a sense of technical precision and data integrity.

For mobile-specific views, `headline-lg` should scale down to 24px and `headline-md` to 20px to ensure long legal titles do not break the layout.

## Layout & Spacing
The layout uses a **Fixed Grid** model for the main content area to maintain an editorial feel, while the sidebar remains fixed to the left.

- **Grid:** A 12-column grid is used for the main content area with 24px gutters.
- **Sidebar:** A 280px fixed-width sidebar houses the navigation and primary filters.
- **Desktop:** Content is centered with a max-width of 1200px.
- **Reflow:** On smaller screens, the sidebar collapses into a drawer menu (hamburger), and the 12-column grid collapses into a single-column stack for the list items.
- **Rhythm:** Use multiples of 4px for all padding and margins to ensure a tight, mathematical alignment.

## Elevation & Depth
To maintain a "clean and trustworthy" feel, this design system avoids heavy shadows. 
- **Low-contrast outlines:** Primary depth is created using 1px borders in `border-neutral` (#E2E8F0) rather than shadows.
- **Tonal layers:** The background uses a subtle off-white (#F8FAFC), while cards and main content areas use pure white (#FFFFFF) to create a "lifted" effect without visual clutter.
- **Hover States:** Subtle ambient shadows (0px 4px 12px rgba(0,0,0,0.05)) are used only for interactive cards to signify clickability.

## Shapes
A **Soft (0.25rem)** roundedness is applied to most components. This subtle rounding softens the "harshness" of legal data while maintaining a professional, structured appearance. 
- Buttons and input fields use the standard 0.25rem (4px) radius.
- Large containers and cards use `rounded-lg` (8px).
- Status tags/chips use a full pill-shape to distinguish them from interactive buttons.

## Components
- **Buttons:** Primary buttons use the `primary_color_hex` with white text. Ghost buttons (border only) are used for secondary actions like "Add Recipient."
- **Status Chips:** Small, non-interactive tags for "Legislative Notice" (Blue tint) and "Enactment" (Green tint). They use a semi-transparent background of the brand color with full-opacity text for high contrast.
- **Cards:** Legislative items are housed in cards with a 1px border. The title is always `headline-sm`, and metadata (Agency, Date) is set in `body-sm`.
- **Input Fields:** Clean, rectangular fields with a 1px border. Focus states use a 2px `primary_color` border and a soft glow.
- **Lists:** Data rows use a subtle zebra-striping or 1px bottom border to separate items clearly.
- **Mailing Dashboard:** Use a split-view layout where the recipient list is on the right and the email composition/summary is on the left to facilitate a "check-and-send" workflow.