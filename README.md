# AIN1 Excel Pro — README

## Overview
AIN1 Excel Pro is a single-file HTML workspace for Excel-focused review and learning. The current build combines a public landing page with the main workspace and is positioned as a free tool for VBA review, formula analysis, workbook diagnostics, structured learning, and guided improvement.

## Current Build
- **Product:** AIN1 Excel Pro
- **Version:** v5.4.0.4
- **Author:** Matthew Naliponguit Sugang
- **License:** CC BY-NC-ND 4.0
- **Format:** Standalone HTML application (landing page + workspace in one file)

## What is included in this build
### 1) Marketing-first landing page
The landing page acts as the public front door before users enter the workspace.

Included sections:
- Hero / product intro
- Features
- Models
- Workflow
- Landing CTA to open the workspace
- Landing footer with credits and license

### 2) Main workspace
The workspace remains the main operating area and includes:
- VBA Code Editor with live monitoring
- Excel Formula Checker
- VBA Code Templates Library
- Excel Functions & Formula Templates
- Arithmetic Operators
- VBA Constants & Declarations
- Excel Criteria Expressions
- Excel Keyboard Shortcuts
- Conditional Formatting Library
- Absolute Cell References (F4)
- Common & Complete Errors in Excel Formulas
- Comprehensive System Self-Test
- License & About

## Major updates reflected in this version
### Landing + workspace combination
- Landing page and workspace are combined into one HTML file.
- Users can open the workspace directly from the landing page.
- The landing page is positioned as a cleaner product presentation layer before the tool opens.
- The product is presented as a **free tool**, with no pricing wall.

### Workspace navigation and UX
- Collapsible navigation panel
- Desktop and mobile/tablet navigation handling
- Mobile overlay behavior for sidebar open/close
- Workspace and landing page are designed to coexist without changing the core tool identity

### VBA Code Editor
- Live monitoring while typing
- Syntax highlighting / line numbers / code editor layout
- Code Analysis notification panel remains available
- Editor-focused workflow stays intact

### Formula Checker
- Live formula analysis
- Overlay-based formula highlighting
- Integrated syntax help / suggestions UI
- Formula Analysis notification panel remains available
- Built-in AIN1Engine formula pipeline for tokenizer, AST, rules, optimizer, explanation, dependencies, and calculation

### Code Templates improvements
- VBA Code Templates Library is included as a dedicated workspace view
- Template search box added
- Category filter added
- Empty-state handling for unmatched filters
- Template browsing is improved for larger libraries

### Reliability / analyzer improvements
- False-positive prevention was added to the VBA lint flow
- Statement structure validation is layered into the analyzer
- Multi-line `If` checking is now shape-aware instead of relying on a simplistic `If ... Then` match
- Formula analysis also uses a local deterministic engine layer for deeper interpretation

### Formula and workbook learning resources
- Excel functions reference
- Criteria expressions reference
- Operators reference
- Constants & declarations reference
- Conditional formatting library
- Error reference for common Excel formula mistakes
- Self-test utilities for internal validation

### Responsive compatibility
- Mobile / tablet responsive adjustments are present
- Sidebar behavior changes automatically for mobile vs desktop
- Landing page sections also adapt for smaller screens

## Important intentional design direction
- The build is **free access**.
- The landing page is used to explain and market the product first.
- The workspace is the actual tool area for editing, checking, browsing templates, and using reference views.
- Execution Tool-related functionality is not the focus of this combined landing/workspace direction.

## How users move through the app
1. Open the HTML file
2. Start on the landing page
3. Review Features / Models / Workflow
4. Click **Open Workspace**
5. Use the workspace tools
6. Return to the landing page when needed

## Notes for maintenance
- Keep the landing page and workspace styles aligned so they feel like one product.
- Avoid changing Code Editor logic unless needed.
- Avoid changing Formula Checker logic unless needed.
- Template search/filter should stay synchronized with the template library data source.
- Analyzer changes should prefer token-aware / context-aware checks to reduce false positives.
- Preserve the footer credits and license block.

## License / attribution
AIN1 Excel Pro © 2026 by Matthew Naliponguit Sugang is licensed under **CC BY-NC-ND 4.0**.

Unauthorized modification, resale, or redistribution is prohibited.
