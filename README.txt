CORNERSTONE PDK Monitor — Release 2.1.0
=======================================
CORNERSTONE Open Platform PDK · University of Southampton · June 2026

WHAT THIS IS
------------
PDK Monitor is a single-file, browser-based dashboard for managing an
open-source photonic PDK: building blocks, layouts/GDS, S-parameters,
characterisation and test data, fabrication notes, documentation, and the
Component Development Index (CDI 1a-4b) that captures each block's maturity.
It runs entirely in your browser — nothing to install, no server — and syncs
through a shared GitHub repository.

QUICK START
-----------
1. Open  cornerstone-pdk-monitor-v2.1.0.html  in a modern browser
   (Chrome, Edge, Firefox, or Safari). Double-click it, or host it on
   GitHub Pages and open the URL.
2. Accept the Terms of Use prompt (nothing is sent anywhere).
3. Browse in Read-only mode — no sign-in needed. To contribute, click the
   mode pill (top-right) and sign in with a GitHub token (user-admin) or the
   admin password.

See USER_MANUAL.pdf (or .docx / .md) for the full guide, and CHANGELOG.md for
what changed in this release.

BUNDLE CONTENTS
---------------
  cornerstone-pdk-monitor-v2.1.0.html   The dashboard application.
  cdi_calculator.html                   Standalone CDI Calculator (also built
                                        into the dashboard: Analysis -> CDI).
  CDI_BB_template_rich.xlsx             Rich BB spec/measurement template for
                                        the calculator's Excel import.
  USER_MANUAL.pdf / .docx / .md         The user manual.
  CHANGELOG.md                          Release notes for v2.1.0.
  README.txt                            This file.
  CORNERSTONE_PDK_Monitor_Introduction.pptx
                                        Introductory slide deck.
  dashboard.json                        The baked dashboard model (reference
                                        export).

PRIVACY
-------
The application collects no telemetry. Your working copy is stored in your
browser (IndexedDB / localStorage). GitHub is the source of truth for shared
content. A backup .json is downloaded before any merge or de-duplication.

CORNERSTONE Open Platform PDK — pdk.cornerstone@soton.ac.uk
