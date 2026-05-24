# CLAUDE_ACTUARIAL.md

Behavioral guidelines to reduce common spreadsheet, coding, and modeling mistakes in actuarial work. Merge with peer-review checklists and regulatory (e.g., HKRBC/IFRS) instructions as needed.

**Tradeoff:** These guidelines bias toward transparency, strict auditability, and absolute numerical accuracy over speed, complex automation, or over-engineered architectures.

## 1. Think Before Modeling

**Don't assume. Don't hide confusion. Surface methodology tradeoffs.**

Before modifying a model, setting assumptions, or writing code:
- **State data and methodology assumptions explicitly.** If data is missing or corrupted, flag it immediately; do not interpolate silently.
- **Surface interpretations:** If an insurance product text or regulatory guideline has multiple interpretations, present them to peer review—do not choose a path in secret.
- **Push back on complexity:** If a simpler, linear calculation approach can achieve the same precision as a complex stochastic simulation, pitch the simpler approach first.

## 2. Simplicity & Auditability First

**Minimum logic that solves the business problem. Nothing speculative.**

- **No "Ghost" Features:** Do not build multi-scenario capabilities, product features, or stress tests unless explicitly requested.
- **No Abstractions for Single Runs:** Write explicit, cell-by-cell or step-by-step se
- **Reject Over-Engineering:** Do not use deeply nested IF statements, overly complex array formulas, or obscure VBA macros when a simple helper column or standard lookup functions (XLOOKUP, INDEX/MATCH) can solve the problem.
- **The Audit Test:** "Could an external auditor or a peer actuary trace this workbook from raw policy data to the final reserve summary without an explanation document?" If the answer is no, simplify the model structure.

## 3. Surgical Changes & Model Parity

**Touch only what you must. Clean up only the errors your changes introduce.**
When building new pricing models, reserving sheets, or automated scripts:
- **Centralized Assumptions:** Keep your underlying assumptions (interest rates, mortality tables, lapse rates) in one clearly marked, dedicated input section. Never hardcode numbers directly inside calculation formulas.
- **Direct Connections:** Keep the path between raw data inputs and final actuarial summaries as short as possible. Avoid fragmenting a single calculation across dozens of hidden or intermediary tabs.

When editing existing pricing models, reserving sheets, or automated scripts:
- **Zero Collateral Drift:** Do not "improve" adjacent cell formulas, format sheets, or rewrite legacy macros that are not broken. 
- **Match the Existing Pattern:** Match the model’s established style, variable naming conventions, and column layouts perfectly, even if you prefer a different structure.
- **Orphan Cleanup:** If your update changes a dynamic risk factor or a table loop, remove any input parameters or helper cells that *your* changes rendered useless. Do not touch pre-existing dead logic.

## 4. Goal-Driven & Balanced Verification

**Define explicit control totals. Loop until numbers reconcile perfectly.**

Transform actuarial tasks into verifiable numerical checkpoints before considering a task complete:
- **"Quantify Estimate"** $\rightarrow$ "isolate and quantify the exact financial impact where it does."
- **"Fix data query bug"** $\rightarrow$ "Tie final output record counts back to the raw policy level data feed control totals."
- **Build Self-Reconciling Blocks:** Every model update must include explicit error-checking cells (e.g., `[Total Assets] - [Total Liabilities + Surplus] = 0`). A model is not complete because it runs; it is complete when it balances to the penny.

## 5. Explicit State & Formulas

**Don't hide the mechanics. Surface the raw actuarial math.**

- **Transparent Logic:** Actuarial math (like discounting cash flows or calculating reserves) should be laid out step-by-step. Break long formulas into sequential, logical helper columns rather than condensing everything into a single, unreadable cell.

- **Explicit Over Hidden:** Favor visible, auditable cell mechanics over hidden processes. Avoid using heavy background VBA scripts to manipulate data behind the scenes when a visible formula structure can do it transparently.