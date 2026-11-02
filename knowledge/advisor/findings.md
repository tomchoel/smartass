# Review findings

Adapted from Anthropic financial-services `audit-xls` (Apache-2.0; see NOTICE.md). `core/check.py` produces this
table after every answer; `sv_answer` returns it as `review: {ok, findings: [{severity, text}]}`.

| Severity | Meaning | Raised for |
|----------|---------|------------|
| **Critical** | A number on the answer cannot be trusted. `ok` is false. | A reconciliation check that is not OK; a formula error (`#REF!`, `#VALUE!`, `#N/A`, `#DIV/0!`, `#NAME?`); a source sheet that differs from the fetch |
| **Warning** | The numbers are right, but the user must know. | A pull that fell back to the Smart View ad hoc grid; POV problems or unknown members; an answer without reconciliation checks; a hardcoded number inside a formula (formula style); a material variance without a driver; engine warnings |
| **Info** | Cosmetic or best practice. | Layout lint; an always-commented line (revenue, personnel, EBITDA) without a driver |

## How the agent uses it

- **Critical**: say so first and do not present the affected numbers as final. Fix the spec or the pull and run
  again.
- **Warning**: mention each one in the reply, in one line.
- **Info**: mention only when asked.
- Do not change the user's own cells to make a finding go away. Report first, fix on request.

## What is checked by construction

Every derived number is an Excel formula on the source sheets (no hardcoded results), the sign convention is
FinMax's (costs negative), and the source sheets carry a provenance header (cube, POV, time, route).
