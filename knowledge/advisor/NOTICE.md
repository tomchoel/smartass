# Notice — vendored advisor text

`variance-commentary.md` in this folder is adapted from the Anthropic
**financial-services-plugins** marketplace (<https://github.com/anthropics/financial-services-plugins>):

| File | Source |
|------|--------|
| `variance-commentary.md` | `month-end-closer/skills/variance-commentary/SKILL.md` |

Copyright Anthropic, PBC. Licensed under the Apache License, Version 2.0
(<https://www.apache.org/licenses/LICENSE-2.0>). You may not use this file except in compliance with the License.
Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on
an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.

## Changes made for smartass (Telenor Norge)

- The materiality threshold is Telenor's, read from `knowledge/catalog.toml` `[materiality]`: 10 % of the absolute
  comparison value counted only from NOK 500k, or NOK 5 million; either is enough.
- The always-comment list is revenue, personnel cost and EBITDA (catalog `materiality.always`), not
  revenue/headcount cost/cash.
- A driver that is not backed by data is allowed only when visibly labelled "Hypothesis — not verified in data".
- The engine (`core/commentary.py`) flags the lines; the agent writes the drivers.
- Wording shortened and adapted to Smart View source sheets and the FinMax sign convention.
