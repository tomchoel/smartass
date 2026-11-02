# Drill-through API research (EPM Cloud Data Integration)

Date: 2026-11-02. Web research only. No tenant was called and nothing was logged in to.
Customer Connect and My Oracle Support (MOS) pages are behind a login or timed out. Where a claim rests only on a search snippet from such a page, it is marked "(snippet)".

## 1. Is there an official API for drill-through source lines?

- **Not found.** The Data Integration REST chapter lists 10 APIs: Get Pipeline Details, Export/Import Data Mapping, Run Integrations, Retrieve Job Status, Export/Import Data Integration, Run a Pipeline, Lock/Unlock POV and Run Data Rules. None returns drill-through or landing-page lines. https://docs.oracle.com/cloud/latest/epm-common/PREST/fdmee_rest_apis.htm
- The Planning REST API covers jobs, members, applications, planning units, data slices and substitution or user variables. It has no drill-through resource. https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/prest/planning_rest_apis.html
- **Closest official route: Data Management reports.** You `POST /aif/rest/V1/jobs` with `jobType: "REPORT"` (BI Publisher templates, output pdf/xlsx/html/excel), then download the output file.
  - Roles: **Service Administrator or Power User**.
  - Shipped templates are POV-level (TB, dimension map and so on). Oracle says custom SQL can extend them, but building that is admin work.
  - https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/prest/fdmee_execute_report.html
- The same reports run through EPM Automate `runDMReport`, with output going to `outbox/reports`. https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/cepma/epm_auto_run_dm_report.html
- **Auth for the official REST APIs:**
  - Basic auth: `identitydomain.username` with a password in the `Authorization` header. https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/prest/authentication_overview.html
  - OAuth 2: needs an OAuth client set up by a Domain Administrator. The token must carry a user context (device-code flow for SSO or SAML users). https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/prest/authentication_oath.html
  - Neither page mentions a browser session cookie or `sso_token` as accepted credentials. Our 401 on the web session matches that.
  - Basic auth also needs an identity-domain password. Purely federated SSO users usually do not have one. This is inferred, not stated on the page.
- **Smart View in-sheet drill** (results land in a new Excel sheet) exists only for **EPM Integration Agent query** drills, which an admin configures. Landing-page drills open in the browser. https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/svuuc/drill_through_to_source_102x7269afe2.html

## 2. Is the FdmeeMain.jsp URL or `/aif/ui/model/drilldown/getData` documented?

- **`FdmeeMain.jsp?module=aif.drilldown`: not found** in any Oracle doc, blog or forum I could reach.
- **`/aif/ui/model/drilldown/getData` and `/aif/ui/model/view?page=DRILL_THROUGH`: not found.**
  - No docs, blogs or GitHub code mention them. A GitHub code search for `drilldown/getData` returned only our own private repo (tomchoel/smartass) and unrelated projects.
  - The body shape and the meaning of `limit`/`offset`, `partitionkey` and `loadId` are undocumented. They are internal UI endpoints and may change in any monthly update.
- **Where the URL shape comes from:** it is an Essbase drill-through URL template that the EPM application (here Data Integration) registers on the cube.
  - The Essbase MaxL `create drillthrough` doc shows a template with the tokens `$SSO_TOKEN$`, `$CONTEXT$` and `$ATTR(ds,pos,gen,level.edge)$`. Essbase substitutes them per cell. This is the likely origin of our `sso_token=` and `attribute=<dim>.id.<member>` parameters.
  - The token grammar itself is not explained there. https://docs.oracle.com/en/database/other-databases/essbase/21/esssr/create-drillthrough.html
  - Essbase admin guide: "The URL XML is populated by the ERP or EPM applications that deployed the Essbase database… not recommended that the Administrator edit". https://docs.oracle.com/cd/E26232_01/doc.11122/esb_dbag/ch13s02s02.html
- Data Integration names the drill regions `FDMEE_<scenario member>`. Drill works only for data loaded in **Full** workflow mode, not Full No Archive or Simple. https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/diepm/integrations_application_detail_options_100x63f4cf96.html
- MOS note 1636621.1 "FDMEE: Drill-through URL Details" exists, but it is behind a MOS login and was not read. https://support.oracle.com/knowledge/Enterprise%20Performance%20Management%20and%20Business%20Intelligence/1636621_1.html
- Our own repo already has a verified finding: `sso_token` alone does not authenticate, and the IDCS gate redirects to sign-in (docs/archive/specs/2026-10-15-sv-drill-design.md). No external source contradicts it.

## 3. EPMFDM-140373 ("too wide") and drill limits

- **The code EPMFDM-140373 itself: not found** on the public web. Its text, as captured in our logs, is "more than 3 dimensions have more than 1000 child members". That matches Oracle's documented summary-drill rules:
  - "Drill is limited to 1000 descendant members for a dimension". The limit applies to each dimension. https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/diepm/integrations_defining_drill_details.html
  - The application option **"Summary Drill Behavior when more than 1000 descendants"** has two settings:
    - **Ignore:** drops that dimension from the filter. "**A maximum of three dimensions can be ignored.**"
    - **Limit:** uses only the first 1,000 members, so results can be incomplete.
    - Source: the Application Details options page linked in section 2.
  - **Inference:** EPMFDM-140373 fires when Ignore is set and a fourth dimension exceeds 1,000 descendants. Oracle does not state this mapping explicitly.
  - Under Ignore, a summary drill can return rows that are *not* under the parent you drilled from. It is only correct when the ignored dimensions meet unique members of the other dimensions. A parent with more than 1,000 descendants "can result inconsistent results" (Defining Drill Details).
  - Oracle recommends enabling drill only from parents with fewer than 1,000 descendants (custom drill region, admin).
- **Workarounds (user side):** drill at lower levels so that at most 3 dimensions are over 1,000 descendants (split by child entity or account, or go to level 0). This is what we already do.
  - Also check that the drilled sum ties to the cube: under Ignore it may be too wide, and under Limit too narrow.
  - Admin-side options: switch Ignore/Limit, or narrow the drill region in Application Drill Details.
- **Row limits on the landing page / getData:** not found in any doc. The `limit=500` is the UI's page size and is not a documented cap.
- A Customer Connect post says "Drill through works only if we have less than 1000 dimension members" (snippet). https://community.oracle.com/CustomerConnect/discussion/512288/drill-through-feature-in-epbcs
- Another post reports that drill totals did not match the form when a dimension had more than 1,000 descendants (snippet). https://community.oracle.com/CustomerConnect/discussion/840474/drill-through-source-data-that-is-displayed-does-not-equal-to-the-amount-displayed-on-data-form
- Unrelated but similar: Essbase's own relational drill-through has a 1,000-member SQL `IN` cap (`dtr.in.clause.limit`). This is on-prem Essbase only. https://docs.oracle.com/en/database/other-databases/essbase/21/essoa/expand-limit-sql-clauses-drill-reports.html

## 4. Speeding up many parallel Playwright pages (fallback)

- **Block heavy resources** with `context.route("**/*", ...)` and abort when `request.resource_type` is image, font or media.
  - Blocking stylesheets often works but can break SPAs that measure layout. Never block script, xhr or fetch.
  - Sources: https://playwright.dev/docs/network , https://www.zenrows.com/blog/blocking-resources-in-playwright , https://scrapfly.io/blog/answers/how-to-block-resources-in-playwright
  - Set the route on the context rather than each page, so every new tab inherits it.
- **One browser, few contexts, many pages.** Contexts are cheap and pages in one context share the SSO cookies. Do not start a browser per drill.
  - Reuse a warm tab by navigating it to the next drill URL. The JS and CSS are then cached, which beats opening a new tab.
  - Sources: https://playwright.dev/docs/browser-contexts , https://webscraping.ai/faq/playwright/what-are-the-different-browser-contexts-and-when-should-i-use-them
- **Do not wait for full load.** Use `wait_until="commit"` (or `"domcontentloaded"`) and resolve as soon as the `getData` request is seen (`page.wait_for_request` / `expect_request`), then close or reuse the tab. No public source quantifies the gain for this SPA.
- **Cap concurrency.** Our numbers (~5 s alone vs ~26 s each in parallel) point to server-side or session serialisation, or client CPU, not network. Measure throughput at 2, 4 and 6 pages and pick the knee. No Oracle guidance on per-session concurrency was found.
- **Smaller rendering cost:** use a small viewport, headless mode, `service_workers="block"` on the context, and the Chromium flag `--disable-gpu`. These help somewhat but are no substitute for skipping the page.
- **Best fallback, needs verification on our tenant:** once one drill page has loaded, call `getData` for *other* cells directly with `context.request` (it shares cookies).
  - This only works if the getData body carries the cell attributes. If the JSP load instead registers a server-side drill context (partitionkey, loadId) that getData depends on, it will not.
  - That is a hypothesis. No external source covers it.

## Implications for us

1. There is no documented, non-admin API for drill lines. Official routes need OAuth or an admin (DM `REPORT` jobs with Service Administrator or Power User role, or an Integration Agent drill). Our browser-session replay of the internal `getData` stays the only path, and it is undocumented and can break with any monthly EPM update.
2. EPMFDM-140373 matches the documented "Ignore, at most 3 dimensions over 1000 descendants" rule. Splitting to children is the right workaround, and we should check that drilled sums tie to the cube, because Ignore can widen and Limit can truncate.
3. Biggest speed win to test: skip the page render for each drill (reuse one warm session and call getData directly if its body is self-contained). Otherwise block images and fonts, wait for the getData request instead of full load, reuse tabs and tune concurrency empirically.
