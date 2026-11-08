# Evaluation of the independent PEPA Management Dash review

*Checked 8 Nov 2026 against `pbix/PEPA Management Dash.pbix` and its extracted metadata in `docs/models/_extracted/`. The review being evaluated is [`PEPA_Management_Dash_Change_List.xlsx`](PEPA_Management_Dash_Change_List.xlsx), which reviewed `1472405d-PEPA_Management_Dash.pbix` on 6 Nov 2026. This document covers its 10 High-severity findings. The other 38 (Medium, Low, Info) are not evaluated here.*

## Summary

| ID | Finding | Verdict | Severity: review → this evaluation | Also affects | Already in our docs? |
|---|---|---|---|---|---|
| DX-01 | `Hours` drops the last selected day | **Confirmed** (one sub-point doesn't apply) | High → High | Compliance Monitoring Dash (7 visuals) | DAX shown on the model page, not flagged |
| DX-02 | `Hours2` has the same defect, plus zero-fill | **Confirmed** | High → High | PEPA Exec Dash (6 visuals); Compliance Monitoring Dash (1 visual) | DAX shown, not flagged |
| DX-04 | `TotalCost` is row-by-row with hard-coded numbers | **Confirmed**; the missing-salary part depends on PQ-08 | High → High | PEPA Exec Dash (same pattern, 15 visuals) | DAX shown, not flagged |
| DX-05 | `ExpectedHours` overstated, so `Utilisation` understated | **Confirmed in code**; holiday end date not checked | High → High | Compliance Monitoring Dash (measure exists, no visual) | DAX shown, not flagged |
| PQ-04 | Assess queries multiply rows | **Confirmed in code**; row counts not checked | High → **Medium** | Identical queries in 3 other models | M shown; unused columns listed in `gaps.md` |
| PQ-07 | Pasted holiday list stops at 2025 | **Partly confirmed**: pasted table confirmed, end date not checked | High → **Medium** | Same table in 6 models, used only here | `gaps.md` flags the Holidays source for 2 other models |
| PQ-08 | Role misspelt in `Salaries` | **Mechanism confirmed**; the misspelling itself not checked | High → High if confirmed | `Salaries` loaded in 3 other models | No |
| SM-01 | `CalcStage → StageOrder` keys never match | **Relationship confirmed unused**; key mismatch likely, not checked | High → **Low** | Same relationship in Compliance Monitoring Dash | Relationship listed, not flagged |
| SM-02 | Person IDs 1 and 18 missing from `RoleTable` | **Mechanism confirmed**; the IDs not checked | High → High if confirmed | — | Relationships listed |
| SM-03 | Application UIDs missing from `Applications` | **Mechanism confirmed**; counts not checked | High → **Medium** | Likely worse in Compliance Monitoring Dash | Relationships listed |

- **The four DAX findings matter most.** The measure code alone confirms them, and they change numbers on visible pages. They also reach **PEPA Exec Dash** and **Compliance Monitoring Dash**, which the review didn't cover.
- **Six findings depend on specific values:** holiday dates, role spellings, orphan keys and row counts. This evaluation didn't read row data (see [Method](#method-and-limits)), so those parts stay open. The queries under [Checks to close the open points](#checks-to-close-the-open-points) settle them in a few minutes in Power BI Desktop.
- **Three severities look too high:**
  - PQ-04: the review itself says results are unchanged.
  - SM-01: an inactive relationship that nothing activates can't change any number.
  - SM-03: by the review's own counts, it affects about 1% of application UIDs.

## Method and limits

**What was read:**
- DAX measures and calculated columns, Power Query (M) and relationships, from `docs/models/_extracted/PEPA Management Dash.json`.
- Which visuals use which fields, from the report definition inside the `.pbix`.
- The model's storage statistics: distinct-value counts per column and refresh times. These are metadata, not cell values.

**What was not read:** row data. CLAUDE.md forbids reading it, mainly because `Salaries` holds pay data. Claims that rest on specific values are marked "not checked".

**Snapshot difference:** our copy was last refreshed 2026-10-27 21:18, the time stored in the model. The review used a copy saved later, and its figures differ slightly from ours:

| Figure | Review | Our copy |
|---|---|---|
| Distinct `StartTime` values | 18,496 | 18,247 |
| Holiday dates | 161 | 159 distinct (the review may count 2 repeated dates) |
| Last timesheet date | 2026-11-05 | after our refresh |

None of the problems below depend on which snapshot is used.

**Verdicts:**
- **Confirmed:** the code or metadata proves the claim.
- **Confirmed in code / mechanism confirmed:** the defect exists in the logic, but how big it is depends on values that weren't checked.
- **Partly confirmed:** some of the claim holds, and the rest wasn't checked or doesn't apply.

## Findings

### DX-01 · `TimesheetData[Hours]` · Confirmed

- **The measure:** `CALCULATE(SUM(TimesheetData[Minutes]), FILTER(TimesheetData, TimesheetData[StartTime] >= MIN('DateSlicer'[Date]) && TimesheetData[StartTime] <= MAX('DateSlicer'[Date]))) / 60`.
- **Slicer dates are all midnight.** `DateSlicer` is `SELECTCOLUMNS(DateTable, "Date", DateTable[Date])`, and `DateTable` is `CALENDAR(DATE(2020,1,1), DATE(2030,12,31))`.
- **Timesheet entries have a time of day.** `StartTime` has 18,247 distinct values, while the date-only `WorkDate` has 433. So any entry after 00:00 on the last selected day fails `<= MAX(...)` and is dropped. A one-day selection returns blank.
- **The slicer only works through these measures.** `DateSlicer` has no relationships, so MIN and MAX are the only way it reaches the numbers.
- **Visual count matches the review:** 15 visuals use `[Hours]`, across Page 1, Application Time, Timesheet, Complaint & Audit Time, Team, Effort on Application and Duplicate of Timesheet.
- **Correction:** the "non-contiguous selection" point doesn't apply as the report is built. All three `DateSlicer` slicers (Application Time, Complaint & Audit Time, Team) are relative-date slicers, which always select a continuous range. With a relative slicer, the day dropped is usually today, or the last day of the last full period.
- **Also affects:**
  - Compliance Monitoring Dash: identical measure on 7 visuals (Timesheet, Complaint & Audit Time, Time on Job, StandUp).
  - ELT Frontpage: identical measure, but not on any visual.
- **Fix direction:** compare dates rather than datetimes, for example `TimesheetData[WorkDate] <= MAX('DateSlicer'[Date])`. The review's "DAX Fixes" sheet takes this approach.

### DX-02 · `TimesheetData[Hours2]` · Confirmed

- **Same filter as `Hours`,** wrapped in `DIVIDE(..., 60)` and `COALESCE(..., 0)`, so it has the same end-day defect.
- **Visual count matches the review:** 8 visuals (5 cards and 1 table on Team, plus one each on Application Time and Complaint & Audit Time).
- **The zero-fill effect is real.** `COALESCE` to 0 means the measure is never blank, so tables and charts list every person or category, including those with no hours.
- **Also affects:**
  - PEPA Exec Dash: identical measure on 6 visuals on Main, and it feeds that report's `TotalCost`.
  - Compliance Monitoring Dash: 1 visual.
  - ELT Frontpage: not on any visual.

### DX-04 · `TimesheetData[TotalCost]` · Confirmed (the missing-salary part depends on PQ-08)

- **The measure:** `SUMX(TimesheetData, [Hours] * (LOOKUPVALUE(Salaries[Salary], Salaries[Role], RELATED(RoleTable[Role])) / 1816) * 1.65)`, wrapped in `COALESCE(Cost, 0)`.
- **It is slow.**
  - Calling `[Hours]` inside `SUMX` turns each row into a filter (context transition), and `[Hours]` then runs its own `FILTER` over the whole table for every row.
  - The salary `LOOKUPVALUE` also runs once per row.
- **The constants are hard-coded:** 1816 and 1.65 are in the measure. It also inherits DX-01's end-day defect.
- **The duplicate-row risk is real DAX behaviour.** The row filter covers every column value, so N identical rows each see all N copies and are counted N×N times. Whether duplicate rows exist wasn't checked.
- **Visual count matches the review:** 5 visuals, all on Team.
- **Also affects:**
  - PEPA Exec Dash: its `TotalCost` uses the same pattern with `[Hours2]`, on 15 visuals on Main, an executive-facing page.
  - ELT Frontpage: its `TotalCost` is a different formula, fixed to May 2025 by hard-coded dates. It isn't on any visual.

### DX-05 · `RoleTable[ExpectedHours]` · Confirmed in code; holiday end date not checked

- **The measure:** `CALCULATE(COUNTROWS(FILTER(CROSSJOIN(VALUES(RoleTable[ID]), VALUES(DateSlicer[Date])), WEEKDAY(DateSlicer[Date], 2) <= 5 && NOT DateSlicer[Date] IN VALUES(Holidays[ObservedHoliday]))) * 8)`. It counts 8 hours per weekday, skips dates in `Holidays`, and the outer `CALCULATE` has no filter arguments, so it does nothing.
- **Counting every ID is the bigger problem.**
  - `RoleTable[ID]` has 75 distinct values, but only 49 people appear in `TimesheetData[Person]`, across all dates.
  - When all IDs are in scope, the total's expected hours can be about 50% higher than for the people actually logging time. That pulls the total `Utilisation` down far more than the missing holidays do.
- **Holiday end date (2025-12-31): not checked.** The holiday rows are embedded in the query (see PQ-07).
- **SM-02 would make it worse.** If timesheet person IDs are missing from `RoleTable`, `VALUES(RoleTable[ID])` also returns a blank member through the active `TimesheetData[Person]` relationship, which adds one phantom person.
- **Where it shows:** `Utilisation` (`IF(ISBLANK([Hours]/[ExpectedHours]), 0, [Hours]/[ExpectedHours])`) is on 1 visual, on Team. It also inherits DX-01 through `[Hours]`.
- **Also affects:** Compliance Monitoring Dash has identical `ExpectedHours` and `Utilisation`, but not on any visual.

### PQ-04 · `AssessOnDecisionDay` / `AssessLastFlagExport` · Confirmed in code; row counts not checked · High → Medium

- **The steps exist as named:**
  - `Custom4` builds `AssessorsWithIndex`, `Custom5` builds `MaxIndex`, and `Custom6` builds `AssessorsWithLastFlag`.
  - They are followed by `Expanded AssessorsWithLastFlag`, `Filtered Rows2` (`each true`), `Changed Type`, and `Added Custom`, which creates `CalcStage`.
- **The rows do multiply.** By `Custom4`, each row is one timesheet entry that still carries its Application+Day group. `Custom4` builds that group's Assess rows, and the expand step creates one row per Assess row.
- **The extra columns aren't used:**
  - `AssessorsWithLastFlag.*`, `AssessorsWithIndex` and `MaxIndex` are dropped by the final `Expanded GroupedRows` step.
  - The M `CalcStage` is loaded as `GroupedRows.CalcStage`, but no visual, relationship or DAX uses it. `gaps.md` lists it among the 16 of 20 unreferenced `AssessOnDecisionDay` columns.
- **`AssessLastFlagExport` is not de-duplicated.** It only selects, renames and adds flag columns, has no `Table.Distinct`, and ends with another `each true` filter.
- **Results are unchanged but refresh is slower.**
  - `TimesheetData[IsLastAssess]` and `[IsDeciderPresent]` read this table with `MAXX(FILTER(AssessLastFlagExport, ...))`, so duplicates don't change their values.
  - Those calculated columns scan the whole table once per timesheet row at every refresh.
- **Row counts (15,307 in both tables): not checked.** The equal counts are expected anyway, because `AssessLastFlagExport` keeps every row of `AssessOnDecisionDay`, so they don't show inflation on their own.
- **Also affects:** the same M runs in Compliance Monitoring Dash, ELT Frontpage and PEPA Exec Dash, so the refresh cost is paid four times.
- **Severity:** Medium. It costs refresh time and memory but produces no wrong numbers, which the review's own impact column also says.

### PQ-07 · `Holidays` · Partly confirmed · High → Medium

- **Confirmed: it's a pasted ("Enter Data") table** with its rows embedded in the query, so it never updates by itself.
- **Identical in all six models that load it,** with the same Power Query and 159 distinct dates in each.
- **End date (2025-12-31): not checked.**
- **Its only consumer is `ExpectedHours` (DX-05).**
  - `Holidays[ObservedHoliday]` isn't referenced in ELT Frontpage, ELT Permit Timeliness Report, PEPA Exec Dash or Permit SPE Dash.
  - According to the Flow Register, application working-day counts come from Power Automate flows that read holidays from SharePoint, not from this table.
- **Fix direction:** the Dashboards Master File says Holidays should come from the `Authorisations - PublicHolidayDates` SharePoint list (`gaps.md` → "Source differs from Master File"). Pointing the query at that list would fix this permanently.
- **Severity:** Medium. It only affects the one `Utilisation` visual.

### PQ-08 · `Salaries` role spelling · Mechanism confirmed; misspelling not checked

- **`Salaries` is a pasted ("Enter Data") table.**
- **The role name links staff to `Salaries` in three places, all by exact text:**
  1. The `TimesheetData` query step `Merged Queries2`, a left join on `RoleTable.Role = Salaries.Role`.
  2. The model relationship `RoleTable[Role] → Salaries[Role]` (many-to-one, both directions, active).
  3. The lookups `TimesheetData[PrimaryFunction] = LOOKUPVALUE(Salaries[PrimaryFunction], Salaries[Role], …)` and the one inside `TotalCost`.
- **The consequences are visible in the code:**
  - A blank `PrimaryFunction` falls through `CalcStage`'s `SWITCH` to its catch-all, `"4. Notify"`.
  - `AssessOnDecisionDay` keeps only rows whose function is Assess or Decide, so those people drop out of it.
  - The salary lookup returns blank, so their cost is 0.
- **The counts fit but don't prove it.** `RoleTable[Role]` and `Salaries[Role]` each have 24 distinct values, which matches the review's picture of 22 shared roles plus 2 unmatched on each side.
- **Not checked:** the spelling itself. Reading `Salaries` rows is barred; the check below compares role names only.
- **Also affects:** `Salaries` is also loaded in Compliance Monitoring Dash, ELT Frontpage and PEPA Exec Dash. If it's the same pasted table, they have the same problem (not checked).

### SM-01 · `TimesheetData[CalcStage] → StageOrder[Stage]` · Relationship confirmed unused · High → Low

- **The relationship is as described:** many-to-one, both directions, inactive.
- **It is never switched on.** The model's only `USERELATIONSHIP` is in the `Assigned Person` measure, for `Applications[AssignedTo]`. So this relationship can't filter or sort anything, whatever the key values are.
- **The key mismatch is likely but not checked.**
  - `CalcStage` (DAX) returns exactly `"1. Review"`, `"2. Assess"`, `"3. Decide"`, `"4. Notify"` and `"5. Monitor"`.
  - `StageOrder[Stage]` has 5 distinct values and is the target of the active `Applications[Stage]` relationship. That suggests it holds unnumbered names like `Applications[Stage]`, but its rows are embedded and weren't read.
- **Nothing relies on `StageOrder` for sorting.**
  - `CalcStage` is on 26 visuals, while `StageOrder[Stage]` is on only 1 (Application Status) and `StageOrder[Number]` on none.
  - The numbered labels already sort correctly on their own.
- **Also affects:** the identical relationship exists in Compliance Monitoring Dash.
- **Severity:** Low. It's clutter to remove, and it doesn't change any number.

### SM-02 · Person IDs missing from `RoleTable` · Mechanism confirmed; IDs not checked

- **The three relationships:**

  | Relationship | Active | Filter direction |
  |---|---|---|
  | `TimesheetData[Person] → RoleTable[ID]` | Yes | Single |
  | `Assignments[AssignedTo] → RoleTable[ID]` | Yes | Single |
  | `Applications[AssignedTo] → RoleTable[ID]` | No (used only by `Assigned Person`) | Both |

  `RoleTable[ID]` has 75 distinct values, as the review says.
- **What a missing ID causes, from the code:**
  - The person shows as blank.
  - The `TimesheetData` query joins `RoleTable` on `Person = ID`, so the role is null and there's no `Salaries` match. As in PQ-08, that means cost 0 and stage `"4. Notify"`.
  - `ExpectedHours` gains a phantom blank person (see DX-05).
- **For `Applications[AssignedTo]`, only the `Assigned Person` measure is affected,** because its relationship is inactive.
- **Not checked:** whether IDs 1 and 18 really appear.

### SM-03 · Application UIDs missing from `Applications` · Mechanism confirmed; counts not checked · High → Medium

- **The four relationships are all active:**

  | Relationship | Filter direction |
  |---|---|
  | `TimesheetData[AppUID] → Applications[UID]` | Both |
  | `Commentary[AppUID] → Applications[UID]` | Single |
  | `Assignments[AppUID] → Applications[UID]` | Single |
  | `TimesheetData[AppUID] → Complaints&Audits[AppUID]` | Single |

- **The review's "117 complaint/audit references (expected)" is right.** The last relationship confirms that `AppUID` holds both application and complaint/audit references.
- **Archiving doesn't explain the missing UIDs in this report.** Here `Applications` already appends `ApplicationsArchive`. Two leads:
  - The Flow Register says the `RemoveDeleted` flow deletes `Applications` items that have no matching regime record (an inferred link, not confirmed).
  - The UIDs may have been typed wrongly in the timesheet or assignment.
- **It is likely worse in Compliance Monitoring Dash.** There, `Applications` does **not** append `ApplicationsArchive`, but `TimesheetData` does include `TimesheetArchive`. So timesheets for archived applications will have no match.
- **Scale:** `TimesheetData[AppUID]` has 3,448 distinct values and `Applications[UID]` has 3,520. The review's 41 missing UIDs are about 1% of timesheet UIDs.
- **Severity:** Medium. Effort for those applications can't be attributed to a regime or stage, but the share is small.

## Checks to close the open points

Run these in Power BI Desktop, in DAX query view, on PEPA Management Dash. They read values inside Desktop only and export nothing, and none of them returns salary amounts.

```dax
// PQ-07 / DX-05: last holiday date, row count vs distinct dates
EVALUATE ROW("Last holiday", MAX(Holidays[ObservedHoliday]), "Rows", COUNTROWS(Holidays), "Distinct dates", DISTINCTCOUNT(Holidays[ObservedHoliday]))

// PQ-08: role names with no match on the other side (role names only)
EVALUATE EXCEPT(VALUES(RoleTable[Role]), VALUES(Salaries[Role]))
EVALUATE EXCEPT(VALUES(Salaries[Role]), VALUES(RoleTable[Role]))

// SM-01: CalcStage labels with no StageOrder row
EVALUATE FILTER(VALUES(TimesheetData[CalcStage]), NOT TimesheetData[CalcStage] IN VALUES(StageOrder[Stage]))

// SM-02: person IDs missing from RoleTable
EVALUATE FILTER(VALUES(TimesheetData[Person]), NOT TimesheetData[Person] IN VALUES(RoleTable[ID]))
EVALUATE FILTER(VALUES(Assignments[AssignedTo]), NOT Assignments[AssignedTo] IN VALUES(RoleTable[ID]))

// SM-03: timesheet UIDs found in neither Applications nor Complaints&Audits
EVALUATE FILTER(VALUES(TimesheetData[AppUID]),
    NOT TimesheetData[AppUID] IN VALUES(Applications[UID]) && NOT TimesheetData[AppUID] IN VALUES('Complaints&Audits'[AppUID]))

// DX-01: entries at midnight; DX-04: duplicate timesheet rows; PQ-04: row multiplication
EVALUATE ROW(
    "StartTime at 00:00", COUNTROWS(FILTER(TimesheetData, HOUR(TimesheetData[StartTime]) = 0 && MINUTE(TimesheetData[StartTime]) = 0)),
    "Last StartTime", MAX(TimesheetData[StartTime]),
    "TimesheetData rows", COUNTROWS(TimesheetData),
    "Duplicate TimesheetData rows", COUNTROWS(TimesheetData) - COUNTROWS(DISTINCT(TimesheetData)),
    "AssessOnDecisionDay rows", COUNTROWS(AssessOnDecisionDay),
    "Distinct assess entries", COUNTROWS(SUMMARIZE(AssessOnDecisionDay,
        AssessOnDecisionDay[GroupedRows.AppUID], AssessOnDecisionDay[GroupedRows.Person], AssessOnDecisionDay[GroupedRows.StartTime]))
)
```

How to read the results:
- **PQ-04:** if "AssessOnDecisionDay rows" is clearly larger than "Distinct assess entries", the rows are multiplying.
- **DX-04:** any duplicate timesheet rows are counted N×N times by `TotalCost`.
- **SM-01, SM-02, SM-03:** each of those queries should return no rows.
