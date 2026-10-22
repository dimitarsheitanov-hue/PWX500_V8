V5.1.38 BUILD FIX6 -- TARGETED NAVIGATION AND SPEED FIX

Source: BUILD FIX5. No changes to: CSV/HOP/PGMX/History/Search workers,
TimesImportWorker, times DB, converter, schedule logic, whitelist aliases,
PowerSet optimization settings, MIXX Save-as confirmation safety.

FIX6 changes (main PWX500_Agent_V3.ps1 only):
1. Add programs: recognizes DISABLED Confirm as a geometry anchor without
   requiring UIA FromPoint to hit that disabled button. Never clicks Confirm
   until PGMX selection.
2. Cumulative batch: searches exact folder in center list with UIA
   ItemContainerPattern and falls back to bounded PageDown for WPF Lists that
   do not provide ScrollPattern. Requires verified OPEN breadcrumbs.
3. PowerSet finished? One process-wide UIA enumeration per poll rather than
   six, preserving the two-independent-result-label decision and 60m timeout.
   Does NOT change any optimizer settings or automate Maestro by guessing.
4. BatchEngine has start/end phase timing markers in the existing logs.
5. Extra GitHub static/runtime regression tests for disabled button,
   virtualized center list, and two-signal PowerSet completion.

IMPORTANT: The user-observed 14 minutes have NOT been profiled yet on the
actual Windows workstation. The result of FIX6 must be tested with real
Maestro/PowerSet. We do NOT promise that 30-60 seconds are achievable.

TEST:
- Keep the last working version as backup. Use a separate test copy.
- Start Stage 2 only (do not repeat Stage 1) for an existing PREPARED batch.
- Observe Add programs: open FOUND -> Общи_пакети -> exact batch folder.
- Then check Save as, Choose file and Export XLS.
- Inspect PWX500_AutoMix.log (or operator Mix log) and
  POWERSET_AUTOMATION_LOG.txt for timestamps / PERF lines.
- Never delete a PREPARED batch manifest or real times database to 'reset'.
