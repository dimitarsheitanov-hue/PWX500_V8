CAM Agent V5.1.38 - BUILD FIX3 / 2026-10-22

Cause: Stage 2 checked whether TimeImportProcess was non-null, not whether it had exited.
Fix: Only actually running workers block Stage 2; the time-import running flags still protect unfinished DB writes.
Scope: Only batch worker guard in src/CAM_Agent_BatchEngine_V5_1_38.ps1 and its build regression test.
No changes to PWX500_Agent_V3.ps1, HOP/PGMX/CSV/Search workers, Maestro/PowerSet automation or times rules.
This is a complete GitHub source/build package, not an already compiled EXE.
Compile using Actions -> Build CAM Agent V5.1.38 EXE -> Run workflow.
After build, download and extract the COMPLETE compiled ZIP with all worker files.
Windows UI automation must still be verified on the target machine.
