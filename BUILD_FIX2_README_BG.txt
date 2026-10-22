CAM Agent V5.1.38 - BUILD FIX2 - EXE batch engine scope repair

Problem observed in a compiled V5.1.38 EXE:
  Sync-PgmxAutoExclusionsFromDailyFolders is not recognized.

Cause addressed: the new CAM_Agent_BatchEngine_V5_1_38.ps1 was loaded by
runtime dot-source. In compiled PS2EXE invocation, its functions were not
available where the GUI startup called them.

FIX2 changes ONLY build/Build-CAM-Agent.ps1:
- Assembles the compiled EXE script with all 21 batch functions embedded.
- Keeps the full original V5.1.38 source / standalone batch module unchanged.
- Checks that every batch function is present exactly once before compilation.
- Keeps the external module file inside the compiled ZIP for compatibility
  with the original PowerShell source and troubleshooting.
- Does not change CSV, HOP, PGMX, MIXX, PowerSet, time or schedule logic.

GITHUB:
1. Use an empty GitHub repository to avoid mixing old builds.
2. Upload ALL contents, including .github and .gitattributes.
3. Actions -> Build CAM Agent V5.1.38 EXE -> Run workflow.
4. Download the full CAM_Agent_V5_1_38_COMPILED_PACKAGE.zip artifact.
5. Keep all files beside CAM_Agent_V5_1_38.exe.

Build checks can validate script structure, but the Maestro/PowerSet GUI
must still be tested on the target Windows machine.
