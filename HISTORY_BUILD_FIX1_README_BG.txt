CAM Agent V5.1.38 - BUILD FIX1 (GitHub compilation only)
1. Same V5.1.38 agent and the same existing runtime workers as the original package.
2. Fixed build/Test-CoreIntegrity.ps1: LF/CRLF-only changes no longer fail validation.
3. Any other source change still fails validation. Binary files remain exact SHA256 checked.
4. Added .gitattributes to prevent Windows Git from changing original runtime file bytes.
5. Upload all files, including hidden .github and .gitattributes, to a clean GitHub repository.
6. Actions -> Build CAM Agent V5.1.38 EXE -> Run workflow.
7. Extract the complete compiled ZIP, not the lone EXE.
This fix does not change the production agent logic or tested worker files.
