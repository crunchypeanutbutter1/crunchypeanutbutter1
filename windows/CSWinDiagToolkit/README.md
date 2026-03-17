# CSWinDiag Portable Toolkit Runner

This folder is designed to be copied onto a user workstation and run locally.

## What this does
- Verifies/elevates to Administrator (UAC prompt).
- Temporarily stages this toolkit under `%ProgramFiles%\AdminTools` (required by CSWinDiag).
- Runs `cswindiag.exe` automatically.
- Creates an easy-to-spot Desktop folder named:
  - `CSWinDiag-ReadyToSend-<COMPUTERNAME>-<TIMESTAMP>`
- Copies the newest `CSWinDiag*.zip` collection into that Desktop folder.
- Adds a `README-Transfer.txt` for handoff instructions.

## How to use
1. Place `cswindiag.exe` in this same folder.
2. Copy this whole folder to the target Windows machine.
3. Run `Run-CSWinDiagToolkit.cmd` (or `Invoke-CSWinDiagToolkit.ps1`).
4. Approve UAC/admin prompts if requested.
5. Wait for CSWinDiag to complete (typically 3–4 minutes).
6. Copy the Desktop output folder to your flash drive.

## Notes
- CSWinDiag must run from a folder under `%ProgramFiles%`; this tool handles that automatically.
- By default, the staged `%ProgramFiles%\AdminTools` copy is deleted when done.
- To keep staged files for troubleshooting, run PowerShell manually with `-KeepProgramFilesCopy`.
