# VBA Project Versioning Rule
Whenever you make a modification to any VBA code (`.bas` or `.cls` files) in this project, you MUST:
1. Increment the version number in the file name (e.g., from `modDemoReports_V9.35.bas` to `modDemoReports_V9.36.bas`).
2. Update the internal version constant (e.g., `SYSTEM_VERSION = "9.36"`).
3. Add a comment header at the top of the file (or near the version constant) indicating the new version, the current date and time, and a brief summary of the changes made.
This rule ensures basic organized work practices for version control and auditing.
