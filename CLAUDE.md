Project context for this session:



\- This is IMP9177 — 3H Pharmaceuticals LLC QMS Portal

\- SPFx web part project, main file:

&#x20; src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts

\- All changes go to that file via targeted string replacements

\- Build commands (run in order after any change):

&#x20;   .\\node\_modules\\.bin\\heft clean

&#x20;   .\\node\_modules\\.bin\\heft build --production

&#x20;   .\\node\_modules\\.bin\\heft package-solution --production

\- Deploy command:

&#x20;   Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive

&#x20;   Add-PnPApp -Path ".\\sharepoint\\solution\\imp-9177-spfx.sppkg" -Scope Site -Overwrite -Publish

\- Git push after every successful deploy:

&#x20;   $env:PATH += ";C:\\Program Files\\Git\\bin;C:\\Program Files\\Git\\cmd"

&#x20;   git add .

&#x20;   git commit -m "describe what changed"

&#x20;   git push origin main

\- SP site: https://adbccro.sharepoint.com/sites/IMP9177

\- Never use patch scripts in this session — edit the file directly

\- Read the relevant section of the file before making any change

