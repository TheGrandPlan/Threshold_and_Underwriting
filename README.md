# Underwriting Analysis Apps Script

Production deployment copies live in `apps-script/production/` and are the ONLY files you copy into the Google Apps Script editor.

## Folders
- `apps-script/production/` - Minified/streamlined deployment `.gs` files.
- `originals/` - Archived verbose or legacy sources (excluded from Git; contains no secrets committed here). Do not deploy.

## Deployment (Manual Copy/Paste Workflow)
1. Open Google Sheet > Extensions > Apps Script of the target project.
2. Optionally delete all existing code files there (make a backup first if uncertain).
3. For each file in `apps-script/production/`:
   - Create (or open) a file in Apps Script with the same name (without directories).
   - Paste contents.
4. Ensure Script Properties (File > Project Properties > Script Properties) contain:
   - PARENT_FOLDER_ID
   - TEMPLATE_FILE_ID
   - N8N_WEBHOOK_URL (optional)
   - CENTRAL_URL (optional)
   - API_KEY (for geocoding)
   - SERVICE_ACCOUNT_EMAIL (optional)
   - OWNER_EMAIL (optional)
5. Save. Run a small test: execute `checkThresholdAndProcess` on a filtered sheet row to confirm folder/file creation.
6. Set any triggers (e.g., time-driven) for periodic execution (`geocodeAllPendingAddresses`, sorter, threshold checker as needed).

## Production Scripts
- `TACS.gs` - Threshold scan & asset creation.
- `GeoCodeAllPendingAddresses.gs` - Resumable batch geocoder.
- `AutoSorter.gs` - Multi-column sort (AB asc, X desc, R desc).
- `DeleteDuplicates.gs` - De-duplicate by address.

## Version Control Guidance
- Edit logic in root working copies (or future `/src`) if you re-expand; regenerate/refresh production files after changes.
- Commit only production copies plus supporting docs/config (avoid secrets).
- `.gitignore` excludes `originals/` and common secret patterns.

## Future Enhancements (Optional)
- Add TypeScript & build step to auto-produce production `.gs`.
- Integrate clasp for push/pull to Apps Script.
- Add simple test harness using clasp + local mocks.

## Safety Notes
- Never commit actual API keys or IDs you consider sensitive. Script Properties hold them safely.
- Review log output for any ERROR: markers after first run.

---
Last updated: INITIAL.
