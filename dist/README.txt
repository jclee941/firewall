firewall-policy-automation distribution

Primary artifact:
- firewall-policy-automation.xlsm : ready-to-use macro-enabled workbook
  (built on Linux by scripts/build_xlsm.py using pyOpenVBA + openpyxl;
   contains modules FirewallPolicyAutomation and FirewallRouteAnalysis,
   and the sheets requests, firewalls, firewall_ranges, settings,
   header_aliases, processing_log, route_results, secui_cli,
   vendor_cli_templates, service_catalog, sample-request-format, usage).

How to (re)build the .xlsm (no Windows / Excel / PowerShell needed):
1. From the repo root, create the venv once:
     python3 -m venv .venv && ./.venv/bin/pip install pyOpenVBA openpyxl pytest
2. Build:
     ./.venv/bin/python scripts/build_xlsm.py
   Output: dist/firewall-policy-automation.xlsm

How to use the workbook:
1. Open dist/firewall-policy-automation.xlsm in Microsoft Excel (enable macros).
2. Register SECUI firewall devices in the firewalls sheet.
3. Define source/destination ranges in firewall_ranges. 대상방화벽 is computed
   from these rows and written to requests / route_results.
4. Put request workbooks in a folder, set settings!B2 (request_folder),
   then reopen the workbook. Workbook_Open merges requests, analyzes routes,
   and refreshes secui_cli automatically.
5. To run manually, press F9 for the RunFirewallAutomationOutputs full workflow, or run
   MergeFirewallRequestFolder, AnalyzeRequestRoutes, and ConvertRequestsToSecuiCli
   in that order from the Alt+F8 macro dialog.
6. route_results shows route validation status, target firewall, matched
   ranges, path details, and original-file tracking for review.
7. ConvertRequestsToSecuiCli groups rows when the merge does not broaden
   access: same firewall + destination + service merges sources, and same
   firewall + source + destination merges services. Different source/service
   combinations stay split to avoid over-permitting.
8. ANY / ALL / * / 0.0.0.0/0 are emitted as policy ANY values without creating
   group objects.

SECUI service notes:
- vendor_cli_templates stores the SECUI CLI command template.
- service_catalog is a reference-only sheet for SECUI service notation
  examples such as tcp/443 and udp/53.
- route_results is the review surface for validation/routing diagnostics.
- requests stays as a clean request list; source tracking is kept in the hidden
  _request_tracking sheet and surfaced through route_results.

F9 runs the full output workflow. Individual macros are also available through
the Alt+F8 macro dialog.
Build is automated: scripts/build_xlsm.py (Linux, no Excel/PowerShell). CI and
tag-driven release workflows build/test/publish this .xlsm automatically.
