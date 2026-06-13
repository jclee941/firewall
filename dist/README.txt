firewall-policy-automation distribution

Primary artifact:
- firewall-policy-automation.xlsm : ready-to-use macro-enabled workbook
  (built on Linux by scripts/build_xlsm.py using pyOpenVBA + openpyxl;
   contains modules FirewallPolicyAutomation and FirewallRouteAnalysis,
   and the sheets requests, firewalls, firewall_ranges, settings,
   header_aliases, processing_log, secui_batch, secui_cli,
   vendor_cli_templates, service_catalog, sample-request-format, usage).

How to (re)build the .xlsm (no Windows / Excel / PowerShell needed):
1. From the repo root, create the venv once:
     python3 -m venv .venv && ./.venv/bin/pip install pyOpenVBA openpyxl pytest
2. Build:
     ./.venv/bin/python scripts/build_xlsm.py
   Output: dist/firewall-policy-automation.xlsm

How to use the workbook:
1. Open dist/firewall-policy-automation.xlsm in Microsoft Excel (enable macros).
2. Register firewall devices in the firewalls sheet.
3. Register source/destination firewall ranges in the firewall_ranges sheet.
4. Put request workbooks in a folder, set settings!B2 (request_folder),
   then run macro MergeFirewallRequestFolder (Alt+F8).
5. To re-run only the firewall range analysis, run macro AnalyzeRequestRoutes
   (Alt+F8).
6. Review requests columns including 검증상태, 대상방화벽, 방화벽경로,
   출발매칭대역, 목적매칭대역, 대역경로, 매칭근거; and the processing_log
   sheet.
7. For SECUI output, run ConvertRequestsToSecuiBatch or
   ConvertRequestsToSecuiCli and review secui_batch / secui_cli.

SECUI service notes:
- vendor_cli_templates stores the SECUI CLI command template.
- service_catalog is a reference-only sheet for SECUI service notation
  examples such as tcp/443 and udp/53. It is not route input; route
  calculation uses firewalls and firewall_ranges.

Macros are invoked via the Alt+F8 macro dialog. The workbook does not ship
with on-sheet buttons; you may add your own and assign the macros if desired.
Build is automated: scripts/build_xlsm.py (Linux, no Excel/PowerShell). CI and
tag-driven release workflows build/test/publish this .xlsm automatically.
