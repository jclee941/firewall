firewall-policy-automation distribution

Primary artifact:
- firewall-policy-automation.xlsm : ready-to-use macro-enabled workbook
  (built on Linux by scripts/build_xlsm.py using pyOpenVBA + openpyxl;
   contains modules FirewallPolicyAutomation and FirewallRouteAnalysis,
   and the sheets requests, network_definitions, firewalls, routing_paths,
   settings, processing_log, sample-request-format, usage).

How to (re)build the .xlsm (no Windows / Excel / PowerShell needed):
1. From the repo root, create the venv once:
     python3 -m venv .venv && ./.venv/bin/pip install pyOpenVBA openpyxl pytest
2. Build:
     ./.venv/bin/python scripts/build_xlsm.py
   Output: dist/firewall-policy-automation.xlsm

How to use the workbook:
1. Open dist/firewall-policy-automation.xlsm in Microsoft Excel (enable macros).
2. Register IP ranges -> zones in the network_definitions sheet.
3. Register firewalls in the firewalls sheet.
4. Register zone-to-zone routing in the routing_paths sheet.
5. Put request workbooks in a folder, set settings!B2 (request_folder),
   then run macro MergeFirewallRequestFolder (Alt+F8).
6. To re-run only the path analysis, run macro AnalyzeRequestRoutes (Alt+F8).
7. Review requests columns: target_firewalls, firewall_path, zone_path,
   source_zone, destination_zone, validation_status, validation_message,
   match_details; and the processing_log sheet.

Macros are invoked via the Alt+F8 macro dialog. The workbook does not ship
with on-sheet buttons; you may add your own and assign the macros if desired.
Build is automated: scripts/build_xlsm.py (Linux, no Excel/PowerShell). CI and
tag-driven release workflows build/test/publish this .xlsm automatically.
