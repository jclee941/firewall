firewall-policy-automation distribution

Primary artifact:
- firewall-policy-automation.xlsm : ready-to-use macro-enabled workbook
  (built on Linux by scripts/build_xlsm.py using pyOpenVBA + openpyxl;
   contains modules FirewallPolicyAutomation and FirewallRouteAnalysis,
   and the sheets requests, firewalls, firewall_ranges, settings,
   header_aliases, processing_log, secui_batch, secui_cli,
   secui_policy_export, policy_analysis, policy_summary,
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
3. Put SECUI device names in requests.대상방화벽 or in the source request files
   (use ; for multiple devices).
4. Put request workbooks in a folder, set settings!B2 (request_folder),
   then reopen the workbook. Workbook_Open merges requests and refreshes
   secui_batch / secui_cli / policy_analysis automatically.
5. To run manually, use MergeFirewallRequestFolder,
   ConvertRequestsToSecuiBatch, and ConvertRequestsToSecuiCli (Alt+F8).
6. ConvertRequestsToSecuiCli groups rows when the merge does not broaden
   access: same firewall + destination + service merges sources, and same
   firewall + source + destination merges services. Different source/service
   combinations stay split to avoid over-permitting.
7. ANY / ALL / * / 0.0.0.0/0 are emitted as policy ANY values without creating
   group objects. Interface tokens come from hidden firewall_ranges
   source_interface / destination_interface helper columns; unmatched ranges
   are emitted as ANY interfaces for review.

SECUI service notes:
- vendor_cli_templates stores the SECUI CLI command template.
- service_catalog is a reference-only sheet for SECUI service notation
  examples such as tcp/443 and udp/53.
- secui_policy_export / policy_analysis / policy_summary are optional existing
  policy review surfaces. CLI generation does not require a policy export.
- requests validation/routing columns are hidden by default. They are for
  manual route review later, not for the default SECUI CLI workflow.

Macros are invoked via the Alt+F8 macro dialog. The workbook does not ship
with on-sheet buttons; you may add your own and assign the macros if desired.
Build is automated: scripts/build_xlsm.py (Linux, no Excel/PowerShell). CI and
tag-driven release workflows build/test/publish this .xlsm automatically.
