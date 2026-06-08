firewall-policy-automation starter files

Files:
- firewall-policy-automation-starter.xlsx: preconfigured sheets (requests, firewalls, settings, sample-request-format)
- FirewallPolicyAutomation.bas: VBA macro module to import into Excel

To make the macro-enabled workbook in Excel:
1. Open firewall-policy-automation-starter.xlsx
2. Save As: firewall-policy-automation.xlsm
3. Press Alt+F11
4. File > Import File... > FirewallPolicyAutomation.bas
5. Run SetupFirewallAutomationWorkbook once
6. Set settings!B2 request_folder or run SelectRequestFolder
7. Run MergeFirewallRequestFolder

Note: creating a real .xlsm with embedded VBA requires Excel or an existing vbaProject.bin. This Linux environment generated the configured workbook and macro source only.
