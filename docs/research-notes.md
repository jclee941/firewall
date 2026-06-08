# Research Notes

## Ready-to-open `.xlsm` generation

Excel macro-enabled files need an embedded `xl/vbaProject.bin`. That file is a binary OLE VBA project, not plain VBA text.

Confirmed approaches:

| Approach | Feasibility | Notes |
|---|---|---|
| Windows Excel COM automation | Works | Requires Microsoft Excel on Windows. Can import `.bas` and save as `.xlsm`. |
| XlsxWriter `add_vba_project()` | Works only with existing `vbaProject.bin` | Official XlsxWriter docs say it embeds a binary VBA project extracted from an existing `.xlsm`; it does not compile `.bas` into `vbaProject.bin`. |
| `vba_extract.py` | Extract only | Extracts `vbaProject.bin` from an existing `.xlsm`; cannot create one from scratch. |
| LibreOffice UNO on this host | Not sufficient | Can create/open workbooks, but did not produce a usable Excel VBA project binary from `.bas` in this environment. |
| Hand-writing `vbaProject.bin` | Not practical | Requires implementing the Office VBA binary/OLE project format. |

Useful references:

- XlsxWriter, “Working with VBA Macros”: explains that `.xlsm` is `.xlsx` plus `vbaProject.bin`, and that `vbaProject.bin` must be extracted from an existing macro workbook.
- XlsxWriter `examples/vba_extract.py`: extracts `xl/vbaProject.bin` from an existing `.xlsm`.

Practical path for this repo:

1. Keep `firewall-policy-automation.xlsx` as the clone-ready configured workbook.
2. Keep `vba/FirewallPolicyAutomation.bas` as the source macro module.
3. If a real starter `.xlsm` is required, create it once on Windows Excel, then commit that `.xlsm` or extract its `vbaProject.bin` and use XlsxWriter to reproduce it.

## Target firewall calculation

Related GitHub examples found:

| Repository | Useful pattern |
|---|---|
| `andreafortuna/VBAIPFunctions` | VBA functions such as `IpSubnetMatch`, `IpSubnetInSubnetMatch`, and `IpFindOverlappingSubnets` distinguish IP-in-subnet from subnet-in-subnet/overlap style matching. |
| `lsambolino/VBASubnetFinderChecker` | Excel/VBA style workflow for checking which subnet an IP belongs to. |
| `foryujian/ipintervalmerge` | Converts IP ranges/CIDR into numeric intervals before merging/comparing. |

Resulting implementation decision:

- The macro now converts each request IP/CIDR and firewall CIDR into numeric IPv4 ranges.
- It matches when the two ranges overlap: `leftStart <= rightEnd And rightStart <= leftEnd`.
- This handles both single IPs and CIDR-to-CIDR overlap.
- Multiple request values can be separated by semicolon, comma, Korean comma, or line breaks.
