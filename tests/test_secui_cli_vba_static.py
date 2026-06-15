import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VBA_POLICY = os.path.join(ROOT, "vba", "FirewallPolicyAutomation.bas")


def test_secui_cli_vba_assigns_dictionary_objects_with_set():
    src = open(VBA_POLICY, encoding="utf-8").read()
    assert "Set serviceFanoutIndex(sourceDestinationKey) = services" in src
    assert "Set cliGroups(groupKey) = group" in src
    assert not re.search(r"^\s*serviceFanoutIndex\(sourceDestinationKey\)\s*=\s*services\s*$", src, re.MULTILINE)
    assert not re.search(r"^\s*cliGroups\(groupKey\)\s*=\s*group\s*$", src, re.MULTILINE)
