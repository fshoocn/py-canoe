"""
Unit tests for test module and test case features:
- TestCase class (properties, verdict_name, to_dict)
- TestModule test case methods (collection, verdict, enable/disable)
- Configuration._match_test_case_name (wildcard and regex)
- Configuration._apply_test_case_selection (enable/disable logic)
- Configuration._find_test_module (lookup helper)
- Configuration.execute_test_module with enable/disable patterns
- Configuration.get_test_module_result (report + test case verdicts)
"""

from unittest.mock import MagicMock, Mock, patch

from py_canoe.core.child_elements.test_case import TestCase
from py_canoe.core.child_elements.test_module import TestModule
from py_canoe.core.configuration import Configuration


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_test_case(name="TC_001", enabled=True, verdict=0, title="Test Case 001"):
    """Create a TestCase instance with mocked COM (bypasses Dispatch)."""
    com = MagicMock()
    com.Name = name
    com.Enabled = enabled
    com.Verdict = verdict
    com.Title = title
    with patch("win32com.client.Dispatch", side_effect=lambda x: x):
        return TestCase(com)


def _make_test_module_wrapper(test_cases=None):
    """Create a TestModule wrapper with mocked COM and test cases.

    Args:
        test_cases: list of dicts with keys 'name', 'enabled', 'verdict', 'title'
    """
    if test_cases is None:
        test_cases = []

    com = MagicMock()
    com.Name = "TestMod1"
    com.FullName = "Env1::TestMod1"
    com.Verdict = 1

    # Build test sequence tree
    items = []
    for tc in test_cases:
        item = MagicMock()
        item.Type = 5  # TestCase
        item.Name = tc["name"]
        item.Enabled = tc.get("enabled", True)
        item.Verdict = tc.get("verdict", 0)
        item.Title = tc.get("title", tc["name"])
        items.append(item)

    sequence = MagicMock()
    sequence.Count = len(items)
    sequence.Item = lambda idx: items[idx - 1]  # 1-based index
    com.TestSequence = sequence

    # Mock test_module_events
    events = MagicMock()
    events.TEST_REPORT_INFORMATION = {
        "success": True,
        "source_full_name": "C:/report.xml",
        "generated_full_name": "C:/report.html"
    }

    with patch("win32com.client.Dispatch", side_effect=lambda x: x), \
         patch("win32com.client.WithEvents", return_value=events):
        tm = TestModule(com)

    return tm


def _make_test_module_with_groups():
    """Create a TestModule with nested TestGroup -> TestSequence -> TestCase."""
    com = MagicMock()
    com.Name = "TestModGroup"
    com.FullName = "Env1::TestModGroup"
    com.Verdict = 2

    # TestGroup item
    group = MagicMock()
    group.Type = 3  # TestGroup
    group.Name = "Group1"

    # TestCase inside group
    tc_in_group = MagicMock()
    tc_in_group.Type = 5  # TestCase
    tc_in_group.Name = "TC_InGroup"
    tc_in_group.Enabled = True
    tc_in_group.Verdict = 1
    tc_in_group.Title = "TC In Group"

    group_sequence = MagicMock()
    group_sequence.Count = 1
    group_sequence.Item = lambda idx: tc_in_group
    group.Sequence = group_sequence

    # Top-level TestCase
    tc_top = MagicMock()
    tc_top.Type = 5
    tc_top.Name = "TC_Top"
    tc_top.Enabled = False
    tc_top.Verdict = 2
    tc_top.Title = "TC Top Level"

    sequence = MagicMock()
    sequence.Count = 2
    items = [tc_top, group]
    sequence.Item = lambda idx: items[idx - 1]
    com.TestSequence = sequence

    events = MagicMock()
    events.TEST_REPORT_INFORMATION = {}

    with patch("win32com.client.Dispatch", side_effect=lambda x: x), \
         patch("win32com.client.WithEvents", return_value=events):
        tm = TestModule(com)

    return tm


# ===========================================================================
# TestCase class tests
# ===========================================================================

class TestTestCaseClass:
    """Test TestCase COM wrapper class."""

    def test_name_property(self):
        tc = _make_test_case(name="MyTestCase")
        assert tc.name == "MyTestCase"

    def test_title_property(self):
        tc = _make_test_case(title="My Title")
        assert tc.title == "My Title"

    def test_enabled_getter(self):
        tc = _make_test_case(enabled=False)
        assert tc.enabled is False

    def test_enabled_setter(self):
        tc = _make_test_case(enabled=False)
        tc.enabled = True
        assert tc.com_object.Enabled is True

    def test_verdict_property(self):
        tc = _make_test_case(verdict=2)
        assert tc.verdict == 2

    def test_verdict_name_passed(self):
        tc = _make_test_case(verdict=1)
        assert tc.verdict_name == "Passed"

    def test_verdict_name_failed(self):
        tc = _make_test_case(verdict=2)
        assert tc.verdict_name == "Failed"

    def test_verdict_name_not_available(self):
        tc = _make_test_case(verdict=0)
        assert tc.verdict_name == "NotAvailable"

    def test_verdict_name_inconclusive(self):
        tc = _make_test_case(verdict=4)
        assert tc.verdict_name == "Inconclusive"

    def test_verdict_name_error_in_test_system(self):
        tc = _make_test_case(verdict=5)
        assert tc.verdict_name == "ErrorInTestSystem"

    def test_to_dict(self):
        tc = _make_test_case(name="TC1", enabled=True, verdict=1, title="Title1")
        d = tc.to_dict()
        assert d == {
            "name": "TC1",
            "title": "Title1",
            "enabled": True,
            "verdict": 1,
            "verdict_name": "Passed"
        }

    def test_repr(self):
        tc = _make_test_case(name="TC1", enabled=True, verdict=1)
        assert "TC1" in repr(tc)
        assert "Passed" in repr(tc)

    def test_value_table_verdict_class_attribute(self):
        assert TestCase.VALUE_TABLE_VERDICT[0] == "NotAvailable"
        assert TestCase.VALUE_TABLE_VERDICT[1] == "Passed"
        assert TestCase.VALUE_TABLE_VERDICT[2] == "Failed"
        assert TestCase.VALUE_TABLE_VERDICT[3] == "None"
        assert TestCase.VALUE_TABLE_VERDICT[4] == "Inconclusive"
        assert TestCase.VALUE_TABLE_VERDICT[5] == "ErrorInTestSystem"


# ===========================================================================
# TestModule test case methods
# ===========================================================================

class TestModuleTestCases:
    """Test TestModule test case collection and access methods."""

    # _collect_test_cases calls TestCase(item) which calls Dispatch internally,
    # so we must keep Dispatch patched during the get_all_test_calls etc. calls.

    def test_get_all_test_cases_empty(self):
        tm = _make_test_module_wrapper(test_cases=[])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x):
            result = tm.get_all_test_cases()
        assert result == {}

    def test_get_all_test_cases_returns_dict(self):
        tm = _make_test_module_wrapper([
            {"name": "TC_001", "enabled": True, "verdict": 1},
            {"name": "TC_002", "enabled": False, "verdict": 2},
        ])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x):
            result = tm.get_all_test_cases()
        assert len(result) == 2
        assert "TC_001" in result
        assert "TC_002" in result
        assert isinstance(result["TC_001"], TestCase)

    def test_get_all_test_cases_preserves_enabled_state(self):
        tm = _make_test_module_wrapper([
            {"name": "TC_001", "enabled": True, "verdict": 0},
            {"name": "TC_002", "enabled": False, "verdict": 0},
        ])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x):
            result = tm.get_all_test_cases()
        assert result["TC_001"].enabled is True
        assert result["TC_002"].enabled is False

    def test_get_all_test_cases_with_nested_groups(self):
        tm = _make_test_module_with_groups()
        with patch("win32com.client.Dispatch", side_effect=lambda x: x):
            result = tm.get_all_test_cases()
        assert len(result) == 2
        assert "TC_Top" in result
        assert "TC_InGroup" in result

    def test_get_test_case_found(self):
        tm = _make_test_module_wrapper([
            {"name": "TC_001", "enabled": True, "verdict": 1},
        ])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x):
            tc = tm.get_test_case("TC_001")
        assert tc is not None
        assert tc.name == "TC_001"

    def test_get_test_case_not_found(self):
        tm = _make_test_module_wrapper([
            {"name": "TC_001", "enabled": True, "verdict": 1},
        ])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x):
            tc = tm.get_test_case("NonExistent")
        assert tc is None

    def test_get_test_case_verdict_found(self):
        tm = _make_test_module_wrapper([
            {"name": "TC_001", "enabled": True, "verdict": 2},
        ])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x):
            assert tm.get_test_case_verdict("TC_001") == 2

    def test_get_test_case_verdict_not_found(self):
        tm = _make_test_module_wrapper([])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x):
            assert tm.get_test_case_verdict("NonExistent") == -1

    def test_get_test_case_enabled_found(self):
        tm = _make_test_module_wrapper([
            {"name": "TC_001", "enabled": False, "verdict": 0},
        ])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x):
            assert tm.get_test_case_enabled("TC_001") is False

    def test_get_test_case_enabled_not_found(self):
        tm = _make_test_module_wrapper([])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x):
            assert tm.get_test_case_enabled("NonExistent") is None

    def test_set_test_case_enabled_success(self):
        tm = _make_test_module_wrapper([
            {"name": "TC_001", "enabled": False, "verdict": 0},
        ])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x):
            result = tm.set_test_case_enabled("TC_001", True)
        assert result is True

    def test_set_test_case_enabled_not_found(self):
        tm = _make_test_module_wrapper([])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x):
            result = tm.set_test_case_enabled("NonExistent", True)
        assert result is False

    def test_get_all_test_case_verdicts(self):
        tm = _make_test_module_wrapper([
            {"name": "TC_001", "enabled": True, "verdict": 1},
            {"name": "TC_002", "enabled": False, "verdict": 2},
        ])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x):
            result = tm.get_all_test_case_verdicts()
        assert result["TC_001"]["verdict"] == 1
        assert result["TC_001"]["verdict_name"] == "Passed"
        assert result["TC_001"]["enabled"] is True
        assert result["TC_002"]["verdict"] == 2
        assert result["TC_002"]["verdict_name"] == "Failed"
        assert result["TC_002"]["enabled"] is False


# ===========================================================================
# Configuration._match_test_case_name
# ===========================================================================

class TestMatchTestCaseName:
    """Test pattern matching logic."""

    # --- Wildcard tests ---

    def test_wildcard_star_matches_all(self):
        assert Configuration._match_test_case_name("TC_001", "*") is True
        assert Configuration._match_test_case_name("Anything", "*") is True

    def test_wildcard_prefix(self):
        assert Configuration._match_test_case_name("SmokeTest_001", "SmokeTest_*") is True
        assert Configuration._match_test_case_name("Regression_001", "SmokeTest_*") is False

    def test_wildcard_suffix(self):
        assert Configuration._match_test_case_name("TC_001_pass", "*_pass") is True
        assert Configuration._match_test_case_name("TC_001_fail", "*_pass") is False

    def test_wildcard_question_mark(self):
        assert Configuration._match_test_case_name("TC_001", "TC_00?") is True
        assert Configuration._match_test_case_name("TC_001", "TC_0?1") is True
        assert Configuration._match_test_case_name("TC_0012", "TC_00?") is False

    def test_wildcard_bracket_seq(self):
        assert Configuration._match_test_case_name("TC_001", "TC_00[123]") is True
        assert Configuration._match_test_case_name("TC_004", "TC_00[123]") is False

    def test_wildcard_bracket_neg_seq(self):
        # fnmatchcase uses [!seq] for negation
        assert Configuration._match_test_case_name("TC_004", "TC_00[!123]") is True
        assert Configuration._match_test_case_name("TC_001", "TC_00[!123]") is False

    def test_wildcard_multiple_stars(self):
        assert Configuration._match_test_case_name("SmokeTest_Module1_Pass", "*Test_*_Pass") is True
        assert Configuration._match_test_case_name("SmokeTest_Module1_Fail", "*Test_*_Pass") is False

    # --- Regex tests ---

    def test_regex_prefix_group(self):
        assert Configuration._match_test_case_name("tc_001", "(?i)^tc_(001|002|003)$") is True
        assert Configuration._match_test_case_name("tc_004", "(?i)^tc_(001|002|003)$") is False
        assert Configuration._match_test_case_name("TC_001", "(?i)^tc_(001|002|003)$") is True

    def test_regex_caret_anchor(self):
        assert Configuration._match_test_case_name("SmokeTest_001", "^Smoke") is True
        assert Configuration._match_test_case_name("Regression_Smoke", "^Smoke") is False

    def test_regex_dollar_anchor(self):
        assert Configuration._match_test_case_name("TC_001_Pass", "Pass$") is True
        assert Configuration._match_test_case_name("TC_001_Fail", "Pass$") is False

    def test_regex_plus_quantifier(self):
        assert Configuration._match_test_case_name("TC_0001", r"TC_0+1") is True
        assert Configuration._match_test_case_name("TC_1", r"TC_0+1") is False

    def test_regex_braces_quantifier(self):
        assert Configuration._match_test_case_name("TC_001", r"TC_\d{3}") is True
        assert Configuration._match_test_case_name("TC_01", r"TC_\d{3}") is False

    def test_regex_square_brackets(self):
        assert Configuration._match_test_case_name("TC_A", r"TC_[A-Z]") is True
        assert Configuration._match_test_case_name("TC_1", r"TC_[A-Z]") is False

    def test_regex_backslash(self):
        assert Configuration._match_test_case_name("TC_001", r"TC_\d+") is True
        assert Configuration._match_test_case_name("TC_abc", r"TC_\d+") is False

    # --- Regex vs glob disambiguation ---

    def test_bracket_range_with_plus_is_regex(self):
        # TC_[0-9]+ → regex (contains \d-like range and +)
        assert Configuration._match_test_case_name("TC_123", r"TC_[0-9]+") is True
        assert Configuration._match_test_case_name("TC_abc", r"TC_[0-9]+") is False

    def test_bracket_range_without_plus_is_glob(self):
        # TC_[0-9] → glob (single char range, no +)
        assert Configuration._match_test_case_name("TC_1", "TC_[0-9]") is True
        assert Configuration._match_test_case_name("TC_a", "TC_[0-9]") is False

    def test_brace_quantifier_is_regex(self):
        # {3} → regex quantifier
        assert Configuration._match_test_case_name("TC_001", r"TC_\d{3}") is True
        assert Configuration._match_test_case_name("TC_01", r"TC_\d{3}") is False

    def test_brace_alternation_is_glob_no_expansion(self):
        # {foo,bar} → treated as fnmatch, but Python fnmatch doesn't support
        # brace expansion, so it becomes a literal match (no match)
        assert Configuration._match_test_case_name("TC_foo", "TC_{foo,bar}") is False
        assert Configuration._match_test_case_name("TC_{foo,bar}", "TC_{foo,bar}") is True

    def test_posix_class_in_bracket_detected_as_regex(self):
        # [[:alpha:]] → detected as regex (contains [:...:] inside brackets)
        # but Python re doesn't support POSIX classes, so use \w equivalent instead
        assert Configuration._match_test_case_name("TC_a", r"TC_[\w]") is True
        assert Configuration._match_test_case_name("TC_1", r"TC_[\w]") is True
        assert Configuration._match_test_case_name("TC_", r"TC_[\w]") is False

    # --- Edge cases ---

    def test_invalid_regex_falls_back_to_literal(self):
        assert Configuration._match_test_case_name("TC[", "TC[") is True
        assert Configuration._match_test_case_name("TC_001", "TC[") is False

    def test_exact_match(self):
        assert Configuration._match_test_case_name("TC_001", "TC_001") is True
        assert Configuration._match_test_case_name("TC_002", "TC_001") is False

    def test_case_sensitive_wildcard(self):
        assert Configuration._match_test_case_name("tc_001", "TC_*") is False
        assert Configuration._match_test_case_name("TC_001", "TC_*") is True

    def test_empty_pattern(self):
        assert Configuration._match_test_case_name("", "") is True
        assert Configuration._match_test_case_name("TC_001", "") is False


# ===========================================================================
# Configuration._apply_test_case_selection
# ===========================================================================

class TestApplyTestCaseSelection:
    """Test enable/disable test case selection logic."""

    def _make_cfg_with_test_cases(self, test_cases):
        """Create Configuration with test module COM mock returning given test cases."""
        cfg = Configuration.__new__(Configuration)

        # Build COM mock items
        items = []
        for tc in test_cases:
            item = MagicMock()
            item.Type = 5
            item.Name = tc["name"]
            item.Enabled = tc.get("enabled", True)
            item.Verdict = tc.get("verdict", 0)
            item.Title = tc.get("title", tc["name"])
            items.append(item)

        com = MagicMock()
        com.Name = "TestMod1"
        sequence = MagicMock()
        sequence.Count = len(items)
        sequence.Item = lambda idx: items[idx - 1]
        com.TestSequence = sequence

        return cfg, com, items

    def test_no_patterns_does_nothing(self):
        cfg, com, items = self._make_cfg_with_test_cases([
            {"name": "TC_001", "enabled": True},
        ])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x), \
             patch("win32com.client.WithEvents", return_value=MagicMock()):
            cfg._apply_test_case_selection(com, (), ())
        assert items[0].Enabled is True

    def test_enable_pattern_enables_matching(self):
        cfg, com, items = self._make_cfg_with_test_cases([
            {"name": "TC_001", "enabled": False},
            {"name": "TC_002", "enabled": False},
        ])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x), \
             patch("win32com.client.WithEvents", return_value=MagicMock()):
            cfg._apply_test_case_selection(com, ["TC_001"], ())
        assert items[0].Enabled is True   # TC_001 enabled
        assert items[1].Enabled is False  # TC_002 unchanged

    def test_disable_pattern_disables_matching(self):
        cfg, com, items = self._make_cfg_with_test_cases([
            {"name": "TC_001", "enabled": True},
            {"name": "TC_slow_001", "enabled": True},
        ])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x), \
             patch("win32com.client.WithEvents", return_value=MagicMock()):
            cfg._apply_test_case_selection(com, ["*"], ["*slow*"])
        assert items[0].Enabled is True   # TC_001 stays enabled
        assert items[1].Enabled is False  # TC_slow_001 disabled

    def test_disable_takes_precedence_over_enable(self):
        cfg, com, items = self._make_cfg_with_test_cases([
            {"name": "TC_slow_001", "enabled": True},
        ])
        with patch("win32com.client.Dispatch", side_effect=lambda x: x), \
             patch("win32com.client.WithEvents", return_value=MagicMock()):
            cfg._apply_test_case_selection(com, ["*"], ["*slow*"])
        assert items[0].Enabled is False


# ===========================================================================
# Configuration._find_test_module
# ===========================================================================

class TestFindTestModule:
    """Test _find_test_module helper."""

    def test_finds_existing_module(self):
        cfg = Configuration.__new__(Configuration)
        mock_tm = Mock()
        cfg._Configuration__test_modules = [
            {"name": "TestMod1", "object": mock_tm, "environment": "Env1"}
        ]
        result = cfg._find_test_module("TestMod1")
        assert result is mock_tm

    def test_returns_none_for_missing_module(self):
        cfg = Configuration.__new__(Configuration)
        cfg._Configuration__test_modules = [
            {"name": "TestMod1", "object": Mock(), "environment": "Env1"}
        ]
        result = cfg._find_test_module("NonExistent")
        assert result is None

    def test_finds_correct_module_among_multiple(self):
        cfg = Configuration.__new__(Configuration)
        mock_tm1 = Mock()
        mock_tm2 = Mock()
        cfg._Configuration__test_modules = [
            {"name": "TestMod1", "object": mock_tm1, "environment": "Env1"},
            {"name": "TestMod2", "object": mock_tm2, "environment": "Env2"},
        ]
        assert cfg._find_test_module("TestMod2") is mock_tm2

    def test_empty_list_returns_none(self):
        cfg = Configuration.__new__(Configuration)
        cfg._Configuration__test_modules = []
        result = cfg._find_test_module("AnyModule")
        assert result is None


# ===========================================================================
# Configuration.execute_test_module with patterns
# ===========================================================================

class TestExecuteTestModuleWithPatterns:
    """Test execute_test_module with enable/disable patterns."""

    def _make_cfg_with_mock_tm(self, verdict=1):
        """Create Configuration with a mock test module that returns the given verdict."""
        cfg = Configuration.__new__(Configuration)
        mock_tm = Mock()
        mock_tm.verdict = verdict
        cfg._Configuration__test_modules = [
            {"name": "TestMod1", "object": mock_tm, "environment": "Env1"}
        ]
        return cfg, mock_tm

    def test_execute_without_patterns_unchanged(self):
        cfg, mock_tm = self._make_cfg_with_mock_tm(verdict=1)

        with patch("py_canoe.core.configuration.Configuration._apply_test_case_selection") as mock_apply:
            result = cfg.execute_test_module("TestMod1")

        assert result == 1
        mock_apply.assert_called_once_with(mock_tm, (), ())

    def test_execute_with_enable_patterns(self):
        cfg, mock_tm = self._make_cfg_with_mock_tm(verdict=1)

        with patch("py_canoe.core.configuration.Configuration._apply_test_case_selection") as mock_apply:
            cfg.execute_test_module("TestMod1", enable_test_cases=["TC_*"])

        mock_apply.assert_called_once_with(mock_tm, ["TC_*"], ())

    def test_execute_with_disable_patterns(self):
        cfg, mock_tm = self._make_cfg_with_mock_tm(verdict=1)

        with patch("py_canoe.core.configuration.Configuration._apply_test_case_selection") as mock_apply:
            cfg.execute_test_module("TestMod1", disable_test_cases=["*slow*"])

        mock_apply.assert_called_once_with(mock_tm, (), ["*slow*"])

    def test_execute_with_both_patterns(self):
        cfg, mock_tm = self._make_cfg_with_mock_tm(verdict=1)

        with patch("py_canoe.core.configuration.Configuration._apply_test_case_selection") as mock_apply:
            cfg.execute_test_module("TestMod1", enable_test_cases=["*"], disable_test_cases=["*slow*"])

        mock_apply.assert_called_once_with(mock_tm, ["*"], ["*slow*"])

    def test_execute_not_found_returns_zero(self):
        cfg = Configuration.__new__(Configuration)
        cfg._Configuration__test_modules = []

        result = cfg.execute_test_module("NonExistent")
        assert result == 0

    def test_execute_exception_returns_zero(self):
        cfg = Configuration.__new__(Configuration)
        mock_tm = MagicMock()
        mock_tm.start.side_effect = Exception("COM error")
        cfg._Configuration__test_modules = [
            {"name": "TestMod1", "object": mock_tm, "environment": "Env1"}
        ]

        with patch("py_canoe.core.configuration.Configuration._apply_test_case_selection"):
            result = cfg.execute_test_module("TestMod1")

        assert result == 0


# ===========================================================================
# Configuration.get_test_module_result
# ===========================================================================

class TestGetTestModuleResult:
    """Test get_test_module_result method."""

    def _make_cfg_for_result(self, test_cases, verdict=1, report_info=None):
        cfg = Configuration.__new__(Configuration)
        tm_wrapper = _make_test_module_wrapper(test_cases)
        tm_wrapper.com_object.Verdict = verdict
        if report_info is not None:
            tm_wrapper.test_module_events.TEST_REPORT_INFORMATION = report_info
        cfg._Configuration__test_modules = [
            {"name": "TestMod1", "object": tm_wrapper.com_object, "environment": "Env1"}
        ]
        return cfg, tm_wrapper

    def test_result_structure(self):
        cfg, tm = self._make_cfg_for_result([
            {"name": "TC_001", "enabled": True, "verdict": 1, "title": "TC 1"},
        ])

        with patch("py_canoe.core.configuration.Configuration._find_test_module", return_value=tm.com_object), \
             patch("py_canoe.core.child_elements.test_module.TestModule", return_value=tm), \
             patch("win32com.client.Dispatch", side_effect=lambda x: x):
            result = cfg.get_test_module_result("TestMod1")

        assert "verdict" in result
        assert "verdict_name" in result
        assert "report" in result
        assert "test_cases" in result

    def test_result_verdict(self):
        cfg, tm = self._make_cfg_for_result([], verdict=2)

        with patch("py_canoe.core.configuration.Configuration._find_test_module", return_value=tm.com_object), \
             patch("py_canoe.core.child_elements.test_module.TestModule", return_value=tm), \
             patch("win32com.client.Dispatch", side_effect=lambda x: x):
            result = cfg.get_test_module_result("TestMod1")

        assert result["verdict"] == 2
        assert result["verdict_name"] == "Failed"

    def test_result_report_info(self):
        cfg, tm = self._make_cfg_for_result([], report_info={
            "success": True,
            "source_full_name": "C:/test.xml",
            "generated_full_name": "C:/test.html"
        })

        with patch("py_canoe.core.configuration.Configuration._find_test_module", return_value=tm.com_object), \
             patch("py_canoe.core.child_elements.test_module.TestModule", return_value=tm), \
             patch("win32com.client.Dispatch", side_effect=lambda x: x):
            result = cfg.get_test_module_result("TestMod1")

        assert result["report"]["success"] is True
        assert result["report"]["source_full_name"] == "C:/test.xml"
        assert result["report"]["generated_full_name"] == "C:/test.html"

    def test_result_test_cases(self):
        cfg, tm = self._make_cfg_for_result([
            {"name": "TC_001", "enabled": True, "verdict": 1, "title": "Test 1"},
        ])

        with patch("py_canoe.core.configuration.Configuration._find_test_module", return_value=tm.com_object), \
             patch("py_canoe.core.child_elements.test_module.TestModule", return_value=tm), \
             patch("win32com.client.Dispatch", side_effect=lambda x: x):
            result = cfg.get_test_module_result("TestMod1")

        assert "TC_001" in result["test_cases"]
        tc = result["test_cases"]["TC_001"]
        assert tc["name"] == "TC_001"
        assert tc["enabled"] is True
        assert tc["verdict"] == 1
        assert tc["verdict_name"] == "Passed"
        assert tc["title"] == "Test 1"

    def test_result_not_found_returns_empty(self):
        cfg = Configuration.__new__(Configuration)
        cfg._Configuration__test_modules = []
        assert cfg.get_test_module_result("NonExistent") == {}

    def test_result_exception_returns_empty(self):
        cfg = Configuration.__new__(Configuration)
        cfg._Configuration__test_modules = []
        with patch("py_canoe.core.configuration.Configuration._find_test_module", side_effect=Exception("error")):
            assert cfg.get_test_module_result("TestMod1") == {}
