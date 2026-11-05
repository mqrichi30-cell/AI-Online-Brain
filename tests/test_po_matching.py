import pytest

from nova_sap.po_matching import build_detail_rows, evaluate_trigger, normalize_po


@pytest.mark.parametrize("raw,expected", [
    ("1234567890", "1234567890"),
    ("  1234567890 ", "1234567890"),
    (1234567890.0, "1234567890"),
    ("1234567890.0", "1234567890"),
    ("1.23456789E+09", "1234567890"),
    ("0001234567890", "1234567890"),
    ("PO-123#", "PO-123"),
    ("", None), (None, None), (float("nan"), None), ("nan", None),
])
def test_normalize(raw, expected):
    assert normalize_po(raw) == expected


def test_keep_leading_zeros_when_configured():
    assert normalize_po("00123", strip_leading_zeros=False) == "00123"


def test_fires_when_sets_equal_with_duplicates():
    nova = ["100", "100", "200", "300"]           # 3 únicas, 100 duplicada
    sap = ["0000000100", "200", "300", "300"]     # 3 únicas, 300 duplicada, ceros de SAP
    r = evaluate_trigger(nova, sap)
    assert r.fired
    assert (r.nova_unique, r.sap_unique) == (3, 3)
    assert r.nova_duplicates == {"100": 2} and r.sap_duplicates == {"300": 2}


def test_not_fired_count_mismatch():
    r = evaluate_trigger(["1", "2", "3"], ["1", "2"])
    assert not r.fired and "Conteo distinto" in r.reason
    assert r.missing_in_sap == ["3"]


def test_not_fired_same_count_but_different_pos():
    r = evaluate_trigger(["1", "2", "3"], ["1", "2", "4"])
    assert not r.fired and r.missing_in_sap == ["3"] and r.extra_in_sap == ["4"]


def test_not_fired_when_sap_has_extra_pos():
    r = evaluate_trigger(["1", "2"], ["1", "2", "3"])
    assert not r.fired


def test_empty_nova_never_fires_by_default():
    assert not evaluate_trigger([], []).fired
    assert evaluate_trigger([], [], allow_empty=True).fired


def test_detail_rows():
    rows = build_detail_rows(["1", "1", "2"], ["2", "3"])
    assert rows == [
        {"po": "1", "veces_nova": 2, "veces_sap": 0, "estatus": "SOLO_NOVA"},
        {"po": "2", "veces_nova": 1, "veces_sap": 1, "estatus": "MATCH"},
        {"po": "3", "veces_nova": 0, "veces_sap": 1, "estatus": "SOLO_SAP"},
    ]
