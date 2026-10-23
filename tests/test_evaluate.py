

def test_excel_round_on_huge_numbers():
    from excel_twin.evaluate import excel_round
    assert excel_round(1e20, 10) == 1e20
    assert excel_round(34.245, 2) == 34.25
