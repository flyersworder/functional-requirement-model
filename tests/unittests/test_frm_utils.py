import datetime
from frm import find_keyword, ws_to_df

def test_find_keyword_lh(sample_lh_work_sheet):
    assert find_keyword(sample_lh_work_sheet, 'origin', 1) == (3,1)
    assert find_keyword(sample_lh_work_sheet, 'arrival time', 1) == (3,3)
    assert find_keyword(sample_lh_work_sheet, 'sum', 1) == (3,16)
    assert find_keyword(sample_lh_work_sheet, 'destination', 18) == (3,31)
    assert find_keyword(sample_lh_work_sheet, 'sum', 18) == (3,46)

def test_find_keyword_pud(sample_pud_work_sheet):
    assert find_keyword(sample_pud_work_sheet, 'origin', 1) == (3,1)
    assert find_keyword(sample_pud_work_sheet, 'arrival time', 1) == (3,3)
    assert find_keyword(sample_pud_work_sheet, 'sum', 1) == (3,16)
    assert find_keyword(sample_pud_work_sheet, 'destination', 18) == (3,32)
    assert find_keyword(sample_pud_work_sheet, 'sum', 18) == (3,47)

def test_ws_to_df(sample_lh_work_sheet):
    df_inbound = ws_to_df(sample_lh_work_sheet, 3, 1, 16)
    df_outbound = ws_to_df(sample_lh_work_sheet, 3, 31, 46)
    assert df_inbound.shape[0] == 187
    assert df_outbound.shape[0] == 89
    assert isinstance(df_inbound['arrival time'][0], datetime.datetime)
    assert isinstance(df_outbound['departure time'][0], datetime.datetime)
