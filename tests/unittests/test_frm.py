from pathlib import Path
from openpyxl import Workbook

def test_frm_init(frm_instance):
    assert frm_instance.PREP_TIME == 5
    assert frm_instance.PPH_OFFLOAD['LOOSE_TRAILER'] == 1002
    assert frm_instance.PPH_OFFLOAD['LOOSE_VAN'] == 835
    assert frm_instance.PIECE_PER_BAG == 27
    assert frm_instance.PARCEL_PER_CAGE_DG == 16
    assert frm_instance.PARCEL_PER_CAGE_NCY == 16
    assert frm_instance.HPC == 5
    assert isinstance(frm_instance.work_book, Workbook)
    assert frm_instance.lh_schedule_sheet.title == 'Input Road LH Schedule'
    assert frm_instance.pud_schedule_sheet.title == 'Input Road PUD schedule'
    assert frm_instance.air_schedule_sheet.title == 'Input Air Schedule'
    assert frm_instance.orig_dest_sheet.title == 'Origin & Destination'
    assert frm_instance.get_air_inbound_df().empty
    assert frm_instance.get_air_outbound_df().empty

def test_get_lh_inbound_df(frm_instance):
    assert frm_instance.get_lh_inbound_df().shape == (187, 16)

def test_get_lh_outbound_df(frm_instance):
    assert frm_instance.get_lh_outbound_df().shape == (89, 16)

def test_get_pud_inbound_df(frm_instance):
    assert frm_instance.get_pud_inbound_df().shape == (38, 16)

def test_get_pud_outbound_df(frm_instance):
    assert frm_instance.get_pud_outbound_df().shape == (194, 16)

def test_get_orig_dest_df(frm_instance):
    assert frm_instance.get_orig_dest_df().shape == (497, 153)

def test_combine_inbound_df(frm_instance):
    assert frm_instance.combine_inbound_df().shape == (225, 13)

def test_combine_outbound_df(frm_instance):
    assert frm_instance.combine_outbound_df().shape == (283, 13)