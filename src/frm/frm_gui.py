#from tkinter import *
#from tkinter import ttk
#import frm

#root = Tk()
#ttk.Label(root, text='Hello GUI World').pack(side=TOP, fill=Y)
#ttk.Button(root, text='Run Model').pack(side=BOTTOM)
#root.title('Functional Requirement Module')
#root.mainloop()
#widget = Label(None, text='Hello GUI World')
#widget.pack()
#widget.mainloop()

import datetime
import streamlit as st
import pandas as pd
import os
import frm

st.title('Functional Requirement Module')

folder = st.sidebar.text_input('Type your folder path here...')

def file_selector(folder_path, filetype='schedule'):
    if folder_path == '':
        folder_path = '.'
    filenames = [None] + os.listdir(folder_path)
    selected_filename = st.sidebar.selectbox(f'Select a {filetype} file', filenames)
    return os.path.join(folder_path, selected_filename) if type(selected_filename) == str else None 

schedule_file = file_selector(folder, 'schedule')
st.write(f'You selected schedule: {schedule_file}')

config_file = file_selector(folder, 'configuration')
st.write(f'You selected configuration: {config_file}')

if schedule_file is not None and schedule_file.lower().endswith(('.xls', '.xlsx')):
    frm_instance = frm.FRM(excel_file=schedule_file, config_file=config_file)
    
    # select schedule type for the hourly volume graph
    schedule_type = st.sidebar.selectbox(
        'Select your schedule type for the hourly volume plot',
        ('linehaul', 'pud', 'air', 'all')
    )
    fig = frm_instance.hourly_volume_graph(schedule_type=schedule_type)
    st.plotly_chart(fig)

    # show the descriptive stats
    stats = frm_instance.descriptive_stats()
    stats

    # select break point (if necessary)
    inbound_ts = frm_instance.inbound_unit_coy()
    start_time = min(inbound_ts['offload time']).floor('1H')
    end_time = max(inbound_ts['finished offload time']).ceil('1H')
    time_range = pd.date_range(start_time, end_time, freq='1H')
    break_point = st.sidebar.selectbox(
        'Select your break point (if necessary)',
        [None] + time_range.to_list())
    break_point = break_point if isinstance(break_point, datetime.datetime) else None

    # show inbound doors
    st.write(f'The number of inbound doors: {frm_instance.inbound_doors_coy()}')

    # select sort windows
    inbound_ts = frm_instance.inbound_unit_coy(break_point=break_point)
    start_time = min(inbound_ts['offload time']).floor('1H')
    end_time = max(inbound_ts['finished offload time']).ceil('1H')
    time_range = pd.date_range(start_time, end_time, freq='1H')
    start_sort1 = st.sidebar.selectbox(
        'Select your start time for sort window 1',
        time_range.to_list())
    end_sort1 = st.sidebar.selectbox(
        'Select your end time for sort window 1',
        time_range.sort_values(ascending=False).to_list())
    start_sort2 = st.sidebar.selectbox(
        'Select your start time for sort window 2 (if necessary)',
        [None] + time_range.to_list())
    end_sort2 = st.sidebar.selectbox(
        'Select your end time for sort window 2 (if necessary)',
        [None] + time_range.to_list())
    sort_window1 = (str(start_sort1), str(end_sort1))
    if isinstance(start_sort2, datetime.datetime) and isinstance(end_sort2, datetime.datetime):
        sort_window2 = (str(start_sort2), str(end_sort2))

    st.write(f'Your sort windows 1 starts at {sort_window1[0]} ends at {sort_window1[1]}')
    st.write('Sort Capacity COY (sorter simulation per hour based on linear programming algorithms):')
    sort1 = frm_instance.sort_capacity_coy(break_point=break_point, sort_window=sort_window1, ts='1H')
    sort1
    st.write('The Number of outbound doors:')
    outbound_doors1 = frm_instance.outbound_doors_coy(break_point=break_point, sort_window=sort_window1, ts='1H')
    outbound_doors1.index.name = 'unique destination'
    outbound_doors1.name = '#doors'
    outbound_doors1.rename(index={'doors per 1H': 'total'})
    outbound_doors1
    if frm_instance.air_schedule_sheet:
        dist1 = frm_instance.schedule_type_dist(sort_window=sort_window1, ts='1H')
        st.write('The Number of outbound doors:')
        dist1

    if 'sort_window2' in locals():
        st.write(f'Your sort windows 1 starts at {sort_window2[0]} ends at {sort_window2[1]}')
    
    

    output_path = st.sidebar.text_input('Type your output folder here...')
    if st.button('Export Results') and 


if __name__ == '__main__':
    os.system('streamlit run frm_gui.py')


