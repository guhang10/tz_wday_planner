#!/usr/bin/python3
from datetime import datetime, timedelta, timezone
import pytz
import json
import pandas as pd
import calendar


## range specified as EST datetime
range = {
    'name': 'NA',
    'start': '2022-12-05T00:00:00-0500',
    'end': '2022-12-11T23:59:59-0500',
    'format': '%Y-%m-%dT%H:%M:%S%z',
}


## location #1 EST time
locations = [
    {
        'name': 'NA',
        'tz': timezone(timedelta(hours=-5)),
        'whours': {
            0: ['08:00', '18:00'],  # mon
            1: ['08:00', '18:00'],  # tue
            2: ['08:00', '18:00'],  # wed
            3: ['08:00', '18:00'],  # thu
            4: ['08:00', '12:00'],  # fri
        }
    },
    {
        'name': 'APAC_shift_1',
        'tz': timezone(timedelta(hours=+10)),
        'whours': {
            0: ['08:00', '17:30'],  # mon
            1: ['08:00', '17:30'],  # tue
            2: ['08:00', '17:30'],  # wed
            3: ['08:00', '17:30'],  # thu
            4: ['08:00', '12:00'],  # fri
        }
    },
    {
        'name': 'APAC_shift_2',
        'tz': timezone(timedelta(hours=+10)),
        'whours': {
            0: ['11:00', '20:30'],  # mon
            1: ['11:00', '20:30'],  # tue
            2: ['11:00', '20:30'],  # wed
            3: ['11:00', '20:30'],  # thu
            4: ['08:00', '12:00'],  # fri
        }
    },
    {
        'name': 'UK',
        'tz': timezone(timedelta(hours=+00)),
        'whours': {
            0: ['08:00', '17:30'],  # mon
            1: ['08:00', '17:30'],  # tue
            2: ['08:00', '17:30'],  # wed
            3: ['08:00', '17:30'],  # thu
            4: ['08:00', '12:00'],  # fri
        }
    },
    {
        'name': 'NL',
        'tz': timezone(timedelta(hours=+1)),
        'whours': {
            0: ['08:00', '17:30'],  # mon
            1: ['08:00', '17:30'],  # tue
            2: ['08:00', '17:30'],  # wed
            3: ['08:00', '17:30'],  # thu
            4: ['08:00', '12:00'],  # fri
        }
    }
]

# timestamp alignment, should be 30% for this purpose
def ceil_dt(dt, delta):
    original_tz = dt.tzinfo
    if original_tz:
        dt = dt.astimezone(pytz.UTC)
        dt = dt.replace(tzinfo=None)
    dt = dt + ((datetime.min - dt) % delta)
    if original_tz:
        dt = pytz.UTC.localize(dt)
        dt = dt.astimezone(original_tz)
    return dt

# test if a timestamp is in a range (same day)
def is_in_range(time, range):
    time_sec = time.split(':')[0]*3600 + time.split(':')[1]*60
    range_start = range[0].split(':')[0]*3600 + range[0].split(':')[1]*60
    range_end = range[1].split(':')[0]*3600 + range[1].split(':')[1]*60
    if range_start <= time_sec <= range_end:
        return True
    else:
        return False

# work out the range in 30 min epoch blocks
range_start = ceil_dt(datetime.strptime(range['start'], range['format']), timedelta(minutes=30))
range_end = ceil_dt(datetime.strptime(range['end'], range['format']), timedelta(minutes=30))

blocks = [range_start]
block_time = range_start
block_days = [calendar.day_name[range_start.weekday()]]

while block_time < range_end:
    block_time = block_time + timedelta(minutes=30)
    blocks.append(block_time)
    block_days.append(calendar.day_name[block_time.weekday()])

local_blocks = {f'Week-days ({range["name"]})': [{'time': ref_day} for ref_day in block_days]}

for location in locations:
    local_blocks[location['name']] = []
    for b_index,block in enumerate(blocks):
        local_block_time = block.astimezone(location['tz'])
        local_day = local_block_time.weekday()
        local_time = local_block_time.strftime('%H:%M')
        local_block = {
            'time': local_time,
            'day': local_day,
            'in_office': local_day in location['whours'] and is_in_range(local_time, location['whours'][local_day])
        }
        ## add a '_w' to time if it's working hours, used for styling later
        if local_block['in_office']:
            local_block['time']
        local_blocks[location['name']].append(local_block)

## create the dataframe
df = pd.DataFrame([[block['time'] for block in local_blocks[location]] for location in local_blocks],
                   index=[location for location in local_blocks])

## get index mapping of all in hour cells
highlighted_cells = []
for row_index,blocks in enumerate(local_blocks.values()):
    for column_index,block in enumerate(blocks):
        if 'in_office' in block and block['in_office']:
            highlighted_cells.append([row_index, column_index])

def style_specific_cell(x):
    color = 'background-color: green; color: white'
    df1 = pd.DataFrame('', index=x.index, columns=x.columns)
    for cell_addr in highlighted_cells:
        df1.iloc[cell_addr[0], cell_addr[1]] = color
    return df1

def style_weekday_cell(x):
    ## weekday color code
    week_day_colors = {
        'Monday': 'background-color: #9fb5a5',
        'Tuesday': 'background-color: #b0b9d6',
        'Wednesday': 'background-color: #c7d6b0',
        'Thursday': 'background-color: #edecb9',
        'Friday': 'background-color: #c6b0d6',
        'Saturday': 'background-color: #D3D3D3',
        'Sunday': 'background-color: #D3D3D3'
    }
    if x in week_day_colors:
        return week_day_colors[x]

df.to_excel('working_time_by_regions_NA.xlsx', sheet_name='working_time_by_regions_NA')
## apply style
df.style.apply(style_specific_cell, axis=None).applymap(style_weekday_cell).to_excel('working_time_by_regions_NA.xlsx', freeze_panes=(0, 1))

