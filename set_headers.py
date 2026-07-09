import urllib.request, urllib.parse, json, time

RELAY = 'https://script.google.com/macros/s/AKfycbx_TbS9YF-XXnm1wOQVFH7_UveazRuzB-i2LKhNcA5Uvs2KTPiMjMTTGmng1axQuuXc1g/exec'

HEADERS = {
    'Command_Register': [
        'Task_ID', 'Serial_No', 'Created_Date', 'Task_Title', 'Description',
        'Category', 'Category_ID', 'Project_Name', 'Priority',
        'Assigned_To', 'Assigned_By', 'Due_Date', 'Status',
        'Completion_Date', 'Officer_Remarks', 'Staff_Remarks',
        'Points_Awarded', 'Approved_By', 'Risk_Level', 'Last_Updated'
    ],
    'Correspondence_Tracker': [
        'Corr_ID', 'Serial_No', 'Date', 'Type', 'Subject',
        'Allocated_To', 'Allocated_By', 'Current_Stage',
        'Submission_Mode', 'Completion_Date', 'Status', 'Remarks',
        'Project_Name', 'Last_Updated'
    ],
    'IGP_Command_Works': [
        'Work_ID', 'Serial_No', 'Date', 'Work_Title', 'Project_Name',
        'Description', 'Assigned_To', 'Assigned_By',
        'Priority', 'Status', 'Due_Date',
        'Completed_Date', 'Remarks', 'Approved_By', 'Last_Updated'
    ],
    'VC_Schedule': [
        'VC_ID', 'Serial_No', 'Subject', 'Date', 'Time',
        'Platform', 'Project_Name', 'VC_Link', 'Organiser', 'Speaker',
        'Allotted_Persons', 'Status', 'Linked_Postponed_ID',
        'Postpone_Reason', 'Created_By', 'Remarks', 'Entry_Date'
    ],
}

def post_json(payload_dict):
    payload = json.dumps(payload_dict).encode('utf-8')
    req = urllib.request.Request(
        RELAY, data=payload,
        headers={'Content-Type': 'application/json'},
        method='POST'
    )
    # Follow redirects manually
    opener = urllib.request.build_opener(urllib.request.HTTPRedirectHandler())
    try:
        with opener.open(req, timeout=30) as r:
            return r.read().decode('utf-8')
    except Exception as e:
        return f'ERROR: {e}'

for sheet, cols in HEADERS.items():
    print(f'\n>>> {sheet} ({len(cols)} columns)')
    print('    Columns:', ', '.join(cols))

    # Use action=update with row=1 to write header row
    payload = {
        'sheet': sheet,
        'action': 'update',
        'row': 1,
        'rowNumber': 1,
        'data': cols
    }
    result = post_json(payload)
    print('    Result:', result[:200])
    time.sleep(2)

print('\nAll done!')
