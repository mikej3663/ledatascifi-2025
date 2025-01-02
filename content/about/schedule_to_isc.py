import pandas as pd
from datetime import datetime

def create_event_with_alarms(date_str, start_time, end_time, summary, description, event_type):
    event = [
        "BEGIN:VEVENT",
        f"DTSTART:{date_str}T{start_time}",
        f"DTEND:{date_str}T{end_time}",
        f"SUMMARY:{summary}",
        f"DESCRIPTION:{description}",
        f"UID:fin377-{date_str}-{start_time}"
    ]
    
    # Add appropriate alarms based on event type
    if event_type == 'ASGN':
        # 1 day before
        event.extend([
            "BEGIN:VALARM",
            "TRIGGER:-P1D",
            "ACTION:DISPLAY",
            "DESCRIPTION:Assignment due in 1 day",
            "END:VALARM",
            # 1 hour before
            "BEGIN:VALARM",
            "TRIGGER:-PT1H",
            "ACTION:DISPLAY",
            "DESCRIPTION:Assignment due in 1 hour",
            "END:VALARM"
        ])
    elif event_type == 'Lecture':
        # 15 minutes before
        event.extend([
            "BEGIN:VALARM",
            "TRIGGER:-PT15M",
            "ACTION:DISPLAY",
            "DESCRIPTION:Class starts in 15 minutes",
            "END:VALARM"
        ])
    else:  # Tasks
        # 1 hour before
        event.extend([
            "BEGIN:VALARM",
            "TRIGGER:-PT1H",
            "ACTION:DISPLAY",
            "DESCRIPTION:Task due before FIN377 class starts",
            "END:VALARM"
        ])
    
    event.append("END:VEVENT")
    return event

def create_calendar_file(df_subset, output_ics, start_time, end_time, calendar_name):
    # Start the ICS file content
    calendar_content = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "CALSCALE:GREGORIAN",
        f"X-WR-CALNAME:{calendar_name}",
        "X-WR-TIMEZONE:America/New_York",
        "METHOD:PUBLISH"
    ]
    
    # Process each event
    for _, row in df_subset.iterrows():
        # Convert date from MM/DD/YYYY to YYYYMMDD
        # date_obj = datetime.strptime(row['Date'], '%m/%d/%Y')
        date_str = row['Date'].strftime('%Y%m%d')
        
        # Create description with hyperlink if it exists
        description = ""
        if pd.notna(row['Hyperlink']) and row['Hyperlink'] != '':
            description = f"{row['Hyperlink']}"
        
        # Create event with appropriate alarms
        event = create_event_with_alarms(
            date_str,
            start_time,
            end_time,
            row['Task or Topic'],
            description,
            row['Hbool']
        )
        
        calendar_content.extend(event)
    
    # Close the calendar
    calendar_content.append("END:VCALENDAR")
    
    # Write to file
    with open(output_ics, 'w', encoding='utf-8') as f:
        f.write('\n'.join(calendar_content))

def create_all_calendars(input_xlsx):
    # Read the CSV file
    df = pd.read_excel(input_xlsx,sheet_name='Overall')
    # df['Date'] = pd.to_datetime(df['Date'])
    
    # Filter out rows based on base criteria
    df = df[
        (~df['Hbool'].isin(['Header', 'Extra-Header', 'Faculty', ''])) & 
        (df['Task or Topic'].notna()) & 
        (df['Task or Topic'] != '')
    ]
    
    # Create assignments calendar
    assignments = df[df['Hbool'] == 'ASGN']
    create_calendar_file(
        assignments, 
        'fin377_due_dates.ics',
        '190000',  # 7:00 PM
        '200000',  # 8:00 PM
        'FIN377 due dates'
    )
    
    # Create Class1 calendar (12:10-1:25)
    class1 = df[df['Hbool'] == 'Lecture']
    create_calendar_file(
        class1,
        'fin377_class1.ics',
        '121000',  # 12:10 PM
        '132500',  # 1:25 PM
        'FIN377 class1'
    )
    
    # Create Class2 calendar (1:35-2:50)
    class2 = df[df['Hbool'] == 'Lecture']
    create_calendar_file(
        class2,
        'fin377_class2.ics',
        '133500',  # 1:35 PM
        '145000',  # 2:50 PM
        'FIN377 class2'
    )
    
    # Create Tasks calendar
    tasks = df[(df['Hbool'] != 'ASGN') & (df['Hbool'] != 'Lecture')]
    create_calendar_file(
        tasks,
        'fin377_tasks.ics',
        '090000',  # 9:00 AM
        '093000',  # 9:30 AM
        'Tasks due for FIN377'
    )

# Example usage
if __name__ == "__main__":
    create_all_calendars('Schedule.xlsx')