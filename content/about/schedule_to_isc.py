import pandas as pd
from datetime import datetime

def create_calendar_file(input_csv, output_ics):
    
    global df 
    global row
    
    # Read the CSV file
    df = pd.read_csv(input_csv)
    
    # Filter out rows based on criteria
    df = df[
        (~df['Hbool'].isin(['Header', 'Extra-Header', 'Faculty', ''])) & 
        (df['Task or Topic'].notna()) & 
        (df['Task or Topic'] != '')
    ]
    
    # Start the ICS file content
    calendar_content = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "CALSCALE:GREGORIAN",
        "X-WR-CALNAME:FIN377",
        "X-WR-TIMEZONE:America/New_York"
    ]
    
    # Process each event
    for _, row in df.iterrows():
        # Convert date from MM/DD/YYYY to YYYYMMDD
        date_obj = datetime.strptime(row['Date'], '%m/%d/%Y')
        date_str = date_obj.strftime('%Y%m%d')

        # print(type(row))
        # print(row.loc['HBool'])
        # print('ierwuhfieuhrf')
        # eiruhf
        
        # Determine event type and time
        if 'ASGN' == row['Hbool']:
            start_time = '190000'  # 7:00 PM
            end_time = '200000'    # 8:00 PM
            color = '#FF6B6B'      # Tomato red
        elif 'Lecture' == row['Hbool']:
            start_time = '121000'  # 12:10 PM
            end_time = '132500'    # 1:25 PM
            color = '#98B4A6'      # Sage green
        else:
            start_time = '090000'  # 9:00 AM
            end_time = '093000'    # 9:30 AM
            color = '#FFB347'      # Tangerine        
        # Create event
        event = [
            "BEGIN:VEVENT",
            f"DTSTART:{date_str}T{start_time}",
            f"DTEND:{date_str}T{end_time}",
            f"SUMMARY:{row['Task or Topic']}",
            f"DESCRIPTION:Module {row['Module']}, Week {row['Week']}",
            f"COLOR:{color}",
            "END:VEVENT"
        ]
        
        calendar_content.extend(event)
    
    # Close the calendar
    calendar_content.append("END:VCALENDAR")
    
    # Write to file
    with open(output_ics, 'w') as f:
        f.write('\n'.join(calendar_content))

# Example usage
if __name__ == "__main__":
    create_calendar_file('ScheduleExperiment.csv', 'fin377_calendar.ics')