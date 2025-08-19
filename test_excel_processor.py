import pandas as pd
import os
from datetime import datetime, timedelta

# Create a test Excel file with multiple date sheets
def create_test_excel(file_path):
    # Create a Pandas Excel writer
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        # Create 5 days of data
        for day in range(5):
            # Create a date string in dd.mm.yyyy format
            date = (datetime.now() - timedelta(days=4-day)).strftime('%d.%m.%Y')
            
            # Create sample data
            data = {
                'Name': [f'Agent {i+1}' for i in range(5)],
                'Team Leader': ['John', 'John', 'Sarah', 'Sarah', 'Mike'],
                'Login Time': ['08:45:00', '09:00:00', '08:50:00', '09:15:00', '08:30:00'],
                'Handled Inbound': [65, 50, 70, 45, 55],
                'AHT': ['00:05:30', '00:07:45', '00:04:50', '00:08:30', '00:06:15'],
                'Occupancy': [0.75, 0.72, 0.80, 0.65, 0.78],
                'Bio break': ['00:10:00', '00:15:00', '00:08:00', '00:12:00', '00:09:00'],
                'Lunch': ['00:30:00', '00:35:00', '00:28:00', '00:32:00', '00:29:00'],
                'Tea break': ['00:15:00', '00:12:00', '00:18:00', '00:20:00', '00:16:00'],
                'Working Rate': [0.85, 0.78, 0.90, 0.72, 0.88]
            }
            
            # Create a DataFrame and write to Excel
            df = pd.DataFrame(data)
            df.to_excel(writer, sheet_name=date, index=False)
    
    print(f"Test Excel file created at: {file_path}")

if __name__ == "__main__":
    # Create the test file in the same directory as the main script
    test_file = os.path.join(os.path.dirname(__file__), 'test_data.xlsx')
    create_test_excel(test_file)
    
    print("\nTo test the application:")
    print(f"1. Run: python excel_processor.py")
    print(f"2. Click 'Browse...' and select: {test_file}")
    print("3. Select a team leader from the dropdown")
    print("4. Set the date range")
    print("5. Click 'Process Data' to generate the output")
