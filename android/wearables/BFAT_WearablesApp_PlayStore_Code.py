#Project Name : boAt - App Feedback Analyzer Tool
#Author       : boAt R&D (QA)
#Description  : Google Play Developer API based boAt App Feedback Analyzer
#Note         : Legacy Selenium web-scraping functions have been removed.
#               Review fetching is now handled in main.py via the official
#               Google Play Developer API. Only post-processing / Excel
#               formatting functions remain below.
#********************************************************************#

import os
import re
import openpyxl
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment

# Global variables declaration
global App_Rating
global App_Version
global lowest_date
global latest_date

#***********************************************************************#

def Positive_Reviews():
    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel('Android_WearablesApp_Review.xlsx')
    
    # Filter the DataFrame to keep only rows with ratings 4 and 5
    positive_reviews_df = df[df['Ratings'].isin([4, 5])]
    
    # Save the positive reviews to the same Excel file in a new sheet named 'Positive_Reviews'
    with pd.ExcelWriter('Android_WearablesApp_Review.xlsx', engine='openpyxl', mode='a') as writer:
        positive_reviews_df.to_excel(writer, sheet_name='Positive_Reviews', index=False)
    
    print("Positive reviews saved to 'Android_WearablesApp_Review.xlsx' under sheet name 'Positive_Reviews'.")


def Negative_Reviews():
    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel('Android_WearablesApp_Review.xlsx')

    # Filter the DataFrame to keep only rows with ratings 1, 2, and 3
    negative_reviews_df = df[df['Ratings'].isin([1, 2, 3])].copy()  # Make a copy to avoid modifying the original DataFrame
    
    # Keyword patterns
    keyword_patterns = {
    'Connectivity': r'\b(connect(ion|ing)?|not\s*connected|not\s*getting\s*connect(ed|ion)?|unable\s*to\s*connect|bad\s*connectivity)\b',
    'Device disconnection': r'\bdisconnected?\b',
    'EQ Customization': r'\b(EQ|equalizer)\s*(customization)?\b',
    'Login': r'\b(login|log(in|ged)?|otp|sign\s*in|signing\s*in)\b',
    'Location Permission': r'\b(location|gps|geolocation)\s*(permission|access)?\b',
    'Phone becoming silent': r'\bsilent\b',
    'Vibration Mode': r'\bvibration\s*mode|vibrate\b',
    'SKU Not supported': r"\b(can't\s*connect.*app|not\s*(connect(ing)?|support(ed)?|show(n|ing)|found|recognized|compatible|listed|list|option)\s*(by\s*this\s*app)?|\bthis\s*app\s*doesn't\s*support\s*all\s*boat\s*Wearables|\bairdopes\s*doesn't\s*pair\s*and\s*show\s*in\s*the\s*app|\bboat\s*headphones\s*not\s*listed|\bdevices\s*could\s*not\s*found|\bnot\s*available\s*in\s*list|\bsupported\s*device\s*list\b|\bthis\s*model\s*is\s*not\s*(listed|shown|supported|found)\s*in\s*the\s*app\b|\bthis\s*model\s*does\s*not\s*show\s*this\s*model\s*on\s*app\b|\bcannot\s*pair\s*it\s*with\s*this\s*app\b|\bnot\s*listed\s*in\s*app\b|\bnot\s*list\b|\bno\s*option\b|\boption\s*is\s*not\s*available\s*in\s*this\s*app\b|can't\s*see\b|\bdidn'?t\s*find\s*my\s*device\b|\bdoes\s*not\s*show\s*other\s*models\b|\bunable\s*to\s*find\s*my\s*product\b|\bdidn'?t\s*found\s*my\s*device\b|\bcan't\s*found\s*my\s*device\b|\bno\s*support\s*for\s*old\s*devices\b|\bnot\s*supported\s*for\s*old\s*devices\b)",  # Added 'No support for old devices' and 'Not supported for old devices'
    'In Ear Detection (Dev)': r'\bin\s*ear\s*detection|ear\s*detection\b',
    'Auto Pause (dev)': r'\b(auto.*pause|pause.*auto)\b',
    'Positive Feedback': r'^(?!.*(not|never|no)).*\b(good|super|nice|nyc|best|beautiful|great|amazing|awesome|impressed)\b',
    'Powering off (Device)': r'\b(auto\s*on\s*off|powering\s*off|auto\s*shutdown)\b',
    'Crash': r'\b(crash(ing)?|hang(ing)?|freeze(ing)?|carshed)\b',
    'Package box (Device)': r'\b(open\s*box|package\s*box)\b',
    'Battery': r'\b(battery|battery\s*percentage|battery\s*level|case\s*battery|earbud\s*battery)\b',
    'Calibration': r'\b(calibration|3D)\b',
    'Firmware': r'\b(firmware|firm\s*ware)\b'
}


    # Compile patterns with re.IGNORECASE flag
    compiled_patterns = {keyword: re.compile(pattern, re.IGNORECASE) for keyword, pattern in keyword_patterns.items()}
    
    # Analyze reviews and boAt replies
    reasons = []
    final_reasons = []
    boat_reply_reasons = []  # To store boAt Reply reasons

    # Loop through the reviews and boAt replies
    for review, reply in zip(negative_reviews_df['Reviews'], negative_reviews_df['boAt Reply']):
        found_reasons = []
        found_reply_reasons = []

        # Ensure reply is a string (replace NaN or None with empty string)
        if pd.isna(reply):
            reply = ''

        # Check patterns in the reviews
        for keyword, pattern in keyword_patterns.items():
            if re.search(pattern, review, re.IGNORECASE):
                found_reasons.append(keyword)

        reasons.append(', '.join(found_reasons) if found_reasons else 'Generic Comments')

        # Check patterns in the boAt replies
        for keyword, pattern in keyword_patterns.items():
            if re.search(pattern, reply, re.IGNORECASE):
                found_reply_reasons.append(keyword)

        boat_reply_reasons.append(', '.join(found_reply_reasons) if found_reply_reasons else 'Generic Reply')

        # Final reason logic:
        # If the boAt reply contains "SKU Not supported", override the final reason with "SKU Not supported"
        if 'SKU Not supported' in found_reply_reasons:
            final_reasons.append('SKU Not supported')
        elif len(found_reasons) > 1:
            final_reasons.append('Decision yet to be confirmed')
        else:
            final_reasons.append(', '.join(found_reasons) if found_reasons else 'Generic Comments')

    # Add the columns to the DataFrame
    negative_reviews_df['Reason'] = reasons
    #negative_reviews_df['boAt Reply Reasons'] = boat_reply_reasons
    negative_reviews_df['Consolidated Reason'] = final_reasons

    # Save the negative reviews with reasons to the same Excel file in a new sheet named 'Negative_Reviews'
    with pd.ExcelWriter('Android_WearablesApp_Review.xlsx', engine='openpyxl', mode='a') as writer:
        negative_reviews_df.to_excel(writer, sheet_name='Negative_Reviews', index=False)

    print("Negative reviews and reasons saved to 'Negative_Reviews' sheet.")
    

def SheetNameChange():
    directory_path = os.path.dirname(os.path.realpath(__file__))  # Get the current script's directory
    # Load the Excel file
    file_path = 'Android_WearablesApp_Review.xlsx'
    current_sheet_name = 'Sheet1'
    new_sheet_name = 'WearablesApp_Review_Data'

    try:
        # Load the Excel file
        wb = openpyxl.load_workbook(file_path)
        
        # Check if the current sheet name exists
        if current_sheet_name in wb.sheetnames:
            # Get the sheet object
            sheet = wb[current_sheet_name]
            
            # Change the sheet name to 'Negative_Reviews'
            sheet.title = new_sheet_name
            
            # Save the modified workbook with the new sheet name
            wb.save(file_path)
            
            print(f"The sheet name '{current_sheet_name}' has been changed to '{new_sheet_name}'.")
        else:
            print(f"The sheet '{current_sheet_name}' does not exist in the Excel file.")
        
        # Close the Excel file
        wb.close()

    except FileNotFoundError:
        print(f"The file '{file_path}' does not exist.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")   



def Sorting_Xlsx(sheet_name):
    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel('Android_WearablesApp_Review.xlsx', sheet_name=sheet_name)

    # Load the XLSX file
    workbook = openpyxl.load_workbook('Android_WearablesApp_Review.xlsx')

    # Select the worksheet where you want to set column widths
    worksheet = workbook[sheet_name]

    # Define the column widths you want to set (e.g., columns A, B, and C)
    column_widths = {'A': 15, 'B': 15, 'C': 8, 'D': 42, 'E': 12.5, 'F': 11, 'G': 11, 'H': 10, 'I': 10, 'J': 48, 'K': 15, 'L': 21} # You can adjust the values as needed

    # Set the column widths
    for column, width in column_widths.items():
        worksheet.column_dimensions[column].width = width

    # Iterate through the rows and columns to apply wrap text and middle align
    for row in worksheet.iter_rows():
        for cell in row:
            # Set wrap text to True
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="center")

    # Save the workbook with changes to the same file
    workbook.save('Android_WearablesApp_Review.xlsx')

    # Close the workbook
    workbook.close()

    print(f"Xlsx modification done for sheet '{sheet_name}'.")
    

def Sorting_Summary_Xlsx(sheet_name):
    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel('Android_WearablesApp_Review.xlsx', sheet_name=sheet_name)

    # Load the XLSX file
    workbook = openpyxl.load_workbook('Android_WearablesApp_Review.xlsx')

    # Select the worksheet where you want to set column widths
    worksheet = workbook[sheet_name]

    # Define the column widths you want to set (e.g., columns A, B, and C)
    column_widths = {'A': 25, 'B': 15, 'C': 21, 'D': 23, 'E': 25, 'F': 28} # You can adjust the values as needed

    # Set the column widths
    for column, width in column_widths.items():
        worksheet.column_dimensions[column].width = width

    # Iterate through the rows and columns to apply wrap text and middle align
    for row in worksheet.iter_rows():
        for cell in row:
            # Set wrap text to True
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="center")

    # Save the workbook with changes to the same file
    workbook.save('Android_WearablesApp_Review.xlsx')

    # Close the workbook
    workbook.close()

    print(f"Xlsx modification done for sheet '{sheet_name}'.")
    
        

def Review_Date_Calculator():
    global lowest_date
    global latest_date
    # Load the Excel file
    df = pd.read_excel('Android_WearablesApp_Review.xlsx')

    # Replace 'DateTimeColumn' with the actual name of the column containing your date and time data
    df['Date Time'] = pd.to_datetime(df['Date Time'])

    # Extract date part
    df['DateOnly'] = df['Date Time'].dt.date

    # Find the earliest and latest date
    lowest_date = df['DateOnly'].min()
    latest_date = df['DateOnly'].max()


    # Print the results
    print("Lowest Date in the DateTimeColumn:", lowest_date)
    print("Latest Date in the DateTimeColumn:", latest_date)



def Negative_Review_Analysis():
    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel('Android_WearablesApp_Review.xlsx', sheet_name='Negative_Reviews')

    # Count occurrences of each "Consolidated Reason"
    reason_counts = df['Consolidated Reason'].value_counts()

    # Print the counts for each "Consolidated Reason"
    print("Negative Review Analysis:")
    for reason, count in reason_counts.items():
        print(f"{reason}: {count}")
        
    return reason_counts


def Summary():
    global App_Rating
    global App_Version
    global lowest_date
    global latest_date
        
    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel('Android_WearablesApp_Review.xlsx')

    # Calculate review counts
    positive_reviews_count = len(df[df['Ratings'].isin([4, 5])])
    negative_reviews_count = len(df[df['Ratings'].isin([1, 2, 3])])

    # Calculate summary statistics
    app_name = "Android Wearables App"  # Replace with your app name
    app_rating = App_Rating + '*'
    app_version = App_Version
    review_data = str(lowest_date) + ' to ' + str(latest_date)

    # Define app details data
    app_details_data = {
        'App Name': [app_name],
        'App Rating': [app_rating],
        'Latest App Version': [app_version],  # Keep the app version unchanged
        'Review Date': [review_data]
    }

    # Create the app details DataFrame
    app_details_df = pd.DataFrame(app_details_data)
    
    # Create a DataFrame for the review counts
    posneg_review_counts_data = {
        'Category': ['Positive Reviews (4 & 5)', 'Negative Reviews (<=3)'],
        'Count': [positive_reviews_count, negative_reviews_count]
    }
    posneg_review_counts_df = pd.DataFrame(posneg_review_counts_data)
    
    # Create a DataFrame for the negative review analysis
    negative_reason_counts = Negative_Review_Analysis()
    negative_review_counts_df = pd.DataFrame({'Summary (-ve reviews)': negative_reason_counts.index, 'Count': negative_reason_counts.values})

    # Count the occurrences of 'Positive Feedback' in negative review analysis
    positive_feedback_count = negative_reason_counts.get('Positive Feedback', 0)

    # Ensure 'Count' column can hold string values
    negative_review_counts_df['Count'] = negative_review_counts_df['Count'].astype(object)

    # Update the 'Latest App Version' column for the 'Positive Feedback' row with the feedback message
    if positive_feedback_count > 0:
        feedback_message = f"Reviews are good but ratings are <= 3"
        
        # Find the row for 'Positive Feedback' in negative review counts DataFrame
        positive_feedback_row_index = negative_review_counts_df[negative_review_counts_df['Summary (-ve reviews)'] == 'Positive Feedback'].index[0]

        # Update the 'Latest App Version' column of 'Positive Feedback' with the feedback message
        negative_review_counts_df.at[positive_feedback_row_index, ''] = feedback_message  # Set the feedback message

        # Append this message to the Latest App Version in the app details
        app_details_df.at[0, 'Latest App Version'] = app_version  # Keep the app version unchanged
        #app_details_df.at[0, 'Latest App Version'] += f" (See feedback in negative reviews)"  # Optional: Add a note

    # Load the workbook
    wb = load_workbook('Android_WearablesApp_Review.xlsx')

    # Check if the 'Summary' sheet exists
    if 'Summary' in wb.sheetnames:
        ws = wb['Summary']
        # Append app details and review counts to 'Summary' sheet
        for r in dataframe_to_rows(app_details_df, index=False):
            ws.append(r)
        for row in dataframe_to_rows(posneg_review_counts_df, index=False):
            ws.append(row)
        ws.append([])  # Add an empty row
        for row in dataframe_to_rows(negative_review_counts_df, index=False):
            ws.append(row)
        
    else:
        # Create a new 'Summary' sheet and write app details and review counts
        ws = wb.create_sheet(title='Summary')
        for r in dataframe_to_rows(app_details_df, index=False):
            ws.append(r)
        ws.append([])  # Add an empty row
        for row in dataframe_to_rows(posneg_review_counts_df, index=False):
            ws.append(row)
        ws.append([])  # Add an empty row
        for row in dataframe_to_rows(negative_review_counts_df, index=False):
            ws.append(row)
    
    # Save the workbook
    wb.save('Android_WearablesApp_Review.xlsx')
    print('Summary sheet has been updated with app details, review counts, and negative review analysis.')


def Summary_Enhance():
    # Load the workbook
    wb = load_workbook('Android_WearablesApp_Review.xlsx')

    # Select the "Summary" sheet
    sheet = wb['Summary']

    # Make the 1st row from column A to D bold and set background color to blue
    for col in range(1, 5):  # Columns A to D
        cell = sheet.cell(row=1, column=col)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="B9D9EB", end_color="B9D9EB", fill_type="solid")
        #cell.fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")

    # Apply center alignment to columns B, C, and D
    for col in range(2, 5):  # Columns B to D
        for row in range(1, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=col)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            
    # Make the 4th row from column A to D bold and set background color
    for col in range(1, 3):  # Columns A to D
        cell = sheet.cell(row=4, column=col)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        
    # Make the 8th row from column A to D bold and set background color to yellow
    for col in range(1, 3):  # Columns A to D
        cell = sheet.cell(row=6, column=col)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Make the 8th row from column A to D bold and set background color
    for col in range(1, 3):  # Columns A to D (1 to 4 inclusive)
        cell = sheet.cell(row=5, column=col)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="66CC66", end_color="66CC66", fill_type="solid")


    # Make the 8th row from column A to D bold and set background color
    for col in range(1, 4):  # Columns A to D
        cell = sheet.cell(row=8, column=col)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Save the workbook (same file)
    wb.save('Android_WearablesApp_Review.xlsx')
    print("Summary Sheet's Last changes has been updated")


def close_excel_file(file_path):
    """
    Close the Excel file to release any locks on it.
    """
    try:
        workbook = openpyxl.load_workbook(file_path)
        workbook.close()
        print("Closed the Excel file.")
    except Exception as e:
        print(f"Error closing the Excel file: {e}")


