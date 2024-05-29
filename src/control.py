import argparse
import os
import pandas as pd
import win32com.client as win32

def get_open_excel_workbook(workbook_name):
    """
    Connect to an already open Excel workbook.

    :param workbook_name: The name of the workbook to get (without the extension).
    :return: The workbook object if found, otherwise None.
    """
    try:
        excel = win32.GetObject(None, 'Excel.Application')
        for workbook in excel.Workbooks:
            if workbook_name in workbook.Name:
                return workbook
        return None
    except Exception as e:
        print(f"Error: {e}")
        return None

def open_excel_workbook(path):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    workbook = excel.Workbooks.Open(path)
    return workbook

def close_excel_workbook(workbook):
    excel = workbook.Application
    workbook.Close(SaveChanges=False)
    excel.Quit()

def save_operation(control_path, mode, map, playbook, files_folder):
    df = pd.read_excel(control_path, sheet_name='playground', skiprows=13, usecols="G:N")
    sanitized_mode = mode.replace('|', '-')
    sanitized_map = map.replace('|', '-')
    file_name = f"{sanitized_map}@{sanitized_mode}@{playbook}.csv"
    file_path = os.path.join(files_folder, file_name)
    df.to_csv(file_path, index=False)

def load_operation(control_path, mode, map, playbook, files_folder):
    sanitized_mode = mode.replace('|', '-')
    sanitized_map = map.replace('|', '-')
    csv_file_path = os.path.join(files_folder, f"{sanitized_map}@{sanitized_mode}@{playbook}.csv")
    
    if not os.path.exists(csv_file_path):
        csv_file_path = os.path.join(files_folder, "eraser.csv")
    
    df = pd.read_csv(csv_file_path)
    
    wb = get_open_excel_workbook("control")
    if wb is None:
        wb = open_excel_workbook(control_path)
    
    sheet = wb.Sheets("playground")
    
    for r_idx, row in df.iterrows():
        for c_idx, value in enumerate(row):
            if pd.isna(value):
                sheet.Cells(r_idx + 15, c_idx + 7).Value = ""
            else:
                sheet.Cells(r_idx + 15, c_idx + 7).Value = value
    
    wb.Save()

def main():
    parser = argparse.ArgumentParser(description='Process some parameters.')
    parser.add_argument('--mode', type=str, required=True, help='Mode of operation')
    parser.add_argument('--map', type=str, required=True, help='Map parameter')
    parser.add_argument('--playbook', type=str, required=True, help='Playbook parameter')
    parser.add_argument('--files_folder', type=str, required=True, help='Files folder')
    parser.add_argument('--control_path', type=str, required=True, help='Path to the control Excel file')
    parser.add_argument('--direction', type=str, required=True, help='Direction (save or load)')

    args = parser.parse_args()

    if args.direction == 'save':
        save_operation(args.control_path, args.mode, args.map, args.playbook, args.files_folder)
    elif args.direction == 'load':
        load_operation(args.control_path, args.mode, args.map, args.playbook, args.files_folder)
    else:
        print("Invalid direction specified. Use 'save' or 'load'.")

if __name__ == '__main__':
    main()
