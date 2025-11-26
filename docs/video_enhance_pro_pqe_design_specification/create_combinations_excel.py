
import pandas as pd
import xlsxwriter

def create_interactive_excel_file():
    """
    Generates a colored Excel file with checkboxes for each combination.
    """
    resolutions = [
        "1920x1080", "2400x1200", "2560x1440", "2880x1620",
        "3036x1708", "3840x2160", "5120x1600", "6240x1172",
        "2560x1600", "2880x1800", "2240x1260"
    ]
    frame_rates = ["60Hz", "120Hz", "144Hz"]
    lane_counts = [2, 4, 8]
    color_depths = ["8 bit", "10 bit"]

    data = []
    index = 1
    for res in resolutions:
        for fr in frame_rates:
            for lanes in lane_counts:
                for depth in color_depths:
                    data.append({
                        "索引": index,
                        "分辨率": res,
                        "帧率": fr,
                        "LANE数": f"{lanes} LANE",
                        "色深": depth
                    })
                    index += 1

    df = pd.DataFrame(data)
    
    # Add a placeholder for the checkbox column header
    df["完成"] = ""

    output_filename = "video_combinations_interactive.xlsx"

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name='Combinations', index=False)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets['Combinations']

    # Add a border format.                                            
    border_format = workbook.add_format({'border': 1})                    
                                                                           
    # Apply a conditional format to add a border to all non-blank cells.  
    worksheet.conditional_format(0, 0, len(df), len(df.columns) - 1,      
                                 {'type': 'no_blanks',                    
                                 'format': border_format}) 

    # --- Start Formatting and Checkbox Insertion ---

    # 1. Define color formats
    color_map = {
        "1920x1080": "#DDEBF7", "2400x1200": "#E2F0D9", "2560x1440": "#FFF2CC",
        "2880x1620": "#F8CBAD", "3036x1708": "#D9E1F2", "3840x2160": "#FCE4D6",
        "5120x1600": "#EDEDED", "6240x1172": "#DEEBF7",
        "2560x1600": "#D0E0E3", "2880x1800": "#EAD8D7", "2240x1260": "#D4EFDF",
    }
    color_formats = {res: workbook.add_format({'bg_color': color}) for res, color in color_map.items()}

    # 2. Get column index for the checkbox column
    checkbox_col_idx = df.columns.get_loc("完成")

    # 3. Iterate over the DataFrame to apply formatting and insert checkboxes
    # Note: row_num is 0-indexed, Excel rows are 1-indexed.
    for row_num, row_data in df.iterrows():
        excel_row = row_num + 1
        
        # Apply row color
        resolution = row_data["分辨率"]
        res_format = color_formats.get(resolution)
        if res_format:
            worksheet.set_row(excel_row, None, res_format)

        # Insert a checkbox in the '完成' column
        worksheet.insert_checkbox(excel_row, checkbox_col_idx, {'checked': False})

    # 4. Adjust column widths for better readability
    worksheet.set_column('A:A', 5)   # 索引
    worksheet.set_column('B:B', 18)  # 分辨率
    worksheet.set_column('C:C', 8)   # 帧率
    worksheet.set_column('D:D', 10)  # LANE数
    worksheet.set_column('E:E', 8)   # 色深
    worksheet.set_column('F:F', 8)   # 完成 (Checkbox)

    # Close the Pandas Excel writer and output the Excel file.
    writer.close()

    print(f"Successfully created '{output_filename}' with interactive checkboxes.")

if __name__ == "__main__":
    create_interactive_excel_file()
