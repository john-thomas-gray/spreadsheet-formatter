import os
import glob
import pandas as pd

def process_xls_to_csv(input_file_path, output_file_path):
    df = pd.read_excel(input_file_path, sheet_name=0)

    previous_name = None
    subtitle_to_use = 'Subs M1'

    def update_subtitles(row):
        nonlocal previous_name, subtitle_to_use
        current_name = row.iloc[2].split('says in Spanish')[0].strip().split()[-1]

        if previous_name is None or current_name != previous_name:
            subtitle_to_use = 'Subs M2' if subtitle_to_use == 'Subs M1' else 'Subs M1'

        row.iloc[1] = subtitle_to_use
        previous_name = current_name
        return row

    df = df[df.iloc[:, 2].str.contains('says in Spanish', na=False)]

    def modify_text(text):
        part_after_says_in_spanish = text.split('says in Spanish', 1)[1].strip() if 'says in Spanish' in text else text
        if 'says' in part_after_says_in_spanish:
            says_index = part_after_says_in_spanish.index('says')
            last_period_before_says = part_after_says_in_spanish.rfind('.', 0, says_index)
            return part_after_says_in_spanish[:says_index].strip() if last_period_before_says == -1 else part_after_says_in_spanish[:last_period_before_says + 1].strip()
        return part_after_says_in_spanish

    df = df.apply(update_subtitles, axis=1)
    df.iloc[:, 2] = df.iloc[:, 2].apply(modify_text)

    df.to_csv(output_file_path, index=False)

def process_all_xls(csv_input_dir, xls_output_dir):
    if not os.path.exists(xls_output_dir):
        os.makedirs(xls_output_dir)

    xls_files = glob.glob(os.path.join(csv_input_dir, '*.xls'))

    for xls_file in xls_files:
        base_filename = os.path.basename(xls_file).replace('.xls', '.csv')
        output_file_path = os.path.join(xls_output_dir, base_filename)

        process_xls_to_csv(xls_file, output_file_path)
        print(f"Processed {xls_file} to {output_file_path}")

# Example usage: process all XLS files
csv_input_dir = 'xls_in'
xls_output_dir = 'csv_out'
process_all_xls(csv_input_dir, xls_output_dir)
