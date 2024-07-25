import pandas as pd
import os
import glob

def process_csv_to_xls(input_file_path, output_file_path):
    input_df = pd.read_csv(input_file_path)

    output_df = pd.DataFrame(
        index=range(len(input_df) + 2),
        columns=['TIME IN', 'CHARACTER', 'SCRIPT']
    )

    output_df.at[0, 'SCRIPT'] = 'BEGINNING OF SCRIPT'
    output_df.at[len(output_df) - 1, 'SCRIPT'] = 'END OF SCRIPT'

    if 'Start TC' in input_df.columns:
        output_df['TIME IN'][1:len(input_df['Start TC']) + 1] = input_df['Start TC']
    else:
        raise ValueError("'Start TC' column is not found in the input file.")

    if 'Character' in input_df.columns:
        output_df['CHARACTER'][1:len(input_df['Character']) + 1] = input_df['Character']
    else:
        raise ValueError("'Character' column is not found in the input file.")

    if 'Text' in input_df.columns:
        output_df['SCRIPT'][1:len(input_df['Text']) + 1] = input_df['Text']
    else:
        raise ValueError("'Text' column is not found in the input file.")

    output_df.to_excel(output_file_path, index=False)

def process_all_csvs(csv_input_dir, xls_output_dir):
    if not os.path.exists(xls_output_dir):
        os.makedirs(xls_output_dir)

    csv_files = glob.glob(os.path.join(csv_input_dir, '*.csv'))

    for csv_file in csv_files:

        base_filename = os.path.basename(csv_file).replace('.csv', '_formatted.xlsx')
        output_file_path = os.path.join(xls_output_dir, base_filename)

        process_csv_to_xls(csv_file, output_file_path)
        print(f"Processed {csv_file} to {output_file_path}")

csv_input_dir = 'csv_in'
xls_output_dir = 'xls_out'

process_all_csvs(csv_input_dir, xls_output_dir)
