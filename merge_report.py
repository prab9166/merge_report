import traceback

import pandas as pd
import numpy as np
import xlsxwriter
import time

try:
    client_file = pd.read_excel(r"D:\dataset\quality_master_file\client file.xlsx")  # main client data (final base file)
    merge_report = pd.read_csv(r"D:\dataset\quality_master_file\merged_report.csv")  # system merged file (has Token etc)
    internal_report = pd.read_csv(r"D:\dataset\quality_master_file\edge_report.csv",low_memory=False)  # internal lookup file (extra info)
    merge_output = r"D:\dataset\quality_master_file\final_output.xlsx"  # final output path
    prelim_merger_path = r"D:\dataset\quality_master_file\prelim_merge_output.xlsx"  # first merge output (temp)

except FileNotFoundError as fe:
    print('File not found error', fe)
except Exception as e:
    print('Un-allowed value, check the file path', e)

###***** proper indexing per user input in a merge file ***####

print('Starting with the merge file')
print('\nMerge file columns are :\n')
for i, j in enumerate(merge_report.columns, 1):
    print(f'{i}:{j}')  # show columns with numbers so user can pick

while True:
    try:  # user selects columns from above (can give 1,2 or 3-7 etc)
        merge_report_input = input('\nWhich column(s) do you want to include, from merged report - include all columns separated by comma(,) or can select a range or both')

        idx_list = []  # stores user input numbers (1-based)
        idx_list_fordf = []  # actual index used in df (0-based)
        error_found = False

        for a in merge_report_input.strip().split(','):  # split input
            if '-' in a:  # if range like 3-7, take all in between
                try:
                    a_int, b_int = [int(x) for x in a.strip().split('-')]
                    if int(a_int) > int(b_int) or int(a_int) > len(merge_report.columns) or int(b_int) > len(
                            merge_report.columns):
                        print('\nRange issue or out of column count, try again')  # basic validation
                        error_found = True
                        break

                    for i in range(int(a_int), int(b_int) + 1):
                        idx_list.append(i)  # keep original index
                        a_fordf = i - 1  # convert to 0-based
                        idx_list_fordf.append(a_fordf)

                except ValueError as be:
                    print('\nError in merge report range selection :', be)

            else:
                error_found = False
                a = int(a)
                idx_list.append(a)
                a_fordf = int(a) - 1
                idx_list_fordf.append(a_fordf)

        column = merge_report[merge_report.columns[idx_list_fordf]]
        print('You selected', column)

        if 'Token' not in column:
            print('Token column must be selected, check properly and try again')  # must for merge later
            error_found = True
            continue

        elif error_found:
            continue

        print('\nSelected index - ', list(set(sorted([i for i in idx_list]))))
        merge_report_column = merge_report.columns[idx_list_fordf]
        merged_df = merge_report[merge_report_column]  # filtered merge report
        print('\n')
        print(merged_df.head(5))
        break

    except ValueError as ve:
        print('Error with input in merge file - must be a no. and/or select a range using "-" symbol only : ')
        error_found = True
    except BaseException as be:
        print('An error occurred in merge file selection : ', be)

time.sleep(2)

###***** proper indexing per user input in an internal file ***####

print('\nNow starting with the internal file')
time.sleep(1)
print('\nInternal file columns are:')
for i, j in enumerate(internal_report.columns, 1):
    print(f'{i}:{j}')


while True:
    try:
        internal_selected_list = []
        internal_idx = []
        internal_idx_fordf = []

        internal_report_input = input(
            '\nWhich column would  you lookup from internal file? separated by comma "," ')  # columns to bring from internal file

        error_found1 = False

        for a in internal_report_input.strip().split(','):
            if '-' in a:

                a_idx, b_idx = (int(x) for x in a.strip().split('-'))
                if a_idx > b_idx or a_idx <= 0 or b_idx > len(internal_report.columns) :
                    print('\nFirst number should be smaller, or invalid range, TRY again')  # range check
                    error_found1 = True
                    break

                for i in range(int(a_idx), int(b_idx) + 1):
                    internal_idx.append(i)
                    a_fordf_internal = i - int(1)
                    internal_idx_fordf.append(int(a_fordf_internal))


            else:

                internal_idx.append(int(a))
                a_fordf_internal = int(a) - 1
                internal_idx_fordf.append(a_fordf_internal)

        if error_found1:
            continue

        print('\nIn internal file You selected column ', list(sorted(set(internal_idx))))
        internal_selected_list = internal_report.columns[internal_idx_fordf]

        print(internal_selected_list)
        internal_df = internal_report[internal_selected_list]  # selected internal data
        print('\n')
        print(internal_df.head(5))
        internal_report_list = []
        for i in internal_idx_fordf:
            idx = int(i)
            selected_internal_column = internal_report.columns[idx]
            print(f'\nLooking in internal file for index {idx + 1} : Found - {selected_internal_column}')
            internal_report_list.append(selected_internal_column)

        internal_df = pd.DataFrame(internal_report[internal_report_list])
        print(internal_df)

        # merge merge_report + internal file using Token

        prelim_merge_df = pd.merge(left=merged_df, right=internal_df, left_on='Token', right_on='Token')
        print(prelim_merge_df)

        time.sleep(1)
        # now moving to client file

        print('Now with the client file')

        time.sleep(1)

        print('Lets strat with client file')

        client_data = pd.DataFrame(client_file)
        print('Client sheet columns are:\n')

        break

    except ValueError as ve:
        print('Error in input with internal file - must be numbers or range "-" TRY again: ')
        error_found1 = True
        continue


    except BaseException as be:
        print('Error in internal selection or merge : ', be)
        error_found1 = True
        break


time.sleep(1)

for i, j in enumerate(client_data, 1):
    print(i, j)  # show client columns

while True:
    try:  # drop unwanted columns from client file
        client_data_input = input(
            'Which columns do you want to drop?if multiple, enter the index seperated by comas(",") or  type "0" if none')

        error_found2 = False
        i_list = []

        if client_data_input == '0':
            print('No column selected, will proceed with all columns')
            client_data_df = pd.DataFrame(client_data)
            final_data_df = client_data_df
            print('test- input 0', final_data_df.columns)
            break

        elif ',' in client_data_input:
            for i in client_data_input.strip().split(','):
                i = int(i)
                if i < 0 or i > len(client_data.columns):  # validation
                    print('Invalid input, out of range, TRY again')
                    error_found2 = True
                    break
                i_list.append(i - 1)
            print('You selected', client_data.columns[i_list])

        elif '-' in client_data_input:
            a1_int, b1_int = [int(x) for x in client_data_input.strip().split('-')]
            if b1_int < 0 or b1_int > len(client_data.columns):
                print('Invalid input, TRY again')
                error_found2 = True

            elif a1_int > b1_int:  # validation
                print('\nFirst number should be smaller, TRY again')
                error_found2 = True

            else:
                for i in range(int(a1_int), int(b1_int) + 1):
                    i_list.append(i - 1)
                print('You selected', client_data.columns[i_list])

        else:
            i = int(client_data_input)
            i_list.append(i - 1)

        if error_found2:
            continue

        print('final printing')
        final_data_df_columns = client_data.columns.drop(client_data.columns[i_list])  # dropping cols
        final_data_df = client_data[final_data_df_columns]
        print('test-', final_data_df.columns)
        break

    except ValueError as ve:
        print('An error in client file for: ', ve)
        print('Use numbers only, TRY again')
        error_found2 = True

    except BaseException as be:
        print('\nError in client file\n', be)

print('\nPrelim merge df columns\n', prelim_merge_df.columns)
with pd.ExcelWriter(prelim_merger_path, engine='xlsxwriter') as w:
    prelim_merge_df.to_excel(w)
    print('prelim merge df created')

print('Now we need to know which columns to match with')

while True:
    try:
        print(f'prelim columns + \n')
        for i, j in enumerate(prelim_merge_df.columns, 1):  # choose key from prelim df
            print(i, j)
        prelim_input = int(input('\nWhich column from prelim to pick up for matching\n'))

        error_found3 = False

        if prelim_input > len(prelim_merge_df.columns):
            print('Invalid column, TRY again')
            error_found3 = True
            continue

        else:
            prelim_input_column = prelim_merge_df.columns[prelim_input - 1]
            print('prelim input column : ', prelim_input_column)

        print(f'client sheet columns: \n')
        for i, j in enumerate(final_data_df.columns, 1):  # choose key from client file
            print(i, j)

        client_column_input = int(input('\nWhich column in client file to pick up for lookup\n'))

        if client_column_input > len(final_data_df.columns):
            print('Invalid column, TRY again')
            traceback.print_tb()
            error_found3 = True

        else:
            clientdf_input_column = final_data_df.columns[client_column_input - 1]
            print('client_df input column : ', clientdf_input_column)

        final_merge_df = pd.merge(left=prelim_merge_df, right=final_data_df, left_on=prelim_input_column,
                                  right_on=clientdf_input_column)  # final merge

        empty_col = []
        dupe_col = []
        for col in final_merge_df.columns:  # check empty + duplicate columns

            dupe_col_index = final_merge_df[col]

            if final_merge_df[col].isna().all():
                empty_col.append(col)
            if str(dupe_col_index).endswith('_y'):
                dupe_col.append(col)

        print('Empty columns in final merge file\n')
        print(empty_col)
        final_merge_df = final_merge_df[final_merge_df.columns.drop(empty_col)]  # drop empty
        print('empty columns removed\n')
        print('dupe columns are:\n')
        print(dupe_col)

        with pd.ExcelWriter(merge_output, engine='xlsxwriter') as w:
            final_merge_df.to_excel(w)

        print('Final merge excel created\n')

        break

    except ValueError as ve:
        print("Enter integer only, TRY again : ", ve)
        traceback.print_exc()
        error_found3 = True

    except BaseException as be:
        print("Error occurred : ", be)
        traceback.print_exc()
        error_found3 = True

print('final merge is :\n')
print(final_merge_df.head(5))

# final check again for safety

empty_col = []
for col in final_merge_df.columns:
    if final_merge_df[col].isna().all():
        empty_col.append(col)

if len(empty_col) == 0:
    print('No empty columns found')
else:
    print('empty column still exists - length of empty columns', len(empty_col))
