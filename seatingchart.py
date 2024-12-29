import pandas as pd
import numpy as np

# Load the Excel file
df = pd.read_excel('chart.xlsx')

class Cadet:
    def __init__(self, name, let, block, date_guest, family, friends, beef):
        self.name = name
        self.let = let
        self.block = block
        self.date_guest = date_guest
        self.family = family
        self.friends = friends
        self.beef = beef if isinstance(beef, str) else ""
        self.seatnum = 0

def algorithm():
    seat = 1
    global table_assignments
    table_assignments = {i: [] for i in range(1, 37)}  # Dictionary to track cadets at each table
    let3a_assigned = {i: False for i in range(1, 37)}  # Track if a LET 3A cadet is assigned to each table
    cadet_list = list(cadet_dict.keys())
    for cadet in cadet_list:
        for table in range(1, 37):
            if cadet_dict[cadet].seatnum == 0:
                # Check if the table already has 8 cadets
                if len(table_assignments[table]) < 8:
                    if str(cadet_dict[cadet].friends) != "nan":
                            expansion = any(friend in table_assignments[table] for friend in cadet_dict[cadet].friends.split(';'))
                    else:
                        expansion = False
                    if True:
                                            # Check if the table needs a LET 3A cadet
                        if cadet_dict[cadet].let == '3A' and not let3a_assigned[table]:
                            # Ensure adding the cadet and their guests does not exceed the table limit
                            if str(cadet_dict[cadet].date_guest) != "nan":
                                guests = cadet_dict[cadet].date_guest.split(';') if ";" in cadet_dict[cadet].date_guest else [cadet_dict[cadet].date_guest]
                            else:
                                guests = []
                            if len(table_assignments[table]) + 1 + len(guests) <= 8:
                                cadet_dict[cadet].seatnum = seat
                                table_assignments[table].append(cadet_dict[cadet].name)
                                seat += 1
                                for guest in guests:
                                    if str(guest) != "nan":
                                        if guest not in cadet_dict:
                                            cadet_dict[guest] = Cadet(guest, "", "", "", "", "", "")
                                        cadet_dict[guest].seatnum = seat
                                        table_assignments[table].append(guest)
                                        seat += 1
                                break
                        # Try to seat cadets from the same block together
                        # Try to seat cadets from the same block or with friends together
                        
                        elif any(cadet_dict[assigned_cadet].block == cadet_dict[cadet].block for assigned_cadet in table_assignments[table]) or \
                            expansion == True or len(table_assignments[table]) == 0:
                            # Ensure adding the cadet and their guests does not exceed the table limit
                            if str(cadet_dict[cadet].date_guest) != "nan":
                                guests = cadet_dict[cadet].date_guest.split(';') if ";" in cadet_dict[cadet].date_guest else [cadet_dict[cadet].date_guest]
                            else:
                                guests = []
                            if len(table_assignments[table]) + 1 + len(guests) <= 8:
                                cadet_dict[cadet].seatnum = seat
                                table_assignments[table].append(cadet)
                                seat += 1
                                # Ensure the guest sits next to the cadet
                                #print(cadet_dict[cadet].date_guest)
                                for guest in guests:
                                    if str(guest) != "nan":
                                        if guest in cadet_dict:
                                            cadet_dict[guest].seatnum = seat
                                            table_assignments[table].append(guest)
                                        else:
                                            table_assignments[table].append(guest)
                                        seat += 1
                                break
                            
def beef_check():
    for table in range(1, 37):
        for cadet in table_assignments[table][:]:
            if cadet in cadet_dict:
                if cadet_dict[cadet].beef != "":
                    for beefcadet in table_assignments[table][:]:
                        if beefcadet in cadet_dict:
                            if cadet_dict[beefcadet].name in cadet_dict[cadet].beef:
                                print(f"Found beef with {cadet_dict[beefcadet].name}")
                                print(f"Moving {cadet_dict[cadet].name} and their guests to a new table")
                                for new_table in range(1, 37):
                                    if new_table != table:
                                        if str(cadet_dict[cadet].date_guest) != "nan":
                                            guests = cadet_dict[cadet].date_guest.split(';') if ";" in cadet_dict[cadet].date_guest else [cadet_dict[cadet].date_guest]
                                        else:
                                            guests = []
                                        if len(table_assignments[new_table]) + 1 + len(guests) < 8:
                                            if not any(cadet_dict[assigned_cadet].name in cadet_dict[cadet].beef.split(',') for assigned_cadet in table_assignments[new_table]):
                                                table_assignments[table].remove(cadet)
                                                table_assignments[new_table].append(cadet)
                                                for guest in guests:
                                                    if guest in table_assignments[table]:
                                                        table_assignments[table].remove(guest)
                                                        table_assignments[new_table].append(guest)
                                                break

def create_dictionary():
    cadet_dict = {}
    for i in range(2, len(df)):
        cadet_dict[df.iloc[i, 0]] = Cadet(df.iloc[i, 0], df.iloc[i, 1], df.iloc[i, 2], df.iloc[i, 3], df.iloc[i, 4], df.iloc[i, 5], df.iloc[i, 6])
    return cadet_dict
def seatfamily():
    for table in range(1, 37):
        if len(table_assignments[table]) == 0:
            for cadet in cadet_dict:
                if len(table_assignments[table]) >= 8:
                    continue
                if str(cadet_dict[cadet].family) != "nan":
                    family = cadet_dict[cadet].family.split(';') if ";" in cadet_dict[cadet].family else [cadet_dict[cadet].family]
                    if type(family) == list:
                        for member in family:
                            if member not in table_assignments:
                                table_assignments[table].append(member)
                            
                    else:
                        if family not in table_assignments:
                            table_assignments[table].append(family)
                    cadet_dict[cadet].family = np.nan

                    
def print_table_assignments():
    for table, cadets in table_assignments.items():
        print(f"Table {table}:")
        for cadet in cadets:
            if cadet in cadet_dict:
                print(f"  {cadet_dict[cadet].name}")
            else:
                print(f"  {cadet}")
        print()
def exceptionchecking():
    for table in table_assignments:
        if len(table_assignments[table]) > 8:
            print(f"Table {table} has too many cadets seated")
import pandas as pd

def output_seating_chart():
    # Create a dictionary to store the seating chart
    data = {}
    for table in range(1, 37):
        table_data = [f'Table {table}']  # Add the table name as the first entry
        for cadet in table_assignments[table]:
            table_data.append(cadet)
        data[f'Table {table}'] = table_data

    # Create a list of DataFrames for each row of 6 tables
    dfs = []
    for i in range(0, 36, 6):
        row_data = {f'Table {j+1}': data[f'Table {j+1}'] for j in range(i, i+6)}
        dfs.append(pd.DataFrame(dict([(k, pd.Series(v)) for k, v in row_data.items()])))

    # Concatenate the DataFrames horizontally
    final_df = pd.concat(dfs, axis=1)

    # Write the DataFrame to an Excel file with formatting
    with pd.ExcelWriter('tables.xlsx', engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False, header=False, sheet_name='Seating Chart')
        workbook = writer.book
        worksheet = writer.sheets['Seating Chart']

        # Set the column width and enable text wrapping
        for col_num, col in enumerate(final_df.columns):
            max_length = max(final_df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(col_num, col_num, max_length, workbook.add_format({'text_wrap': True}))



# Call the function to output the seating chart

cadet_dict = create_dictionary()
algorithm()
for i in range(5):
    beef_check()
seatfamily()
#print_table_assignments()
exceptionchecking()
output_seating_chart()