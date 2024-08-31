import pandas as pd
import numpy as np 


# reading data from excel file
all_data_df=pd.read_excel('input.xlsx')

# print (all_data_df.iloc[0:5,:])

# Remove rows without blockname where this column is NaN
data_df = all_data_df.dropna(subset=[all_data_df.columns[3]])

# Replace NaN with 0 in the entire DataFrame
data_df_filled = data_df.fillna(0)


# Convert the filled DataFrame to a NumPy array
data = data_df_filled.to_numpy()
print (data[0:5,:])

# making lists of drills and blocknames
drill_list=[]
block_list=[]
for item in data:
    if item[1] not in drill_list:
        drill_list.append(item[1])
    if item[3] not in block_list:
        block_list.append(item[3])

# extracting results from data array
drill_block_info=[]
for drill in drill_list:
    for block in block_list:
        days=[]
        fe=0
        fe_waste=0
        waste=0
        for item in data:
            if item[1] == drill and item[3] == block:
                sum=0
                days.append(item[2])
                fe+=item[4]
                fe_waste+=item[5]
                waste+=item[6]
                contractor=item[0]
                sum+=(fe+fe_waste+waste)
        drill_block_info.append([contractor , drill, block, days, fe, fe_waste, waste, sum])



# deleting items of results whithout working days (blocks that drill did not operat on them)
drill_block_info = [item for item in drill_block_info if item[3] != []]
# print (drill_block_info[0:4])

# Convert the results to a DataFrame to write in excel file
data_df=pd.DataFrame(data, columns=['Contract', 'Drill', 'Working days', 'Block', 'Fe', 'Fe-W', 'W', 'Sum','تطابق','وضعیت حفاری','تایید کانی کاوان'])
result_df = pd.DataFrame(drill_block_info, columns=['Contract', 'Drill', 'Block', 'Working days', 'Fe', 'Fe-W', 'W', 'Sum'])


# Write this results to Excel file with  new sheets 
with pd.ExcelWriter('input.xlsx', engine='openpyxl', mode='a') as writer:  # mode='a' to append to existing file
    # Check if 'result' sheet exists and remove it
    if 'result' in writer.book.sheetnames:
        del writer.book['result']
    if 'data' in writer.book.sheetnames:
        del writer.book['data']
    result_df.to_excel(writer, sheet_name='result', index=False)
    data_df.to_excel(writer, sheet_name='data', index=False)

