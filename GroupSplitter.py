import pandas as pd
import random, os

OUTPUT_DIR = 'output/'
FILE_NAME = 'MOCK_DATA.xlsx'
FAC_COL_NAME = 'Faculty'
NUM_GROUPS = 2

random.seed(42)

# Read the Excel file into a DataFrame
data = pd.read_excel(FILE_NAME)

# Group the employees by department and count the number of employees in each department
nongNong_fac_count = data[FAC_COL_NAME].value_counts()
nongNong_total_count = data.shape[0]

# Display the statistics
# print("Faculty\tNongNong count")

group_size = len(data) // NUM_GROUPS
print(f"Count nongNong = {nongNong_total_count}")
print(f"Group needed = {NUM_GROUPS} | Group size = {group_size}" + "\n---" )

# Initialize empty groups
groups = [[] for _ in range(NUM_GROUPS)]

# Distribute employees equally across the groups
for department, count in nongNong_fac_count.items():
    nongNongs = data[data[FAC_COL_NAME] == department].index.tolist()
    random.shuffle(nongNongs)

    for i, nongNong in enumerate(nongNongs):
        group_index = i % NUM_GROUPS
        groups[group_index].append(nongNong)

# Allocate any remaining employees randomly to groups
remaining_nongNongs = data[
    ~data.index.isin(
        [assigned_nongNong for group in groups for assigned_nongNong in group]
    )
]

random.shuffle(remaining_nongNongs.index)

for i, nongNong in enumerate(remaining_nongNongs.index, start=1):
    group_index = i % NUM_GROUPS
    groups[group_index].append(nongNong)

# create directory for output
os.makedirs('output', exist_ok=True)

# Display personnel statistics for each group
for group_num, group in enumerate(groups, start=1):
    print(f"\nGroup {group_num} Personnel Statistics:")
    group_data = data.loc[group]
    group_nongNong_counts = group_data[FAC_COL_NAME].value_counts()
    total_nongNongs = len(group_data)
    print(group_nongNong_counts)
    print(f"Group No. {group_num} " + f"Total NongNongs: {total_nongNongs}")

    # Export the list of employees in each group to an Excel file, and save in the location
    group_filename = OUTPUT_DIR + f"Group_{group_num}_raknong66.xlsx"
    group_data.to_excel(group_filename, index=False)
    print(f"Exported group {group_num} NongNongs to {group_filename}")
