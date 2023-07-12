import pandas as pd
import random

# Read the Excel file into a DataFrame
data = pd.read_excel("./ep.xlsx")

# Group the employees by department and count the number of employees in each department
nongNong_counts = data["Department"].value_counts()

# Display the statistics
print("Department\tNongNong count")
print(nongNong_counts)

# Define the number of groups and the desired group size
num_groups = 6
group_size = len(data) // num_groups

# Initialize empty groups
groups = [[] for _ in range(num_groups)]

# Distribute employees equally across the groups
for department, count in nongNong_counts.items():
    nongNongs = data[data["Department"] == department].index.tolist()
    random.shuffle(nongNongs)

    for i, nongNong in enumerate(nongNongs):
        group_index = i % num_groups
        groups[group_index].append(nongNong)

# Allocate any remaining employees randomly to groups
remaining_nongNongs = data[
    ~data.index.isin(
        [assigned_nongNong for group in groups for assigned_nongNong in group]
    )
]

random.shuffle(remaining_nongNongs.index)

for i, nongNong in enumerate(remaining_nongNongs.index):
    group_index = i % num_groups
    groups[group_index].append(nongNong)

# Display personnel statistics for each group
for group_num, group in enumerate(groups):
    print(f"\nGroup {group_num} Personnel Statistics:")
    group_data = data.loc[group]
    group_nongNong_counts = group_data["Department"].value_counts()
    total_nongNongs = len(group_data)
    print(group_nongNong_counts)
    print(f"Total NongNongs: {total_nongNongs}")

    # Export the list of employees in each group to an Excel file
    group_filename = f"Group_{group_num}_ragNong.xlsx"
    group_data.to_excel(group_filename, index=False)
    print(f"Exported group {group_num} NongNongs to {group_filename}")