import pandas as pd
import random, os

OUTPUT_DIR = 'output/'
FILE_NAME = 'FILE.xlsx'
FAC_COL_NAME = 'faculty'
NUM_GROUPS = 28
FACTOR = 5
random.seed(42)

# Read the Excel file into a DataFrame
data = pd.read_excel(FILE_NAME)

# Group the students by department and count the number of students in each department
nongNong_counts = data[FAC_COL_NAME].value_counts()

# Display the statistics
print("Department\tNongNong count")
print(nongNong_counts)
print("\t --------------- \t")

# Initialize empty groups
groups = [[] for _ in range(NUM_GROUPS)]

# Separate departments with fewer students than the desired minimum group size
min_group_size = NUM_GROUPS + FACTOR  # Set the desired minimum group size
small_departments = nongNong_counts[
    (nongNong_counts <= min_group_size) & (nongNong_counts > 0)
].index.tolist()

random.shuffle(small_departments)

# Distribute students from small departments to individual groups
for department in small_departments:
    department_students = data[data[FAC_COL_NAME] == department].index.tolist()

    # Randomly choose a group until we find one that is not full
    while True:
        group_index = random.randint(0, NUM_GROUPS-1)
        if len(groups[group_index]) < min_group_size:
            break

    groups[group_index].extend(department_students)

    # Remove the small department from the nongNong_counts
    nongNong_counts.drop(department, inplace=True)


# Calculate the desired group size after accounting for small departments
group_size = (len(data) - len(small_departments)) // NUM_GROUPS

# Distribute the remaining students across the groups
for department, count in nongNong_counts.items():
    students = data[data[FAC_COL_NAME] == department].index.tolist()
    random.shuffle(students)

    # Distribute students to each group
    for i, student in enumerate(students):
        group_index = i % NUM_GROUPS
        groups[group_index].append(student)

# Calculate the target size for each group
target_size = (len(data) - len(small_departments)) // NUM_GROUPS

# Balance the group sizes
for group_num, group in enumerate(groups):
    while len(group) > target_size:
        # Find a random student in the group
        student_to_remove = random.choice(group)

        # Remove the student from the group
        group.remove(student_to_remove)

        # Find another group to transfer the student
        transfer_group_num = (group_num + 1) % NUM_GROUPS
        groups[transfer_group_num].append(student_to_remove)

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
