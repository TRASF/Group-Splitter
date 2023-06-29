import pandas as panda
import random

def read_excel_files(fileName):

    # Load the information from the desired Excel file
    nongNongsList = panda.read_excel(fileName, index_col = 0)

    # Determine the total number of nongNong in each Department
    departmentCounter = nongNongsList['Department'].value_counts()

    print("Department Statistic Counter\n")
    # Prevent output of the line ' Name: count, dtype: int64 '
    print(departmentCounter.to_string( index = True ))
    print(f"Total:\t\t  ",len(nongNongsList))

    return [nongNongsList, departmentCounter]

def group_splitter(numberOfGroup, nongNongs):

    # split list data from read_excel_files 
    totalNongNong, departmentCounter = nongNongs

    # Determine how many nongNongs are in each group.
    groupSize = len(totalNongNong) // numberOfGroup
    print(f"Nong per group:\t   {groupSize}\n")

    groups = [ [] for _ in range(numberOfGroup) ]
    logCounter = 0          # For log

    for department, nong in departmentCounter.items():
        
        # Create a list that classifies nongNongs according to their respective departments.
        nongList = totalNongNong[ totalNongNong[ 'Department' ] == department ].index.tolist()

        # Shuffle the order indicates that each department has an equal chance of being chosen first.
        random.shuffle(nongList)
        
        # Dividing up the nong into a group in a fair way.
        for index, nong in enumerate(nongList):
            indicator = index % numberOfGroup
            groups[indicator].append(nong)
            logCounter += 1

    # For counting if someone leftover
    print("---------- Loop Log ----------")
    print(f"Log counter:\t   {logCounter}")
    print("------------------------------")
    return groups


def print_detailed_group_list(nongNongs, groups):

    totalNongNong, departmentCount = nongNongs

    for groupIndicator, group in enumerate(groups):

        print(f"\nGroup {groupIndicator + 1} Personnel Statistics:")

        groupData = totalNongNong.loc[group]
        nongNongsInGroupCount = groupData['Department'].value_counts()
        totalInGroup = len(groupData)
        
        print(nongNongsInGroupCount.to_string( index = True ))
        print(f"Total Nongs in group:\t {totalInGroup}")

def main():

    fileName = 'ep.xlsx'
    numberOfGroup = 6

    nongNongs = read_excel_files(fileName)
    groups = group_splitter( numberOfGroup, nongNongs )

    print_detailed_group_list(nongNongs, groups)


if __name__ == '__main__':
    main()

