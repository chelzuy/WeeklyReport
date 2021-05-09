import pandas as pd
from pandas.core import groupby


#=================== ORIGINAL EXCEL MANIPULATION ===================#

#Reading excel file
data_alarm = pd.read_excel("Alarm.xlsx", sheet_name="Alarm", engine="openpyxl")
data_df = pd.DataFrame(data_alarm)
data_df2 = data_df.dropna(subset=["Device", "Category", "Description"])


new_column_description = []
column_description = list(data_df["Description"])
for value in column_description:
    try:
        if "CPU Utilization(UCD SNMP MIB)" in value:
            new_column_description.append("CPU Utilization(UCD SNMP MIB)")
        elif "CPU Utilization" in value:
            new_column_description.append("CPU Utilization")
        elif "Memory Utilization" in value:
            new_column_description.append("Memory Utilization")
        elif "Device not responding: Probably down or busy" in value:
            new_column_description.append(
                "Device not responding: Probably down or busy")
        elif "threshold violated" in value:
            value_split = value.split(".")
            new_column_description.append(value_split[0])
        elif "is down" in value:
            new_column_description.append(value)
        elif "is shutdown" in value:
            new_column_description.append(value)
        elif "Discards rate" or "Errors rate" or "Memory Usage" or "Disk Utilization" in value:
            # DO NOT DELETE THE SPACE BEFORE ' is'
            value_split = value.split(" is")
            new_column_description.append(value_split[0])
        else:
            new_column_description.append(value)
    except TypeError:
        pass

#=================== NEW EXCEL MANIPULATION ===================#

#Drop Description Column
data_df2.drop("Description", axis=1, inplace=True)

#Inserting Alarm Message Column
data_df2.insert(2, "Alarm Message", new_column_description)

for column in data_df2:       
    data_df2[column] = data_df2[column].replace(r"(\W*(_x000D_)\W*)", '', regex=True)

#Count Occurences of Alarm Message
occurence = data_df2.groupby(["Device", "Category", "Alarm Message"]).size().reset_index(name="Occurences")
data_column = occurence.sort_values("Alarm Message")

new_file = data_column.to_excel("new_file.xlsx", index=False)


