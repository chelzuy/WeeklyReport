import pandas as pd



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
        elif "Discards rate" or "Errors rate" or "Memory Usage" in value:
            # DO NOT DELETE THE SPACE BEFORE ' is'
            value_split = value.split(" is")
            new_column_description.append(value_split[0])
        # elif "Errors rate" in value:
        #         value_split = value.split(" is") # DO NOT DELETE THE SPACE BEFORE ' is'
        #         new_column_description.append(value_split[0])
        # elif "Memory Usage" in value:
        #         value_split = value.split(" is") # DO NOT DELETE THE SPACE BEFORE ' is'
        #         new_column_description.append(value_split[0])
        else:
            new_column_description.append(value.upper())
    except TypeError:
        pass

#=================== NEW EXCEL MANIPULATION ===================#

#Drop Description Column
data_df2.drop("Description", axis=1, inplace=True)

#Inserting Alarm Message Column
data_df2.insert(1, "Alarm Message", new_column_description)
data_df3 = data_df2.sort_values(["Alarm Message"],ascending=False)
for column in data_df3:       
    data_df3[column] = data_df3[column].replace(r"(\W*(_x000D_)\W*)", '', regex=True)

#Count Occurences of Alarm Message
occurence = data_df3.groupby(["Device", "Category", "Alarm Message"]).size().reset_index(name="Occurences")

data_column = occurence.sort_values("Device")

# #Replace _x000D_ to ""
# for column in data_column:       
#     data_column[column] = data_column[column].replace(r"(\W*(_x000D_)\W*)", '', regex=True)

#Save count into a new_file
new_file = data_column.to_excel("new_file.xlsx", index=False)


