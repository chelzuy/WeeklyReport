import pandas as pd
import xlsxwriter

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
        elif "threshold violated" in value:
            value_split = value.split(".")
            new_column_description.append(value_split[0])
        elif "is down" in value:
            new_column_description.append(value)
        elif "is shutdown" in value:
            new_column_description.append(value)
        elif "Device not responding: Probably down or busy" in value:
            new_column_description.append(value)
        # elif "Discards rate" or "Errors rate" or "Memory Usage" or "Disk Utilization" in value:
        #     # DO NOT DELETE THE SPACE BEFORE ' is'
        #     value_split = value.split(" is")
        #     new_column_description.append(value_split[0])
        elif "Discards rate" in value:
            # DO NOT DELETE THE SPACE BEFORE ' is'
            value_split = value.split(" is")
            new_column_description.append(value_split[0])
        elif "Errors rate" in value:
            # DO NOT DELETE THE SPACE BEFORE ' is'
            value_split = value.split(" is")
            new_column_description.append(value_split[0])
        elif "Memory Usage" in value:
            # DO NOT DELETE THE SPACE BEFORE ' is'
            value_split = value.split(" is")
            new_column_description.append(value_split[0])
        elif "Disk Utilization" in value:
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
    data_df2[column] = data_df2[column].replace(
        r"(\W*(_x000D_)\W*)", '', regex=True)

#Count Occurences of Alarm Message
occurence = data_df2.groupby(
    ["Device", "Category", "Alarm Message"]).size().reset_index(name="Occurences")
data_column = occurence.sort_values(
    "Alarm Message", inplace=False, ascending=True)

#=================== CELL BORDER and ALIGN CENTER ===================#

writer = pd.ExcelWriter("new_file.xlsx", engine='xlsxwriter')
data_column.to_excel(writer, sheet_name="Sheet1", index=False)
workbook = writer.book
worksheet = writer.sheets["Sheet1"]

border_format = workbook.add_format(
    {"bottom": 1,
     "top": 1,
     "left": 1,
     "right": 1,
     "align": "vcenter"}
)
# workbook.add_format({'align': 'vcenter'})
worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(data_column), (len(data_column.columns)-1)),
                             {'type': 'no_errors', 'format': border_format})

data_format = workbook.add_format()
data_format.set_align('center')

worksheet.set_column("A:D", 18, data_format)
writer.save()
