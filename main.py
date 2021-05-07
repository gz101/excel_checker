# Import packages.
import pandas as pd 


# Define constants - DO NOT CHANGE ANYTHING BEYOND THIS SECTION.
OLD_FILE = "sample-address-1.xlsx"
NEW_FILE = "sample-address-2.xlsx"
KEY = "key_name"
FIELD_1 = "field_1"
FIELD_2 = "field_2"
FIELD_3 = "field_3"
FIELD_4 = "field_4"


""" ...................................................................
WARNING: DO NOT CHANGE ANYTHING BEYOND THIS LINE
....................................................................""" 
# Define useful functions.
def report_diff(x):
    """
    report_diff x -- Defines a function to show the changes within each
    field within a single data point in a df.
    """
    return x[0] if x[0] == x[1] else '{} ---> {}'.format(*x)


# Body goes here.
# Load data and create columns to track.
old = pd.read_excel(OLD_FILE, "Sheet1", na_values=["NA"])
new = pd.read_excel(NEW_FILE, "Sheet1", na_values=["NA"])
old["version"] = "old"
new["version"] = "new"

"""
What is the key field (column) in the dataset? We use this to determine
what are the new entries and which entries have been newly added.
"""
old_key_all = set(old[KEY])
new_key_all = set(old[KEY])
removed_rows = old_key_all - new_key_all 
added_rows = new_key_all - old_key_all

"""
Combine the two sets of data and drop the duplicates. The data in eaach row
is now unique. All column names (fields) within subset are used for 
comparison.
"""
all_data = pd.concat([old, new], ignore_index=True)
changes = all_data.drop_duplicates(subset=[KEY,
                                           FIELD_1,
                                           FIELD_2,
                                           FIELD_3,
                                           FIELD_4], keep="last")

# Figure out the duplicated rows (by key_name).
dupe_entries = changes[changes[KEY].duplicated() == \
               True][KEY].tolist()
dupes = changes[changes[KEY].isin(dupe_entries)]

# Split the old and new data into separate dataframes.
change_new = dupes[(dupes["version"] == "new")]
change_old = dupes[(dupes["version"] == "old")]

# Drop the temp columns - no longer needed.
change_new = change_new.drop(["version"], axis=1)
change_old = change_old.drop(["version"], axis=1)

# Index on the key_name field.
change_new.set_index(KEY, inplace=True)
change_old.set_index(KEY, inplace=True)

# Combine all the changes together.
df_all_changes = pd.concat([change_old, change_new],
                            axis="columns",
                            keys=["old", "new"],
                            join="outer")

# Moves old and new columns next to each other.
df_all_changes = df_all_changes.swaplevel(axis="columns") \
                 [change_new.columns[0:]]

"""
Combines the different columns using the report_diff function.
If different, both values are captured within a single cell.
The index is reset.
"""
df_changed = df_all_changes.groupby(level=0, axis=1).apply \
             (lambda frame: frame.apply(report_diff, axis=1))
df_changed = df_changed.reset_index()

# Find out what has been removed and what has been added.
df_removed = changes[changes[KEY].isin(removed_rows)]
df_added = changes[changes[KEY].isin(added_rows)]

# Output results into Excel file. Template must exist.
output_columns = [KEY, FIELD_1, FIELD_2, FIELD_3, FIELD_4]
writer = pd.ExcelWriter("my-diff.xlsx")
df_changed.to_excel(writer,"changed", index=False, columns=output_columns)
df_removed.to_excel(writer,"removed",index=False, columns=output_columns)
df_added.to_excel(writer,"added",index=False, columns=output_columns)
writer.save()