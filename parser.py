import pandas as pd

# Function to read and parse the .ods file
def parse_ods(file_path):
    # Read the .ods file into a pandas DataFrame
    df = pd.read_excel(file_path, engine="odf")
    filtered_df = df.iloc[:, [2, 3]]
    #filtered_df.columns = ['Application_Number', 'Result']
    filtered_df = filtered_df.dropna()
    filtered_df = filtered_df.rename(columns=filtered_df.iloc[0]).drop(filtered_df.index[0])
    filtered_df.reset_index(drop=True, inplace=True)
    print(filtered_df)
    return filtered_df

# Path to the .ods file
file_path = '20240714_NDVO_Visa_Decisions.ods'

# Parse the .ods file
df = parse_ods(file_path)
application_number_to_find = 68049422
# Check if the Application Number exists in the DataFrame
if application_number_to_find in df['Application Number'].values:
    decision = df.loc[df['Application Number'] == application_number_to_find, 'Decision'].iloc[0]
    print(f"Decision for Application Number {application_number_to_find}: {decision}")
else:
    print(f"Application Number {application_number_to_find} not found.")
