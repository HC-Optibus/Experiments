import os
import pandas as pd

def clean_timeline(input_folder):
    # Traverse the directory
    for root, dirs, files in os.walk(input_folder):
        # Check if the folder name contains "Caesar"
        if "Caesar" in os.path.basename(root):
            # Look for XLSX file with "Dataset" in its name
            dataset_files = [file for file in files if "Dataset" in file and file.endswith('.xlsx')]
            
            if dataset_files:
                dataset_file = os.path.join(root, dataset_files[0])
                output_folder = os.path.join(root, "Cleaned_TP")
                output_file = os.path.join(output_folder, "Cleaned_TP.xlsx")

                # Ensure the output folder exists, create if not
                os.makedirs(output_folder, exist_ok=True)

                # Read the data and perform cleaning (modify this part based on your cleaning logic)
                df = pd.read_excel(dataset_file, engine='openpyxl', sheet_name='Trips')
                df = df[df['Departure'] != df['Arrival']]
                df['Departure'] = pd.to_datetime(df['Departure'], format='%H:%M')
                
                # Sort the DataFrame based on 'Id' and 'Departure' in ascending order
                df = df.sort_values(by=['Id', 'Departure'])

                # Initialize 'Sub Trip Index' with 1 for all rows
                df['Sub Trip Index'] = 1

                # Function to update 'Sub Trip Index' based on the earliest departure time
                def update_sub_trip_index(group):
                    group['Sub Trip Index'] = range(1, len(group) + 1)
                    return group

                # Apply the function to each group of 'Id'
                df = df.groupby('Id').apply(update_sub_trip_index)

                # Sort the DataFrame back to the original order
                df = df.sort_index()

                # Identify unique 'Id' values
                unique_ids = df['Id'].value_counts() == 1

                # Void 'Sub Trip Index' for rows with unique 'Id'
                df.loc[df['Id'].isin(unique_ids[unique_ids].index), 'Sub Trip Index'] = None

                # Create a dictionary to map 'Service Groups' to numeric identifiers
                service_groups_mapping = {group: index + 1 for index, group in enumerate(df['Service Groups'].unique())}
                # Apply the mapping to create 'temp_Service Groups' column
                df['temp_Service Groups'] = df['Service Groups'].map(service_groups_mapping)

                # Create 'temp_Id' column with values before the first '_'
                df['temp_Id'] = df['Id'].str.split('_').str[0]

                # Convert 'temp_Id' to numeric for sorting
                df['temp_Id'] = pd.to_numeric(df['temp_Id'], errors='coerce')

                # Sort the DataFrame based on 'temp_Service Groups' and 'temp_Id' in ascending order
                df = df.sort_values(by=['temp_Service Groups', 'temp_Id'])

                # Drop the temporary columns
                df = df.drop(['temp_Id', 'temp_Service Groups'], axis=1)

                # Read the other sheets from the original Excel file
                other_sheets = pd.read_excel(dataset_file, engine='openpyxl', sheet_name=None)
                # The 'Trips' sheet will be replaced with the modified DataFrame
                other_sheets['Trips'] = df

                # Save all sheets to a new Excel file
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    for sheet_name, sheet_df in other_sheets.items():
                        sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)

                print(f"Cleaned output saved to {output_file}")

# Run the cleaning function for the current directory
clean_timeline(os.getcwd())
