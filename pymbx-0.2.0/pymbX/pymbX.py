import os
import pandas as pd
import re
import os

def ezclean(microbiome_data, metadata, level="d"):
    # -------------------------------
    # 1. File Extension Checks and Data Reading
    # -------------------------------
    # Check the file extension for microbiome_data
    microbiome_ext = os.path.splitext(microbiome_data)[1].lower().lstrip('.')
    if microbiome_ext not in ['csv', 'xls', 'xlsx', 'txt']:
        return "The file is not csv, xls, xlsx, or txt format. Please check the file type for microbiome data."
    
    # Check the file extension for metadata and read accordingly
    metadata_ext = os.path.splitext(metadata)[1].lower().lstrip('.')
    if metadata_ext == "txt":
        metadata_df = pd.read_csv(metadata, sep="\t", header=0, dtype=str)
    elif metadata_ext == "csv":
        metadata_df = pd.read_csv(metadata, header=0, dtype=str)
    elif metadata_ext in ["xls", "xlsx"]:
        metadata_df = pd.read_excel(metadata, header=0, dtype=str)
        metadata_df = pd.DataFrame(metadata_df)
    else:
        return "Please check the file format of metadata."
    
    # -------------------------------
    # 2. Metadata Header Check
    # -------------------------------
    valid_headers = ["id", "sampleid", "sample id", "sample-id", "featureid", "feature id", "feature-id"]
    first_header = metadata_df.columns[0].strip().lower()
    if first_header not in valid_headers:
        return "Please check the first header of the metadata for file format correction."
    
    # -------------------------------
    # 3. Read Microbiome Data
    # -------------------------------
    # Note: To preserve numeric values, do not force dtype=str.
    if microbiome_ext == "txt":
        microbiome_df = pd.read_csv(microbiome_data, sep="\t", header=0)
    elif microbiome_ext == "csv":
        microbiome_df = pd.read_csv(microbiome_data, header=0)
    elif microbiome_ext in ["xls", "xlsx"]:
        # Skip the first row (mimicking an R skip index row) but do not force string conversion
        microbiome_df = pd.read_excel(microbiome_data, skiprows=1)
    else:
        return "The microbiome file is not in a supported format. Please use txt, csv, xls, or xlsx."
    
    # -------------------------------
    # 4. Taxonomic Level Mapping
    # -------------------------------
    levels_map = {
        "domain": "d__", "Domain": "d__", "DOMAIN": "d__", "D": "d__", "d": "d__",
        "kingdom": "d__", "Kingdom": "d__", "KINGDOM": "d__", "K": "d__", "k": "d__",
        "phylum": "p__", "Phylum": "p__", "PHYLUM": "p__", "P": "p__", "p": "p__",
        "class": "c__", "Class": "c__", "CLASS": "c__", "C": "c__", "c": "c__",
        "order": "o__", "Order": "o__", "ORDER": "o__", "O": "o__", "o": "o__",
        "family": "f__", "Family": "f__", "FAMILY": "f__", "F": "f__", "f": "f__",
        "genera": "g__", "genus": "g__", "Genera": "g__", "GENERA": "g__", "G": "g__", "g": "g__",
        "species": "s__", "Species": "s__", "SPECIES": "s__", "S": "s__", "s": "s__"
    }
    
    level_key = levels_map.get(level.lower())
    if level_key is None:
        return ("The level value should be one of the following: domain, phylum, class, order, "
                "family, genera, species or their abbreviations.")
    level_value = level_key
    
    # -------------------------------
    # 5. Separate Microbiome and Metadata Columns
    # -------------------------------
    common_cols = list(set(microbiome_df.columns).intersection(set(metadata_df.columns)))
    just_microbiome_df = microbiome_df.loc[:, ~microbiome_df.columns.isin(common_cols)]
    just_metadata_df = metadata_df.loc[:, metadata_df.columns.isin(common_cols)]
    
    # Save the separated DataFrames as Excel files
    just_microbiome_df.to_excel("just_microbiome.xlsx", index=False)
    just_metadata_df.to_excel("just_metadata.xlsx", index=False)
    
    # -------------------------------
    # 6. Transpose the Microbiome Data
# Read the entire Excel file without assigning any row as header.
    orig = pd.read_excel("just_microbiome.xlsx", header=None)
    
    # Define the new column headers from the original first column (skip the top-left cell)
    new_columns = orig.iloc[1:, 0].tolist()
    
    # Define the new row index from the original header row (skip the top-left cell)
    new_index = orig.iloc[0, 1:].tolist()
    
    # Extract the data block that excludes the original header row and first column.
    data_block = orig.iloc[1:, 1:]
    
    # Transpose the data block so that:
    # - The rows become the new row labels (matching new_index)
    # - The columns become the new column headers (matching new_columns)
    data_transposed = data_block.T
    
    # Create the new DataFrame with the desired orientation.
    df_new = pd.DataFrame(data_transposed.values, index=new_index, columns=new_columns)
    
    # Save the new DataFrame to an Excel file.
    df_new.to_excel("microbiome_ezy_2.xlsx", index=True)


    #NOTE FOR MYSELF  SINCE I USE THE MEZY 1 WHICH WAS IN TEXT AND MEZY2 WAS THE ONE WITH NUMERICAL
    #I DID THE DIRECT NUMERICAL SO I SKIPPED MEZY_1 AND NAMED IT AS MEZY_2
    
    # -------------------------------
    # 8. Add a Blank "Taxa" Column as the Second Column
    # -------------------------------
    df_ezy2 = pd.read_excel("microbiome_ezy_2.xlsx", header=None)
    first_col = df_ezy2.iloc[:, [0]]
    blank_col = pd.DataFrame([""] * df_ezy2.shape[0])
    remaining_cols = df_ezy2.iloc[:, 1:]
    new_matrix = pd.concat([first_col, blank_col, remaining_cols], axis=1, ignore_index=True)
    new_matrix.iat[0, 1] = "Taxa"
    new_matrix.to_excel("microbiome_ezy_3.xlsx", index=False, header=False)
    
    # -------------------------------
    # 9. Copy microbiome_ezy_3.xlsx to microbiome_ezy_4.xlsx Without Modification
    # -------------------------------
    df_ezy3 = pd.read_excel("microbiome_ezy_3.xlsx", header=None)
    df_ezy3.to_excel("microbiome_ezy_4.xlsx", index=False, header=False)
    
    # -------------------------------
    # 10. Extract and Construct Taxa Names into the "Taxa" Column
    # -------------------------------
    df_ezy4 = pd.read_excel("microbiome_ezy_4.xlsx", header=None)
    
    levels_order = ["d__", "p__", "c__", "o__", "f__", "g__", "s__"]
    level_mapping = {
        "d__": "domain",
        "p__": "phylum",
        "c__": "class",
        "o__": "order",
        "f__": "family",
        "g__": "genus",
        "s__": "species"
    }
    
    blank_counter = [0]
    
    def extract_taxa(x):
        if pd.isna(x):
            return ""
        x = str(x)
        if level_value in x:
            pattern_target = ".*" + re.escape(level_value) + r"\s*([^;]*)(;.*)?$"
            m = re.search(pattern_target, x)
            candidate = m.group(1).strip() if m and m.group(1) is not None else ""
        else:
            candidate = ""
        if candidate != "":
            return candidate
        else:
            try:
                target_idx = levels_order.index(level_value) + 1
            except ValueError:
                target_idx = None
            found_value = ""
            found_marker = ""
            if target_idx is not None and target_idx > 1:
                for i in range(target_idx - 2, -1, -1):
                    if levels_order[i] in x:
                        pattern_prev = ".*" + re.escape(levels_order[i]) + r"\s*([^;]*)(;.*)?$"
                        m_prev = re.search(pattern_prev, x)
                        candidate_prev = m_prev.group(1).strip() if m_prev and m_prev.group(1) is not None else ""
                        if candidate_prev != "":
                            found_value = candidate_prev
                            found_marker = levels_order[i]
                            break
            if found_value != "":
                blank_counter[0] += 1
                return ("unidentified_" + level_mapping[level_value] + "_" + str(blank_counter[0]) +
                        "_at_" + found_value + "_" + level_mapping[found_marker])
            else:
                return ""
    
    if df_ezy4.shape[0] > 1:
        extracted_text = df_ezy4.iloc[1:, 0].apply(extract_taxa)
        df_ezy4.iloc[1:, 1] = extracted_text
    
    df_ezy4.to_excel("microbiome_ezy_5.xlsx", index=False, header=False)
    
    
    
    
    
    
    
    
    
    # ---- Additional Delete First Column Block Start ----
    # Read the microbiome_ezy_5.xlsx file (all data, including header)
    df_ezy5 = pd.read_excel("microbiome_ezy_5.xlsx", header=None)
    
    # Delete the first column entirely
    df_ezy5 = df_ezy5.iloc[:, 1:]
    
    # Save the updated data as microbiome_ezy_6.xlsx without extra row or column names
    df_ezy5.to_excel("microbiome_ezy_7.xlsx", index=False, header=False)
    # ---- Additional Delete First Column Block End ----
    
        
    
    
    
    
    
    
    
    
    
        # ---- Additional Aggregation Block Start ----
   # ---- Additional Aggregation Block Start ----
    # Read the "microbiome_ezy_7.xlsx" file (all data, including the header row) without column names
    df_ezy7 = pd.read_excel("microbiome_ezy_7.xlsx", header=None)
    
    # Separate the header row (first row) and the data rows (remaining rows)
    header_row_7 = df_ezy7.iloc[0, :].tolist()
    data_rows_7 = df_ezy7.iloc[1:, :].copy()
    
    # Ensure the first column is treated as text and all other columns are numeric
    data_rows_7[0] = data_rows_7[0].astype(str)
    for col in data_rows_7.columns[1:]:
        data_rows_7[col] = pd.to_numeric(data_rows_7[col], errors='coerce')
    
    # Define replacement names based on the level_value
    other_names = {
        "d__": "Other_domains",
        "p__": "Other_phyla",
        "c__": "Other_classes",
        "o__": "Other_orders",
        "f__": "Other_families",
        "g__": "Other_genera",
        "s__": "Other_species"
    }
    
    # Determine the appropriate replacement name using the level_value.
    replacement_name = other_names.get(level_value, "Other")
    
    # Replace rows with missing taxa information (NA, blank, "#VALUE", or "nan" as a string)
    data_rows_7[0] = data_rows_7[0].apply(
        lambda x: replacement_name if (pd.isna(x) or str(x).strip() in ["", "#VALUE", "nan"]) else x
    )
    
    # Aggregate the data by the taxa (first) column:
    # For rows with the exact same taxa value, sum the numeric columns column-wise.
    aggregated_df = data_rows_7.groupby(0, as_index=False).sum()
    
    # Set the column names of the aggregated data using the original header row.
    new_header = header_row_7.copy()
    new_header[0] = "Taxa"
    aggregated_df.columns = new_header
    
    # Write the aggregated data (including header) to "microbiome_ezy_8.xlsx"
    aggregated_df.to_excel("microbiome_ezy_8.xlsx", index=False)
    # ---- Additional Aggregation Block End ----

    
        
    ###3333 NOW CONVERTING INTO THE PERCENTAGE

    
    # ---- Additional Column-wise Percentage Block Start ----
    # Read the "microbiome_ezy_8.xlsx" file with header information
    microbiome_ezy8 = pd.read_excel("microbiome_ezy_8.xlsx", header=0)
    
    # Extract the taxa column (first column) and numeric columns (columns 2 onward)
    taxa_column = microbiome_ezy8.iloc[:, 0]
    numeric_columns = microbiome_ezy8.iloc[:, 1:].copy()
    
    # Ensure all numeric columns are properly converted to numeric
    numeric_columns = numeric_columns.apply(pd.to_numeric, errors='coerce')
    
    # Calculate the column-wise percentages:
    # For each numeric column, divide each element by the column sum (ignoring NAs) and multiply by 100.
    # Note: Use axis=1 to align the sums with the columns.
    col_sums = numeric_columns.sum(axis=0, skipna=True)
    percentage_columns = numeric_columns.div(col_sums, axis=1) * 100
    
    # Combine the taxa column with the calculated percentage columns
    microbiome_ezy9 = pd.concat([taxa_column, percentage_columns], axis=1)
    
    # Save the resulting data frame as "microbiome_ezy_9.xlsx" without row names
    microbiome_ezy9.to_excel("microbiome_ezy_9.xlsx", index=False)
    # ---- Additional Column-wise Percentage Block End ----



#####################--------------------------------------------------------------   


###TRANSPOSE THE EXCEL



    # Read the entire Excel file (microbiome_ezy_9.xlsx) without assigning any row as a header.
    df_ezy_9 = pd.read_excel("microbiome_ezy_9.xlsx", header=None)
    
    # Define new column headers from the original first column (skipping the top-left cell)
    col_headers_ezy_9 = df_ezy_9.iloc[1:, 0].tolist()
    
    # Define new row labels from the original header row (skipping the top-left cell)
    row_labels_ezy_9 = df_ezy_9.iloc[0, 1:].tolist()
    
    # Extract the data block that excludes the original header row and first column.
    data_slice_ezy_9 = df_ezy_9.iloc[1:, 1:]
    
    # Transpose the data block so that:
    # - The rows become the new row labels (matching row_labels_ezy_9)
    # - The columns become the new column headers (matching col_headers_ezy_9)
    data_trans_ezy_9 = data_slice_ezy_9.T
    
    # Create the new DataFrame with the desired orientation.
    df_ezy_10 = pd.DataFrame(data_trans_ezy_9.values, index=row_labels_ezy_9, columns=col_headers_ezy_9)
    
    # Save the new DataFrame to an Excel file (microbiome_ezy_10.xlsx).
    df_ezy_10.to_excel("microbiome_ezy_10.xlsx", index=True)




####33CHANGE THE FIRST HEADER FROM THE METADATA 


    # ---- Additional Header Replacement Block Start ----
    # Extract the header of the first column from the metadata file
    metadata_first_header = metadata_df.columns[0]
    
    # Read the "microbiome_ezy_10.xlsx" file; since it was written without column names,
    # the first row actually contains the header information.
    df_ezy10 = pd.read_excel("microbiome_ezy_10.xlsx", header=None, dtype=str)
    
    # Extract the first row as the current header and convert to a list of strings
    new_header = df_ezy10.iloc[0, :].astype(str).tolist()
    
    # Replace the first element with the metadata header
    new_header[0] = metadata_first_header
    
    # Remove the first row (which contained the original header) from the data
    df_ezy10_data = df_ezy10.iloc[1:, :].copy()
    
    # Assign the new header to the data
    df_ezy10_data.columns = new_header
    
    # Save the updated data as "microbiome_ezy_11.xlsx" with proper column names
    df_ezy10_data.to_excel("microbiome_ezy_11.xlsx", index=False, header=True)
    # ---- Additional Header Replacement Block End ----




#####EVERYTHING IS  TEXT FOR THE EASE OF DATA HANDLING SO I WILL CONVERT THE NUMERICALS





    # ---- Additional Numeric Conversion Block for ezy_12 Start ----
    # Read the "microbiome_ezy_11.xlsx" file (all values as text) without column names
    df_ezy_11 = pd.read_excel("microbiome_ezy_11.xlsx", header=None, dtype=str)
    
    # Separate header row (first row) and data rows (remaining rows)
    header_row_ezy_11 = df_ezy_11.iloc[0, :].tolist()
    data_rows_ezy_11 = df_ezy_11.iloc[1:, :].copy()
    
    # Convert all columns except the first in data_rows_ezy_11 to numeric
    for col in data_rows_ezy_11.columns[1:]:
        data_rows_ezy_11[col] = pd.to_numeric(data_rows_ezy_11[col], errors='coerce')
    
    # Create a new Excel file and write the header row as text (transposed) and data rows starting at row 2
    with pd.ExcelWriter("microbiome_ezy_12.xlsx", engine="openpyxl") as writer:
        # Write the header row as a single-row DataFrame at row 1
        pd.DataFrame([header_row_ezy_11]).to_excel(writer, sheet_name="Sheet1", index=False, header=False, startrow=0)
        # Write the data rows starting at row 2 (i.e., startrow=1)
        data_rows_ezy_11.to_excel(writer, sheet_name="Sheet1", index=False, header=False, startrow=1)
    
    
    
#####333 NOW I WILL ADD THE METADATA COLUMNS TO THE CLEANED DATA SET



    # ---- Additional Metadata Merge Block Start ----
    # Read the "microbiome_ezy_12.xlsx" file (with header)
    df_ezy12 = pd.read_excel("microbiome_ezy_12.xlsx", header=0)
    
    # In the metadata, the first column is the key.
    # Instead of using meta_remaining (which lacks the key column),
    # build a keyed DataFrame from the full metadata.
    metadata_keyed = metadata_df.set_index(metadata_df.columns[0])
    
    # Extract the key values from df_ezy12 (the cleaned data)
    df_key = df_ezy12.iloc[:, 0]
    
    # Reindex the keyed metadata with df_key.
    # Then select only the columns corresponding to the original meta_remaining,
    # i.e. all columns except the first.
    appended_metadata = metadata_keyed.reindex(df_key)[metadata_df.columns[1:]].reset_index(drop=True)
    
    # Combine the df_ezy12 data with the appended metadata columns.
    # Reset the index of df_ezy12 to ensure proper alignment.
    df_ezy13 = pd.concat([df_ezy12.reset_index(drop=True), appended_metadata], axis=1)
    
    # Save the combined data as "microbiome_ezy_13.xlsx" with headers
    df_ezy13.to_excel("microbiome_ezy_13.xlsx", index=False)
    # ---- Additional Metadata Merge Block End ----
    
    



##### REMOVING ANY INTERMEDIATE FILES IN THE WORKING DIR 


    # List of intermediate files to be removed (excluding the final file "microbiome_ezy_13.xlsx")
    intermediate_files = [
        "just_microbiome.xlsx",
        "just_metadata.xlsx",
        "microbiome_ezy_2.xlsx",
        "microbiome_ezy_3.xlsx",
        "microbiome_ezy_4.xlsx",
        "microbiome_ezy_5.xlsx",
        "microbiome_ezy_6.xlsx",
        "microbiome_ezy_7.xlsx",
        "microbiome_ezy_8.xlsx",
        "microbiome_ezy_9.xlsx",
        "microbiome_ezy_10.xlsx",
        "microbiome_ezy_11.xlsx",
        "microbiome_ezy_12.xlsx"
    ]
    
    # Remove each intermediate file if it exists
    for f in intermediate_files:
        if os.path.exists(f):
            os.remove(f)

##FINAL RETURN BASED ON THE TAXA LEVEL IN THE INPUT OF THE FUNCTION

    # ---- Final Return Block ----
    final_names = {
        "d__": "mbX_cleaned_domains_or_kingdom.xlsx",
        "p__": "mbX_cleaned_phylum.xlsx",
        "c__": "mbX_cleaned_classes.xlsx",
        "o__": "mbX_cleaned_orders.xlsx",
        "f__": "mbX_cleaned_families.xlsx",
        "g__": "mbX_cleaned_genera.xlsx",
        "s__": "mbX_cleaned_species.xlsx"
    }
    
    final_file_name = final_names[level_value]
    os.rename("microbiome_ezy_13.xlsx", final_file_name)
    
    return final_file_name
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    return "Processing complete."

# Example usage:
# result = ezclean("microbiome2.csv", "metadata2.txt", "s")
# print(result)












#ezclean("microbiome2.csv", "metadata2.txt", "s")
###-_________________________________________________________________________________
import os
import pandas as pd
import pandas as pd
import pandas.api.types as ptypes



def ezviz(microbiome_data, metadata, level, selected_metadata, top_taxa=None, threshold=None):
    import os
    import pandas as pd
    import numpy as np
    # Check the file extension for microbiome_data
    microbiome_ext = os.path.splitext(microbiome_data)[1].lower().lstrip('.')
    if microbiome_ext not in ['csv', 'xls', 'xlsx']:
        return "The file is not csv, xls, or xlsx format. Please check the file type for microbiome data."
    
    # Check the file extension for metadata
    metadata_ext = os.path.splitext(metadata)[1].lower().lstrip('.')
    if metadata_ext == 'txt':
        metadata_df = pd.read_csv(metadata, sep="\t", header=0)
    elif metadata_ext == 'csv':
        metadata_df = pd.read_csv(metadata, header=0)
    elif metadata_ext in ['xls', 'xlsx']:
        metadata_df = pd.read_excel(metadata, header=0, engine='openpyxl')
    else:
        return "Please check the file format of metadata."
    











    # Check the first header of metadata
    valid_headers = ["id", "sampleid", "sample id", "sample-id", 
                     "featureid", "feature id", "feature-id"]
    first_header = metadata_df.columns[0].strip().lower()
    if first_header not in valid_headers:
        return "Please check the first header of the metadata for file format correction."
    
    # Check if the selected_metadata is a valid categorical column in metadata_df.
    # We consider categorical columns as those with dtype 'object' (strings) or 'category'
    if (selected_metadata not in metadata_df.columns or 
        not (ptypes.is_categorical_dtype(metadata_df[selected_metadata]) or 
             ptypes.is_object_dtype(metadata_df[selected_metadata]))):
        return "The selected metadata is either not in the metadata or not a categorical value"
    








    # Read microbiome data based on file extension
    if microbiome_ext == "txt":
        microbiome_df = pd.read_csv(microbiome_data, sep="\t", header=0)
    elif microbiome_ext == "csv":
        microbiome_df = pd.read_csv(microbiome_data, header=0)
    elif microbiome_ext in ["xls", "xlsx"]:
        # Mimic the R behavior of skipping the first row
        microbiome_df = pd.read_excel(microbiome_data, skiprows=1, engine="openpyxl")
    else:
        return "The microbiome file is not in a supported format. Please use txt, csv, xls, or xlsx."







    
    levels_map = {
        "domain": "d__",
        "Domain": "d__",
        "DOMAIN": "d__",
        "D": "d__",
        "d": "d__",
        "kingdom": "d__",
        "Kingdom": "d__",
        "KINGDOM": "d__",
        "K": "d__",
        "k": "d__",
        "phylum": "p__",
        "Phylum": "p__",
        "PHYLUM": "p__",
        "P": "p__",
        "p": "p__",
        "class": "c__",
        "Class": "c__",
        "CLASS": "c__",
        "C": "c__",
        "c": "c__",
        "order": "o__",
        "Order": "o__",
        "ORDER": "o__",
        "O": "o__",
        "o": "o__",
        "family": "f__",
        "Family": "f__",
        "FAMILY": "f__",
        "F": "f__",
        "f": "f__",
        "genera": "g__",
        "Genera": "g__",
        "GENERA": "g__",
        "G": "g__",
        "g": "g__",
        "species": "s__",
        "Species": "s__",
        "SPECIES": "s__",
        "S": "s__",
        "s": "s__"
    }




    # Get the corresponding level code using the lower-case version of level
    level_code = levels_map.get(level.lower())
    if level_code is None:
        return "Invalid taxonomic level provided."
    
    # Ensure that only one of top_taxa or threshold is provided
    if top_taxa is not None and threshold is not None:
        return "Only one of the parameter can be selected between top_taxa and threshold"
    
    # Determine the threshold value based on the provided parameter
    if threshold is not None:
        threshold_value = threshold
    elif top_taxa is not None:
        threshold_value = top_taxa
    else:
        threshold_value = None
    
    ##### File Handling: Run ezclean and use its output #####
    cleaned_file_ezviz = ezclean(microbiome_data, metadata, level)
    print("ezclean returned:", cleaned_file_ezviz)
    
    import os
    if not os.path.exists(cleaned_file_ezviz):
        raise FileNotFoundError(f"The file returned by ezclean does not exist: {cleaned_file_ezviz}")
    







##READING THE CLEANED FILE FROM THE EZCLEAN

    cleaned_data = pd.read_excel(cleaned_file_ezviz, header=0, engine="openpyxl")






    # Identify metadata columns in cleaned_data that are also present in metadata_df
    metadata_cols_in_cleaned = list(set(metadata_df.columns).intersection(cleaned_data.columns))
    
    # Check that the selected_metadata is indeed present in the cleaned data
    if selected_metadata not in metadata_cols_in_cleaned:
        raise Exception(f"The selected metadata column '{selected_metadata}' was not found in the cleaned data.")
    
    # Determine which metadata columns (that are present) should be removed
    metadata_to_remove = set(metadata_cols_in_cleaned) - {selected_metadata}
    
    # Subset the cleaned_data to keep only non-metadata columns plus the selected_metadata column
    subset_data = cleaned_data.loc[:, [col for col in cleaned_data.columns if col not in metadata_to_remove]]
    
    # Write the resulting subset_data to "mbX_cleaning_1.xlsx"
    subset_data.to_excel("mbX_cleaning_1.xlsx", index=False)
    
    
    









  
    
        # Read in the "mbX_cleaning_1.xlsx" file
    data_cleaning_mbX_1 = pd.read_excel("mbX_cleaning_1.xlsx", engine="openpyxl")
    
    # Define invalid values for the selected metadata column
    invalid_values = ["", "Na", "NA", "#VALUE!", "#NAME?"]
    
    # Filter out rows where the selected_metadata column is empty, NA, or contains an invalid value
    data_cleaning_mbX_1 = data_cleaning_mbX_1[
        data_cleaning_mbX_1[selected_metadata].notna() &
        (~data_cleaning_mbX_1[selected_metadata].isin(invalid_values))
    ]
    
    # Group the data by the selected_metadata column and calculate the mean for each numeric column
    numeric_cols = data_cleaning_mbX_1.select_dtypes(include=[np.number]).columns
    data_summary_mbX_2 = data_cleaning_mbX_1.groupby(selected_metadata)[numeric_cols].mean().reset_index()
    
    # Write the summarized data to "mbX_cleaning_2.xlsx"
    data_summary_mbX_2.to_excel("mbX_cleaning_2.xlsx", index=False)
    
        
    




    # Read in the "mbX_cleaning_2.xlsx" file with headers
    data_cleaning_mbX_2 = pd.read_excel("mbX_cleaning_2.xlsx", engine="openpyxl")
    
    # Transpose the DataFrame using .T (this converts rows to columns and vice versa)
    data_transposed_mbX_3 = data_cleaning_mbX_2.T
    
    # Write the transposed DataFrame to "mbX_cleaning_3.xlsx", including the index (row names)
    data_transposed_mbX_3.to_excel("mbX_cleaning_3.xlsx", index=True)
    












    # Read in the "mbX_cleaning_3.xlsx" file without using a header
    raw_data_mbX_3 = pd.read_excel("mbX_cleaning_3.xlsx", header=None, engine="openpyxl")
    
    # Remove the very first row (assumed to be the unwanted index row)
    raw_data_no_index = raw_data_mbX_3.iloc[1:, :]
    
    # Now, the first row of raw_data_no_index is the actual header.
    header_row = raw_data_no_index.iloc[0]
    data_without_index = raw_data_no_index.iloc[1:, :].copy()
    
    # Assign the header row as column names
    data_without_index.columns = header_row
    
    # Write the cleaned data to "mbX_cleaning_4.xlsx" without row names
    data_without_index.to_excel("mbX_cleaning_5.xlsx", index=False)
    
    
    









    
    
    import pandas as pd
    import numpy as np
    import os
    
    # Read in "mbX_cleaning_5.xlsx"
    df_clean5 = pd.read_excel("mbX_cleaning_5.xlsx", engine="openpyxl")
    
    # Assume the first column contains taxon names and the remaining columns are numeric.
    # Compute the row averages (ignoring the first column)
    numeric_matrix = df_clean5.iloc[:, 1:]
    row_avg = numeric_matrix.mean(axis=1, skipna=True)
    
    # Attach the computed averages as a new column "RowAvg"
    df_clean5["RowAvg"] = row_avg
    
    # Sort the data frame in descending order by the row average
    df_sorted = df_clean5.sort_values(by="RowAvg", ascending=False).reset_index(drop=True)
    
    # Define the mapping for the aggregated "Other" row names
    other_name_map = {
        "d__": "Other_domains",
        "p__": "Other_phyla",
        "c__": "Other_classes",
        "o__": "Other_orders",
        "f__": "Other_families",
        "g__": "Other_genera",
        "s__": "Other_species"
    }
    
    # Retrieve the proper "Other" name based on level_code (assumed defined previously)
    other_row_name = other_name_map.get(level_code)
    
    # Initialize final data frame variable
    final_df = None
    
    if threshold is not None:
        # --- THRESHOLD LOGIC ---
        # Keep rows with average >= threshold
        rows_keep = df_sorted[df_sorted["RowAvg"] >= threshold].copy()
        # Rows with average below threshold will be aggregated
        rows_agg = df_sorted[df_sorted["RowAvg"] < threshold].copy()
        
        if not rows_agg.empty:
            # Sum the numeric values (columns 2 to second-to-last, excluding the first taxon column and the last "RowAvg" column)
            agg_values = rows_agg.iloc[:, 1:-1].sum(skipna=True)
            # Create a new row: first column is the Other row name; numeric columns are the aggregated sums.
            agg_row = pd.DataFrame([[""] * df_sorted.shape[1]], columns=df_sorted.columns)
            agg_row.iloc[0, 0] = other_row_name
            agg_row.iloc[0, 1:-1] = agg_values.values
            agg_row.iloc[0, -1] = np.nan
            # Combine the kept rows with the aggregated row.
            final_df = pd.concat([rows_keep, agg_row], ignore_index=True)
        else:
            final_df = df_sorted.copy()
            
    elif top_taxa is not None:
        # --- TOP_TAXA LOGIC ---
        if df_sorted.shape[0] > top_taxa:
            rows_keep = df_sorted.iloc[:top_taxa, :].copy()
            rows_agg = df_sorted.iloc[top_taxa:, :].copy()
            
            if not rows_agg.empty:
                agg_values = rows_agg.iloc[:, 1:-1].sum(skipna=True)
                agg_row = pd.DataFrame([[""] * df_sorted.shape[1]], columns=df_sorted.columns)
                agg_row.iloc[0, 0] = other_row_name
                agg_row.iloc[0, 1:-1] = agg_values.values
                agg_row.iloc[0, -1] = np.nan
                final_df = pd.concat([rows_keep, agg_row], ignore_index=True)
            else:
                final_df = df_sorted.copy()
        else:
            final_df = df_sorted.copy()
    else:
        final_df = df_sorted.copy()
    
    # Remove the helper "RowAvg" column before writing the output
    final_df = final_df.drop(columns=["RowAvg"])
    
    # Define a mapping for visualization data file names based on level_code
    vizDataNames = {
        "d__": "mbX_vizualization_data_domains.xlsx",
        "p__": "mbX_vizualization_data_phyla.xlsx",
        "c__": "mbX_vizualization_data_classes.xlsx",
        "o__": "mbX_vizualization_data_orders.xlsx",
        "f__": "mbX_vizualization_data_families.xlsx",
        "g__": "mbX_vizualization_data_genera.xlsx",
        "s__": "mbX_vizualization_data_species.xlsx"
    }
    
    # Determine the output file name based on level_code
    outputVizFile = vizDataNames.get(level_code)
    
    # Write the final data frame to the dynamic file name without row names
    final_df.to_excel(outputVizFile, index=False)
    
    # --- Later, when reading the file for plotting ---
    
    # Check if the file exists before reading
    if not os.path.exists(outputVizFile):
        raise FileNotFoundError(f"The file {outputVizFile} does not exist in the working directory: {os.getcwd()}")
    
    df_clean6 = pd.read_excel(outputVizFile, engine="openpyxl")
    
    
    
    




    
    
    import re
    import os
    import gc
    import warnings
    import numpy as np
    import pandas as pd
    import matplotlib.pyplot as plt
    
    # --- LEN FOR THE ROTATION OF THE TEXT IN THE X AXIS ---
    
    # Compute maximum header length (excluding the first column header)
    headers_to_check = list(df_clean6.columns[1:])  # all headers except the first
    # Remove all whitespace from each header
    headers_clean = [re.sub(r"\s+", "", h) for h in headers_to_check]
    max_header_length = max([len(h) for h in headers_clean]) if headers_clean else 0
    #print("Maximum header length (excluding the first header) is:", max_header_length)
    
    # Set axis text parameters based on max_header_length
    if max_header_length > 9:
        axis_x_angle = 65
        axis_x_ha = "right"  # matplotlib accepts "right" (equivalent to hjust = 1.0)
    else:
        axis_x_angle = 0
        axis_x_ha = "center"  # equivalent to hjust = 0.5
    
    # --- DYNAMIC PLOTTING CODE BLOCK USING CUSTOM COLORS ---
    
    # Define the mapping for taxonomic levels
    level_names = {
        "d__": "Domains",
        "p__": "Phyla",
        "c__": "Classes",
        "o__": "Orders",
        "f__": "Families",
        "g__": "Genera",
        "s__": "Species"
    }
    
    # Clear unused variables and run garbage collection
    gc.collect()
    
    # Retrieve the descriptive name for the current level from the mapping
    taxon_descriptor = level_names.get(level_code)
    if taxon_descriptor is None:
        raise Exception("Invalid level_code provided.")
    
    # Build dynamic titles for the plot and legend
    plot_title = f"Relative Abundance of Microbial {taxon_descriptor}"
    legend_title = f"Microorganism {taxon_descriptor}"
    
    # Reshape the data from wide to long format
    # The first column contains taxon names; remaining columns represent samples.
    taxa_col = df_clean6.columns[0]
    df_long = df_clean6.melt(id_vars=[taxa_col], var_name="Sample", value_name="Abundance")
    
    # Ensure the taxon column is treated as categorical (preserving the original order)
    df_long[taxa_col] = pd.Categorical(df_long[taxa_col],
                                       categories=df_clean6[taxa_col].tolist(),
                                       ordered=True)
    
    # Define the two base palettes
    tab10 = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
             "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf"]
    set3 = ["#8dd3c7", "#ffffb3", "#bebada", "#fb8072", "#80b1d3",
            "#fdb462", "#b3de69", "#fccde5", "#d9d9d9", "#bc80bd"]
    
    # We want 200 colors in total: 100 values from each palette.
    tab10_rep = (tab10 * 10)[:100]
    set3_rep = (set3 * 10)[:100]
    
    # Interleave the two repeated palettes
    custom_colors = []
    for t, s in zip(tab10_rep, set3_rep):
        custom_colors.append(t)
        custom_colors.append(s)
    
    # Calculate the number of distinct taxa (for distinct colors)
    unique_taxa = df_long[taxa_col].unique()
    n_colors = len(unique_taxa)
    
    # Subset or repeat custom_colors to match the number of taxa
    if n_colors > len(custom_colors):
        warnings.warn("Number of taxa exceeds custom color length; colors will be recycled.")
        palette_colors = (custom_colors * ((n_colors // len(custom_colors)) + 1))[:n_colors]
    else:
        palette_colors = custom_colors[:n_colors]
    
    # Map each taxon to a color
    taxon_to_color = dict(zip(unique_taxa, palette_colors))
    
    # Create a pivot table for the stacked bar chart:
    # Rows = Sample; Columns = taxon names; Values = Abundance
    pivot_df = df_long.pivot(index="Sample", columns=taxa_col, values="Abundance").fillna(0)
    pivot_df = pivot_df.sort_index()  # sort samples alphabetically
    
    # Create the stacked bar plot using matplotlib
    fig, ax = plt.subplots()
    bottom = np.zeros(len(pivot_df))
    # Use the order of taxa as in the original df_clean6 for consistent ordering
    for taxon in df_clean6[taxa_col]:
        if taxon in pivot_df.columns:
            values = pivot_df[taxon].values
            ax.bar(pivot_df.index, values, bottom=bottom,
                   label=taxon, color=taxon_to_color.get(taxon))
            bottom += values
    
    ax.set_title(plot_title, fontsize=18, fontweight="bold", loc="center")
    ax.set_xlabel(selected_metadata)
    ax.set_ylabel("Relative Abundance (%)")
    plt.xticks(rotation=axis_x_angle, ha=axis_x_ha, fontsize=12)
    # Reverse legend order so that it matches the stacking order
    handles, labels = ax.get_legend_handles_labels()
    ax.legend(handles[::-1], labels[::-1], title=legend_title, fontsize=12, title_fontsize=14,
              loc="center left", bbox_to_anchor=(1, 0.5))
    plt.tight_layout(rect=[0, 0, 0.85, 1])


    # Determine dynamic plot dimensions based on the shape of df_clean6
    df_rows, df_cols = df_clean6.shape
    if df_cols > 4:
        plot_width = 12 + 0.7 * (df_cols - 4)
    else:
        plot_width = 12
    if df_rows > 55:
        plot_height = 15 + 0.35 * (df_rows - 55)
    else:
        plot_height = 15
    #print("Calculated plot width:", plot_width, "and height:", plot_height)
    
    fig.set_size_inches(plot_width, plot_height)
    
    # Define output filename dynamically based on level_code
    final_names_viz = {
        "d__": "mbX_viz_domains_or_kingdom",
        "p__": "mbX_viz_phylum",
        "c__": "mbX_viz_classes",
        "o__": "mbX_viz_orders",
        "f__": "mbX_viz_families",
        "g__": "mbX_viz_genera",
        "s__": "mbX_viz_species"
    }
    output_plot_filename = final_names_viz.get(level_code, "mbX_viz") + ".pdf"
    
    # Save the plot with high DPI and the calculated dimensions
    plt.savefig(output_plot_filename, dpi=1200, bbox_inches="tight")
    plt.close(fig)
    #print(f"Output plot '{output_plot_filename}' has been created.")
    
    # Delete temporary cleaning files before closing the function
    temp_files = ["mbX_cleaning_1.xlsx", "mbX_cleaning_2.xlsx", 
                  "mbX_cleaning_3.xlsx", "mbX_cleaning_4.xlsx", 
                  "mbX_cleaning_5.xlsx"]
    for file in temp_files:
        if os.path.exists(file):
            os.remove(file)
            #print("Deleted file:", file)
        # else:
            #print("File not found (already deleted or never created):", file)
    
    print("Done with the visualization, cite us!")
    
    


























####ezviz("microbiome2.csv", "metadata2.txt", "g", "Treatmentdescription", top_taxa=10)

####ezviz("microbiome2.csv", "metadata2.txt", "g", "check1", threshold=0.1)