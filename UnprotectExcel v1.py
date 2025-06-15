import zipfile
import io
import os
import re
import argparse # Import the argparse module

def remove_protection_from_xlsx_regex(input_xlsx_path, output_xlsx_path):
    """
    Removes worksheet and workbook protection tags from an XLSX file
    using regular expressions to modify the XML content directly.

    Args:
        input_xlsx_path (str): Path to the input XLSX file.
        output_xlsx_path (str): Path to save the modified XLSX file.

    Returns:
        bool: True if successful, False otherwise.
    """
    print(f"Processing file: {input_xlsx_path}")
    removed_tags_count = 0
    files_modified_count = 0

    try:
        # Use temporary storage for modified XML data in memory
        temp_xml_data = {}

        # --- Step 1: Read the input XLSX (zip) file ---
        try:
            with zipfile.ZipFile(input_xlsx_path, 'r') as zin:
                all_files_data = {item.filename: zin.read(item.filename) for item in zin.infolist()}
        except FileNotFoundError:
            print(f"Error: Input file not found at {input_xlsx_path}")
            return False
        except zipfile.BadZipFile:
            print(f"Error: Input file is not a valid XLSX (Zip) file: {input_xlsx_path}")
            return False


        # --- Step 2: Iterate through the extracted file data ---
        for filename, file_content in all_files_data.items():
            # --- Step 3: Check if it's a relevant XML file ---
            if filename.endswith('.xml') and ('xl/worksheets/' in filename or filename == 'xl/workbook.xml'):
                print(f" Checking XML file: {filename}")
                try:
                    # --- Step 4: Decode XML content to string ---
                    # Assume UTF-8 encoding, which is standard for OOXML
                    xml_string = file_content.decode('utf-8')
                    # original_xml_string = xml_string # Keep a copy for comparison - not strictly needed for logic

                    # --- Step 5: Use Regex to remove protection tags ---
                    # This regex looks for <sheetProtection ... /> or <sheetProtection>...</sheetProtection>
                    # It handles potential namespaces (like <x:sheetProtection>) and attributes.
                    # It's non-greedy (.+?) to avoid overmatching if multiple tags exist.
                    # It accounts for self-closing tags (<.../>) and tags with content (<...>...</...>).
                    # Using re.IGNORECASE for tag names, though they are typically camelCase.
                    # Added optional namespace prefix handling ([\w]*:)?
                    sheet_prot_pattern = r"<([\w]*:)?sheetProtection.*?(/>|</([\w]*:)?sheetProtection>)"
                    workbook_prot_pattern = r"<([\w]*:)?workbookProtection.*?(/>|</([\w]*:)?workbookProtection>)"

                    # Keep track of whether any changes were made to this specific file
                    file_modified = False

                    # Remove sheetProtection tags
                    modified_xml_string, num_sheet_removed = re.subn(sheet_prot_pattern, '', xml_string, flags=re.IGNORECASE | re.DOTALL)
                    if num_sheet_removed > 0:
                         print(f"  Removed {num_sheet_removed} <sheetProtection> tag(s) from {filename}")
                         removed_tags_count += num_sheet_removed
                         xml_string = modified_xml_string # Update string for next removal
                         file_modified = True

                    # Remove workbookProtection tags
                    modified_xml_string, num_workbook_removed = re.subn(workbook_prot_pattern, '', xml_string, flags=re.IGNORECASE | re.DOTALL)
                    if num_workbook_removed > 0:
                         print(f"  Removed {num_workbook_removed} <workbookProtection> tag(s) from {filename}")
                         removed_tags_count += num_workbook_removed
                         xml_string = modified_xml_string
                         file_modified = True

                    # --- Step 6: Store modified or original content ---
                    if file_modified:
                        # Encode the modified string back to bytes
                        temp_xml_data[filename] = xml_string.encode('utf-8')
                        files_modified_count += 1
                        print(f"  Stored modified version of {filename}")
                    else:
                         # No changes, store original content
                         temp_xml_data[filename] = file_content
                         print(f"  No protection tags found in {filename}. Storing original.")

                except UnicodeDecodeError as e:
                    print(f"  Warning: Could not decode XML file {filename} as UTF-8. Skipping modification. Error: {e}")
                    # Store original content if decoding fails
                    temp_xml_data[filename] = file_content
                except Exception as e:
                    print(f"  Error processing file {filename} with regex: {e}")
                    temp_xml_data[filename] = file_content # Store original on other errors
            else:
                # For non-XML files or irrelevant XMLs, just store original content
                 temp_xml_data[filename] = file_content


        # --- Step 7: Create the output XLSX (zip) file ---
        if files_modified_count > 0:
            print(f"\nWriting output file: {output_xlsx_path} ({files_modified_count} XML file(s) modified)")
        else:
            print("\nNo protection tags found in any relevant XML files. Output file will be a copy of the input.")

        # Ensure the output directory exists
        output_dir = os.path.dirname(output_xlsx_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Created output directory: {output_dir}")


        with zipfile.ZipFile(output_xlsx_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            # Write all files (original or modified) back to the new zip
            for filename, content in temp_xml_data.items():
                # For directories, use writestr with an empty string
                if filename.endswith('/'):
                     zout.writestr(filename, b'')
                else:
                     zout.writestr(filename, content)


        print(f"\nSuccessfully processed and saved to: {output_xlsx_path}")
        print(f"Total protection tags removed: {removed_tags_count}")
        return True

    except Exception as e:
        print(f"An unexpected error occurred during file writing: {e}")
        # Clean up potentially incomplete output file
        if os.path.exists(output_xlsx_path):
             try:
                 os.remove(output_xlsx_path)
                 print(f"Cleaned up incomplete output file: {output_xlsx_path}")
             except OSError as rm_err:
                 print(f"Error trying to clean up output file: {rm_err}")
        return False

# --- Command Line Interface Setup ---
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Removes worksheet and workbook protection from an XLSX file."
    )

    # Add arguments for input and output files
    parser.add_argument(
        "input_file",
        help="Path to the input XLSX file."
    )
    parser.add_argument(
        "output_file",
        help="Path to save the modified XLSX file."
    )

    # Parse the command-line arguments
    args = parser.parse_args()

    # Call the protection removal function with the provided arguments
    if remove_protection_from_xlsx_regex(args.input_file, args.output_file):
        print("\nProtection removal process completed successfully.")
    else:
        print("\nProtection removal process failed.")
