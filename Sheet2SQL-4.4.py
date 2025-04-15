import pandas as pd
import os

# Function to generate an SQL script from an Excel file
def generate_sql(input_excel, output_sql, table_name):
    try:
        # Read the Excel file into a DataFrame
        print(f"Reading the file: {input_excel}")
        df = pd.read_excel(input_excel, dtype=str, keep_default_na=False)
        print(f"File read successfully, processing rows...")

        # Get column headers and replace spaces with underscores
        headers = df.columns
        headers = [header.replace(" ", "_") for header in headers]

        # Start building the CREATE TABLE SQL statement
        create_table_sql = f"CREATE TABLE {table_name} (\n"
        for header in headers:
            create_table_sql += f"    [{header}] varchar(255),\n"
        create_table_sql = create_table_sql.rstrip(",\n") + "\n);\n"

        # Add optional SQL commands for debugging or cleanup
        create_table_sql += f"\n-- SELECT * FROM {table_name};\n"
        create_table_sql += f"-- DROP TABLE {table_name};\n"

        # Start building the INSERT INTO SQL statement
        insert_columns = f"INSERT INTO {table_name} (\n"
        insert_columns += ",\n".join([f"   [{header}]" for header in headers])
        insert_columns += "\n) VALUES\n"

        # Prepare to batch insert values
        insert_values = []
        batch_size = 999  # Maximum number of rows per batch
        batch_count = 0

        # Iterate through each row in the DataFrame
        for i, row in enumerate(df.itertuples(index=False, name=None)):
            # Escape single quotes and backslashes in cell values
            values = ",".join([f"'{str(cell).replace('\'', '\'\'').replace('\\', '/')}'" for cell in row])
            insert_values.append(f"({values})")

            # Write SQL script in batches to avoid exceeding SQL limits
            if (i + 1) % batch_size == 0 or (i + 1) == len(df):
                if batch_count == 0:
                    sql_script = create_table_sql + "\n" + insert_columns + ",\n".join(insert_values) + ";\n"
                else:
                    sql_script += insert_columns + ",\n".join(insert_values) + ";\n"
                
                insert_values = []  # Clear the batch
                batch_count += 1

        # Save the generated SQL script to the output file
        output_file_path = os.path.join(os.path.dirname(input_excel), output_sql)
        with open(output_file_path, 'w', encoding='utf-8') as sqlfile:
            sqlfile.write(sql_script)

        print(f"SQL script created successfully! Output saved to {output_file_path}")

    except Exception as e:
        # Handle any errors that occur during processing
        print(f"An error occurred while processing the file: {e}")
        input("Press Enter to exit...")  

# Main function to handle user input and execute the script
def main():
    # Get the current directory of the script
    current_directory = os.path.dirname(os.path.abspath(__file__))
    print(f"Current directory: {current_directory}")  

    while True:
        try:
            # Prompt the user for the Excel file name
            input_excel = input("Enter the name of the Excel file (without .xlsx) or type 'exit' to quit: ").strip()

            # Exit the program if the user types 'exit'
            if input_excel.lower() == 'exit':
                print(r"""
  *****************************************
  *                                       
  *   This program was written with ❤️    
  *      by   lucafungo   AKA baffon          
  *      AKA   Guybrush Threepwood          
  *      A Pirate!     🏴‍☠️ AHRRR!       
  *                                       
  *         Goodbye, FNS!                
  *****************************************
                """)
                break

            # Append '.xlsx' if not already included in the file name
            if not input_excel.endswith('.xlsx'):
                input_excel += '.xlsx'
                

            # Construct the full file path
            file_path = os.path.join(current_directory, input_excel)
            print(f"Looking for file: {file_path}")

            # Check if the file exists in the current directory
            if not os.path.exists(file_path):
                print(f"File not found in '{current_directory}'. Please ensure the file is in the same directory.")
                continue

            # Ask the user if the table is temporary and get the table name
            is_temp_table = input("Is this a temporary table? (yes/no): ").strip().lower()
            table_name = input("Enter the table name (without #, e.g., 'my_table'): ").strip()

            # Prefix the table name with '#' if it's a temporary table
            if is_temp_table == 'yes':
                table_name = f"#{table_name}"

            # Generate the output SQL file name
            base_name = os.path.splitext(input_excel)[0]
            output_sql = f"{base_name}_output.sql"

            # Call the function to generate the SQL script
            generate_sql(file_path, output_sql, table_name)

        except Exception as e:
            # Handle unexpected errors
            print(f"An unexpected error occurred: {e}")
            input("Press Enter to continue...")

# Entry point of the script
if __name__ == "__main__":
    main()
    input("Press Enter to exit...")