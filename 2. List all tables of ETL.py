import cx_Oracle
import re
import pandas as pd
import winsound

# Prompt for password
password = input("Enter password for VIBE_ETLADMIN: ")

# Connect to the database
connection = cx_Oracle.connect(user="VIBE_ETLADMIN", password=password,
                               dsn=cx_Oracle.makedsn("10.80.100.103", 8004, sid="devobidw"))
cursor = connection.cursor()
print("Connected to server...")


# Get the code of the procedure we are parsing
def get_procedure_code(procedure_name):
    user, name = procedure_name.split('.')
    query = "SELECT TEXT FROM ALL_SOURCE WHERE OWNER = :owner AND NAME = :name ORDER BY LINE"
    cursor.execute(query, owner=user.upper(), name=name.upper())
    code = cursor.fetchall()
    return ''.join([row[0] for row in code])


# Check if a given name is a view or a table
def check_if_view_or_table(user, name):
    query = "SELECT OBJECT_TYPE FROM ALL_OBJECTS WHERE OWNER = :owner AND OBJECT_NAME = :name"
    cursor.execute(query, owner=user.upper(), name=name.upper())
    result = cursor.fetchone()
    return result[0] if result else 'UNKNOWN'


# Modified parser program
def parse_procedure_code(code):
    procedures = re.findall(r"(\w+)\.(\w+)(\(|;)", code, re.IGNORECASE)
    tables = re.findall(r"(INSERT INTO|MERGE INTO|FROM|JOIN) (\w+)\.(\w+)", code, re.IGNORECASE)
    return procedures, tables


# Modified append_data function
def append_data(step, call_stack, origin_user, origin_name, entity_type, user, name, data_list):
    object_type = check_if_view_or_table(user, name)
    data_list.append([step, call_stack, origin_user, origin_name, f"{entity_type} ({object_type})", user, name])


# Recursive procedure parser
def recursive_procedure_parser(proc_name, step, call_stack, data_list):
    procedure_code = get_procedure_code(proc_name)
    procedures, tables = parse_procedure_code(procedure_code)

    print("Working with procedures:")
    print(procedures)

    for index, proc in enumerate(procedures, start=1):
        new_step = step + [index]
        new_call_stack = call_stack + [f"{proc[0]}.{proc[1]}"]
        recursive_procedure_parser(f"{proc[0]}.{proc[1]}", new_step, new_call_stack, data_list)

    print("Working with tables:")
    print(tables)

    for table in tables:
        append_data(step, call_stack, proc_name.split('.')[0], proc_name.split('.')[1], table[0].lower(), table[1],
                    table[2], data_list)


# Initiating the process
def main():
    data_to_append = []
    input_procedure = input("Enter procedure in format 'user.procedure': ")
    file_path = input("Enter the directory where you want to save the Excel file: ")
    print("Parsing begun")
    recursive_procedure_parser(input_procedure, [1], [input_procedure], data_to_append)
    save_to_excel(data_to_append, input_procedure, file_path)


# Save to excel
def save_to_excel(data_list, input_procedure, file_path):
    df = pd.DataFrame(data_list, columns=['step', 'call_stack', 'origin_user', 'origin_name', 'type', 'user', 'name'])
    full_path = f"{file_path}\\{input_procedure}_explored.xlsx"
    df.to_excel(full_path, index=False, sheet_name="Sheet1")
    print(f"Successfully saved to: {file_path} As: {input_procedure}")

# Start the process
if __name__ == "__main__":
    main()

# Closing statements
cursor.close()
connection.close()
winsound.Beep(1000, 500)  # 1000 Hz frequency, 500 ms duration
