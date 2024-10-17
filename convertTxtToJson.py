import json
import ast

# Read the 'clients_list.txt' file
with open('clients_list.txt', 'r') as file:
    lines = file.readlines()

# Convert to a dictionary
try:
    # Use ast.literal_eval to safely evaluate the string as a Python object
    dictionary_list = ast.literal_eval(lines[0])
except (ValueError, SyntaxError) as e:
    print("Error parsing string:", e)

# Convert the list of dictionaries to JSON format if needed
json_output = json.dumps(dictionary_list, indent=4)

# Output the resulting dictionary
print(json_output)


# Write the list of dictionaries as a JSON file
with open('output.json', 'w') as json_file:
    json.dump(json.loads(json_output), json_file, indent=4)

print("JSON file with list of dictionaries has been generated successfully.")
