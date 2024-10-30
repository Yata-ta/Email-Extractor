import sys
import json

def load_json(json_file_path):
    """
    Load the contents of the given JSON file into a dictionary.
    
    :param json_file_path: The path to the JSON file.
    :return: A dictionary with the JSON contents.
    """
    try:
        with open(json_file_path, 'r') as json_file:
            data = json.load(json_file)
            return data
    except FileNotFoundError:
        print(f"Error: The file '{json_file_path}' was not found.")
        sys.exit(1)
    except json.JSONDecodeError:
        print(f"Error: The file '{json_file_path}' is not a valid JSON file.")
        sys.exit(1)

def save_json(json_file_path, data):
    """
    Save the dictionary contents back into the JSON file.
    
    :param json_file_path: The path to the JSON file.
    :param data: The dictionary to be saved.
    """
    try:
        with open(json_file_path, 'w') as json_file:
            json.dump(data, json_file, indent=4)
            print(f"Updated JSON saved to '{json_file_path}'")
    except Exception as e:
        print(f"Error: Unable to save the file '{json_file_path}'. Reason: {e}")
        sys.exit(1)

def email_analysis(data):
    """
    Check if the key 'Processed' exists in the dictionary, and set its value to True.
    If the key doesn't exist, it will be created with the value True.
    
    :param data: The dictionary containing JSON data.
    :return: The updated dictionary.
    """


    ## Process Here




    if 'Processed' not in data:
        print("'Processed' key does not exist. Adding it.")
    else:
        print("'Processed' key exists. Updating its value.")
    
    data['Processed'] = True
    return data

def main():
    # Check if the file path is provided as a command-line argument
    if len(sys.argv) != 2:
        print("Usage: python file.py <file.json>")
        sys.exit(1)

    json_file_path = sys.argv[1]
    
    # Load JSON file contents into a dictionary
    data_dict = load_json(json_file_path)
    
    # Update the 'Processed' key in the dictionary
    data_dict = email_analysis(data_dict)
    
    # Save the updated dictionary back to the JSON file
    save_json(json_file_path, data_dict)

if __name__ == "__main__":
    main()
