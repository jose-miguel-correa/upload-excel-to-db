import os
import importlib
from file_function_mapping import mapping

def run_function_on_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file in mapping:
                file_path = os.path.join(root, file)
                function_name = mapping[file]
                module_name = "functions.%s" % function_name  # Using string formatting
                module = importlib.import_module(module_name)
                process_function = getattr(module, "process_%s" % function_name)
                process_function(file_path)

if __name__ == "__main__":
    folder_path = "./files"  # Replace with the path to your folder
    run_function_on_files(folder_path)
