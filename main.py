import json
import pprint

file_path = "dataset/result_11769.json"

with open(file_path, "r") as file:
    data = json.load(file)

# pprint.pprint(data)

data['']