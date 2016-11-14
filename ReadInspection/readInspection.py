import re
import os

from msOfficeOperations import docC2txt
from models import Anchor

file_name = input("Enter a file name:\n")
file_path = os.path.join(os.getcwd(), file_name)
txt_file_path = docC2txt(file_path)

# Read txt file
content = []
with open(txt_file_path, 'r') as f:
    content = f.readlines()


i = 0
snippet = []
anchors = []  # List of anchor objects
anchor_pattern = re.compile(r"^([A-Z]{1,2}\)|\d{1,2}\))")

# While loop for finding all anchor positions
while i < len(content):
    line = content[i]
    mo = re.search(anchor_pattern, line)
    if mo is not None:
        anchors.append(Anchor(mo.group(0), i))
    i += 1

# Read text between anchor positions - ignoring last one (Since it is useless)
j = 0
while j < len(anchors) - 1:
    lines = []
    for k in range(anchors[j].line_index, anchors[j + 1].line_index):
        lines.append(content[k])
    anchors[j].add_text_snippet(lines)
    j += 1

# Write CSV file
csv_file_path = file_path.rstrip('doc') + 'csv'
with open(csv_file_path, 'w+') as f:
    f.write("Item, Min, Max\n")
    for anchor in anchors:
        if not anchor.isEmpty:
            f.write("{0},{1},{2}".format(
                anchor.name, anchor.min_num, anchor.max_num
            ))

            try:
                f.write(",")
                for item in anchor.list_nums:
                    f.write("," + str(item))
            except AttributeError:
                f.write(",No numbers found")
            f.write("\n")
