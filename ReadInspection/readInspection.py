import re
import os

from msOfficeOperations import docC2txt
from models import Anchor

file_path = os.path.join(os.getcwd(), '50851653.doc')
txt_file_path = docC2txt(file_path)

# Read txt file
content = []
with open(txt_file_path, 'r') as f:
    content = f.readlines()


i = 0
snippet = []
anchors = []  # List of anchor objects
anchor_pattern = re.compile(r"^([A-Z]\)|\d+\))")

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
important_anchors = "A,B,C,D,E,F,G,H,12,J,K,L,M,N,P,21,R,S,T,U,V,W,X,Y,Z".split(',')

csv_file_path = file_path.rstrip('doc') + 'csv'
with open(csv_file_path, 'w+') as f:
    f.write("Item, Min, Max\n")
    for anchor in anchors:

        write_anchor = False
        for important_anchor in important_anchors:
            if anchor.name[:-1] == important_anchor:
                write_anchor = True
        if write_anchor:
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
