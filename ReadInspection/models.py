import re
from fractions import Fraction

class Anchor(object):
    pattern_decimal = re.compile(r"(\d+\.\d+)")
    pattern_fraction2 = re.compile(r"(?<![\/\d])(\d+)\/(\d+)(?![\/\d])") #From this http://stackoverflow.com/questions/1912376/regular-expression-to-match-fractions-and-not-dates
    pattern_fraction = re.compile(r"(\d*\s\d+/\d+)")

    def __init__(self, name, line_index):
        self.name = name
        self.line_index = line_index

        self.min_num = "N/A"
        self.max_num = "N/A"

        self.isEmpty = False

    def add_text_snippet(self, lines_text):
        self.lines_text = lines_text
        string_text = "\n".join(self.lines_text)

        self.list_nums = re.findall(self.pattern_decimal, string_text)
        self.list_nums= [float(x) for x in self.list_nums]

        fractions = re.findall(self.pattern_fraction, string_text)
        for fraction in fractions:
            self.list_nums.append(float(sum(Fraction(s) for s in fraction.split()))) #From this http://stackoverflow.com/questions/1806278/convert-fraction-to-float
        try:
            self.min_num = min(self.list_nums)
            self.max_num = max(self.list_nums)
        except ValueError:
            self.min_num = "N/A"
            self.max_num = "N/A"
            self.isEmpty = True
        except TypeError:
            self.min_num = "Type Error"
            self.max_num = "Type Error"