class Material(object):
    def __init__(self, list_info):
        '''
        Assuming that list_info is fed from parse_SAP_export
        '''
        self.mat_num = list_info[2]
        self.mat_type = list_info[3]
        self.desc = list_info[5]
        self.basic_mat = list_info[6]
        self.amc = list_info[9]


def create_materials_from_SAP_csv(file_path):
    with open(file_path, 'r') as f:
        list_rows = f.readlines()

    list_obj = []
    # Skip header rows
    for i in range(1, len(list_rows)):
        row = list_rows[i]
        list_obj.append(Material(row.split(',')))

    return list_obj
