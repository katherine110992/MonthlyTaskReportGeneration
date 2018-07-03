import os

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
print(ROOT_DIR)

path = 'C:\\Users\\yespindola\\Documents\\MonthlyTaskReportGeneration\\change_name\\'

for filename in os.listdir(path):
    fields = filename.split(' - ')
    date_time = fields[0].split('_')
    name = fields[1]
    date = date_time[0]
    time = date_time[1]
    new_date = date.replace('-', '_')
    new_time = time.replace('-', '_')
    new_filename = new_date + '-' + new_time + ' - ' + name
    print(new_filename)
    os.rename(path+filename, path+new_filename)

