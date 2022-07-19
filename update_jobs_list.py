from os import read
from shareplum import Office365
from shareplum import Site
from shareplum.site import Version
import shareplum as sp
import pandas as pds

#shareplum code
authcookie = Office365('sharepoint site name', username='xxxxxx', password='xxxxxx').GetCookies()
site = Site('specific sharepoint site name', version=Version.v365, authcookie=authcookie)
active_jobs_list = site.List('sharepoint list name')
ajl_items = active_jobs_list.GetListItems()
ajl_jobs = active_jobs_list.GetListItems(fields=['Job Number'])
ajl_data = []
for item in ajl_jobs:
    ajl_data.append(item['Job Number'])


#pandas code
file = ('file location')
excel_data = pds.read_excel(file)
excel_jobs = []
for item in excel_data['Job']:
    excel_jobs.append(item)

#script
def get_differences(ajl_data, excel_data):
    """Gets the differences between the current active jobs list and the
       new active jobs list
    """
    to_append = []
    to_delete = []
    for job in excel_data['Job']:
        if job not in ajl_data:
            to_append.append(job)
    for job_num in ajl_data:
        if job_num not in excel_jobs:
            to_delete.append(job_num)
    return to_append, to_delete

def deleted(to_delete):
    output = []
    for dicti in ajl_items:
        if dicti['Job'] in to_delete:
            output.append(dicti['ID'])
    return output

def get_new_item_info(to_append):
    output = []
    for i in range(len(excel_data)):
        if excel_data.loc[i]['Job'] in to_append:
            output.append(excel_data.loc[i])
    return output

def format_items(new_items):
    output = []
    for item in new_items:
        to_add = dict()
        to_add['Job Number'] = item['Job']
        to_add['Customer'] = item['Customer']
        to_add['Job Name'] = item['Job Name']
        to_add['Job Lead'] = item['Job Lead']
        to_add['Budget'] = item['Budget']
        to_add['Category'] = item['Category']
        output.append(to_add)
    return output

def __main__(ajl_data, excel_data):
    to_append, to_delete = get_differences(ajl_data, excel_data)
    deleted_ids = deleted(to_delete)
    if len(deleted_ids) > 0:
        print('These jobs have been removed:')
        for item in to_delete:
            print(item)
    else:
        print('No jobs have been closed')
    new_items = get_new_item_info(to_append)
    ready_to_add = format_items(new_items)
    if len(ready_to_add) > 0:
        print('These jobs have been added:')
        for item in to_append:
            print(item)
    else:
        print('No new jobs')

__main__(ajl_data, excel_data)
