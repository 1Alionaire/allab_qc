import json
from pathlib import Path
import math

analyst = 'VC'
types_analysis = ["plm_count", "nob_count","tem_count",]


path = Path(fr"C:\Python\allab\allab_qc\data\{analyst}_data.jsonl")

output_file = Path(r"C:\Python\allab\allab_qc\data\VC_NEW_data.jsonl")

grouped_date_file = Path(r"C:\Python\allab\allab_qc\data\analysts_grouped.json")

final_result = {}

with path.open("r", encoding="utf-8") as f:
    data = json.load(f)       

with grouped_date_file.open("r", encoding="utf-8") as f:
    grouped_data_analysts = json.load(f) 

grouped_data_analysts = grouped_data_analysts[analyst]


for record in grouped_data_analysts:
    # how many blank samples per day accroding audit
    limit = 100
    # total count samples per day. 
    plm_total_samples = 0
    # filter all projects by day
    projects_by_date = {key: value for key, value in data.items() if value['date'] == record['date'] }
    # this counter need only for first blank sample in first project of the day
    counter = 0

    #iterate through all projects in this day
    for project, project_info in projects_by_date.items():
        # for type_analysis in types_analysis:
        if project_info["plm_count"] > 0:
            # sum samples before adding samples from current project
            before = plm_total_samples
            # sum samples after adding samples from current project
            after = before + project_info["plm_count"]
            # divide by samples amount how many samples before and after 
            blanks_before = before // limit
            blanks_after = after // limit
            # how many blanks? difference between before and after
            blanks_to_add = blanks_after - blanks_before
            # if first project of the day - it must be one blank
            if counter == 0:
                blanks_to_add = 1
            #add result
            final_result[project] = {'plm_blank':blanks_to_add}
            # new total amount samples per day
            plm_total_samples = after
            # increase counter for day
            counter+=1

print(final_result)