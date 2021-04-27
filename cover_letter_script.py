from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import pandas as pd
from docx2pdf import convert

template = "template.docx"
cl_text = pd.read_csv("cover_letter_text.csv")

job_lst = []
for line_num in range(cl_text.shape[0]):
    job_dict = {}
    for col in cl_text.columns:
        job_dict[col] = cl_text.loc[line_num, col]
    job_lst.append(job_dict)

for job in job_lst:
    document = MailMerge(template)
    document.merge(
                  company_name=job['company_name'],
                  position_name=job['position_name'],
                  job_posting_source=job['job_posting_source'],
                  personalized_introduction=job['personalized_introduction'],
                  personalized_industry_comment=job['personalized_industry_comment']
                    )
    document.write(f'./Name_Cover_Letter_{job['company_name']}.docx')
document.close()

