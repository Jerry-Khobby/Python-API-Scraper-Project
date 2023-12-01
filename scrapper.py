import requests
import xlwt 
from xlwt import Workbook 

API_URL = "https://remoteok.com/api"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
REQUEST_HEADER = {
    "User-Agent": USER_AGENT,
    "Accept-Language": "en-US,en;q=0.5",
}

def get_remote_ok_jobs():
    res = requests.get(url=API_URL, headers=REQUEST_HEADER)
    return res.json()

def output_jobs_to_xls(data):
    wb = Workbook()
    job_sheet = wb.add_sheet("Jobs information")

    headers = list(data[0].keys())
    for i in range(0, len(headers)):
        job_sheet.write(0, i, headers[i])

    for i in range(0, len(data)):
        job = data[i]
        values = list(job.values())
        for x in range(0, len(values)):
            job_sheet.write(i + 1, x, values[x])

    wb.save("Remote_jobs.xls")

if __name__ == "__main__":
    json_data = get_remote_ok_jobs()[1:]
    output_jobs_to_xls(json_data)
