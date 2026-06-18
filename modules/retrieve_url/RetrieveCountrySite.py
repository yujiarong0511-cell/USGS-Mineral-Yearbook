# Input: (1) country_url (str): URL to USGS country homepage; (2) country_name (str): name of the country
    # Variant: (1) country_response: response from requests.get(); (2) country_name (str)
# Output: A list combining two dictionaries whose format is like: {'[country_name]':{'[year1]':'[url1]','[year2]':'[url2]'}}

import requests
from bs4 import BeautifulSoup

# ----------------Integrated Solution----------------
def reports_tables_url(country_url, country_name):
    report_dict = {}
    report_dict[country_name] = {}
    table_dict = {}
    table_dict[country_name] = {}
    
    response = requests.get(url=country_url,
                            headers={
                                'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36 Edg/149.0.0.0"
                            },
                            timeout=10) # DEBUG: Timeout threshold is set at 10 seconds
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.content, features='html.parser')

    reports_html = [para for para in soup.find_all('p') if "Reports:" in para.text] # Given the simple layout of most country homepage, we can directly search for fields embedding "Reports:", in most cases contained in <p> labels
    if len(reports_html) != 0: # Check if reports are available for this country
        reports_html = reports_html[0] # In most cases, all reports of a single country are contained in a single <p> label, thus retain only the first element in the list.
        reports_html_clean = [report for report in reports_html.children if report.name == 'a'] # In the <p> label encapsulating all hyperlinks to reports, URLs and year of the report info are stored in <a> labels.
        for report_label in reports_html_clean: # Retrieve report URL and year info from retained <a> labels.
            year = report_label.get_text(strip=True) # Year of the report
            url = report_label.get('href', default='') # URL of the report
            report_dict[country_name][year] = url # Provisionally save to dictionary
    
    tables_html = [para for para in soup.find_all('p') if "Tables:" in para.text] # Analogous to reports_html
    if len(tables_html) != 0:
        tables_html = tables_html[0]
        tables_html_clean = [table for table in tables_html.children if table.name == 'a']
        for table_label in tables_html_clean:
            year = table_label.get_text(strip=True)
            url = table_label.get('href', default='')
            table_dict[country_name][year] = url
    
    output_list = [report_dict, table_dict] # Concatenate two dictionaries, reports_html and tables_html, to a list for output.
    return output_list

# -----------Variant: Input is replaced with response (to facilitate use of Session Class)--------------
def reports_tables_url_session(country_response, country_name):
    report_dict = {}
    report_dict[country_name] = {}
    table_dict = {}
    table_dict[country_name] = {}
    
    country_response.encoding = 'utf-8'
    soup = BeautifulSoup(country_response.content, features='html.parser')

    reports_html = [para for para in soup.find_all('p') if "Reports:" in para.text] # Given the simple layout of most country homepage, we can directly search for fields embedding "Reports:", in most cases contained in <p> labels
    if len(reports_html) != 0: # Check if reports are available for this country
        reports_html = reports_html[0] # In most cases, all reports of a single country are contained in a single <p> label, thus retain only the first element in the list.
        reports_html_clean = [report for report in reports_html.children if report.name == 'a'] # In the <p> label encapsulating all hyperlinks to reports, URLs and year of the report info are stored in <a> labels.
        for report_label in reports_html_clean: # Retrieve report URL and year info from retained <a> labels.
            year = report_label.get_text(strip=True) # Year of the report
            url = report_label.get('href', default='') # URL of the report
            report_dict[country_name][year] = url # Provisionally save to dictionary
    
    tables_html = [para for para in soup.find_all('p') if "Tables:" in para.text] # Analogous to reports_html
    if len(tables_html) != 0:
        tables_html = tables_html[0]
        tables_html_clean = [table for table in tables_html.children if table.name == 'a']
        for table_label in tables_html_clean:
            year = table_label.get_text(strip=True)
            url = table_label.get('href', default='')
            table_dict[country_name][year] = url
    
    output_list = [report_dict, table_dict] # Concatenate two dictionaries, reports_html and tables_html, to a list for output.
    return output_list


# --------------------------------------------------
def reports_url(country_url, country_name):
    report_dict = {}
    report_dict[country_name] = {}
    
    response = requests.get(url=country_url,
                            headers={
                                'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36 Edg/149.0.0.0"
                            })
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.content, features='html.parser')

    reports_html = [para for para in soup.find_all('p') if "Reports:" in para.text]
    if len(reports_html) != 0:
        reports_html = reports_html[0]
        reports_html_clean = [report for report in reports_html.children if report.name == 'a']
        for report_label in reports_html_clean:
            year = report_label.get_text(strip=True)
            url = report_label.get('href', default='')
            report_dict[country_name][year] = url
    
    return report_dict


def tables_url(country_url, country_name):
    table_dict = {}
    table_dict[country_name] = {}
    
    response = requests.get(url=country_url,
                            headers={
                                'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36 Edg/149.0.0.0"
                            })
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.content, features='html.parser')

    tables_html = [para for para in soup.find_all('p') if "Tables:" in para.text]
    if len(tables_html) != 0:
        tables_html = tables_html[0]
        tables_html_clean = [table for table in tables_html.children if table.name == 'a']
        for table_label in tables_html_clean:
            year = table_label.get_text(strip=True)
            url = table_label.get('href', default='')
            table_dict[country_name][year] = url
    
    return table_dict