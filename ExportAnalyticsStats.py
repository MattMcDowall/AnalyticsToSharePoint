# Initialize values, etc.
import pandas as pd
import requests
import shutil
import xmltodict
import Credentials      # Local file with API keys, etc.

# # # Change these to your paths, etc. # # #
apikey = Credentials.analytics_api  # Set in Credentials.py (within this folder)
report = '%2Fshared%2FUniversity%20of%20Nebraska%20Kearney%2001UON_K%2FReports%2FMass%20weeding%20project%20-%20withdrawn%20items'
sp_folder = '../Collection Development - Reports/'
excel_file = 'Mass weeding project - withdrawn items.xlsx'

baseurl = 'https://api-na.hosted.exlibrisgroup.com'
query = '/almaws/v1/analytics/reports?path={report}&limit=1000&col_names=true&token={token}&apikey={apikey}'

df = pd.DataFrame()
records_list = []
cols = {}
token = ''
query_num = 0

# Get records from Analytics (into a dataframe)
while True:
    query_num += 1
    record_num = 0
    query_msg = "Querying . . . " if query_num == 1 else "Continuing: Query no. " + str(query_num)
    print(query_msg)
    # Query the API
    r = requests.get(''.join([baseurl, query.format(report=report, token=token, apikey=apikey)]))
    # Check whether response was good
    if r.status_code != 200:
        exit_msg = 'Failed querying the API: ' + str(r.status_code) + ' [' + r.reason + ']\n\t\tTo wit: ' + xmltodict.parse(r.text)['web_service_result']['errorList']['error']['errorMessage']
        raise SystemExit(exit_msg)
    rdict = xmltodict.parse(r.content)
    # If there's a resumption token, store it
    token = rdict['report']['QueryResult']['ResumptionToken'] if 'ResumptionToken' in rdict['report']['QueryResult'] else token

    # If there's table schema (i.e., if this is the first response), use it to create the dataframe
    if 'xsd:schema' in rdict['report']['QueryResult']['ResultXml']['rowset']:
        print('Creating header row')
        headerset = rdict['report']['QueryResult']['ResultXml']['rowset']['xsd:schema']['xsd:complexType']['xsd:sequence']['xsd:element']

        # For each column in the response . . .
        colrange = range(1, len(headerset) - 1)  # First column is empty; last column is just "total" count
        for i in colrange:
            col_head = headerset[i]['@saw-sql:columnHeading']
            # . . . create a dict matching generic name with actual heading name . . .
            cols[headerset[i]['@name']] = col_head
            # . . . and populate the df headers
            df[col_head] = None

    # Read regular rows & add them to the df
    for row in rdict['report']['QueryResult']['ResultXml']['rowset']['Row']:
        record_num += 1
        if record_num % 100 == 0:
            print(record_num, end=' . . . ')
        df_row = {}
        for cell in [c for c in row if (c != 'Column0' and c != 'Column31')]:
            df_row.update({cols[cell]: row[cell]})
        records_list.append(df_row)

    # Stop querying if this is the last response
    if rdict['report']['QueryResult']['IsFinished'] == 'true':
        break

    print()
print()

# Create the actual df
print('Creating dataframe.')
df = pd.DataFrame.from_dict(records_list)

# Create the Excel file
# Works best to create the file OUTSIDE SharePoint . . .
print('Exporting to Excel file.')
df.sort_values(by='Withdrawal Date', ascending=False).to_excel(excel_file, index=False)
# . . . then drop it into the synced folder
print('Transferring to SharePoint.')
shutil.move(excel_file, sp_folder + excel_file)
