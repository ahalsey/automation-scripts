# updates PO lines via the Coupa API
# re-opens lines before update if they are closed
# closes lines post-update whether update is successful or fails

import pandas as pd
import requests
import json
import sys

def get_params():
    params = []
    runtime_env = False
    # configure API URL based on runtime environment
    while runtime_env is False:
        runtime_env = input("Enter runtime environment (Test, Dev, or Prod): ").upper()
        if runtime_env == "PROD":
            URL = 'https://coupahost.com'
        elif runtime_env == 'TEST':
            URL = 'https://test.coupahost.com'
        elif runtime_env == 'DEV':
            URL = 'https://dev.coupahost.com'
        else:
            #  reset runtime environment until valid input
            runtime_env = False
            print("Invalid runtime environment!\n")

    params.append(URL)
    params.append(runtime_env)
    # assign file variables based on input
    params.append(input("Enter filename: "))
    params.append(input("Enter sheet name: "))
    params.append(input("Enter API Key: "))
    params.append('updated_po_lines.csv')

    return params


# XML template for putting new PO line data all at once at the PO header level
order_header_template = """<?xml version="1.0" encoding="UTF-8"?>
<order-header>
   <order-lines>
    {all_lines}
   </order-lines>
</order-header>"""

order_line_template = """<order-line>
         <id>{line_id}</id>
         <account>
            <code>{account_code}</code>
            <account-type>
               <name>{chart_of_accounts}</name>
            </account-type>
            <segment-1>{segment_1}</segment-1>
            <segment-2>{segment_2}</segment-2>
            <segment-3>{segment_3}</segment-3>
            <segment-4>{segment_4}</segment-4>
            <segment-5>{segment_5}</segment-5>
         </account>
      </order-line>"""

line_status_template = """<?xml version="1.0" encoding="UTF-8"?>
<order-header>
    <order-lines>
        <order-line>
            <id>{line_id}</id>
            <status>{line_status}</status>
        </order-line>
    </order-lines>
</order-header>"""


def put_request_update_po_ln(po_id, xml, session):
    # Puts the requested XML into a PO ID
    headers = {"accept": "application/xml", "X-COUPA-API-KEY": str(api_key)}
    query_url = '%s/api/purchase_orders/%s' % (URL, po_id)
    response = session.put(query_url, headers=headers, data=xml)
    return response


def get_po_data(po_id, session):
    # Gets the data based on PO ID
    headers = {"accept": "application/json", "X-COUPA-API-KEY": str(api_key)}
    query_url = '%s/api/purchase_orders/%s' % (URL, po_id)
    print("Getting PO Data for PO #%s...\n" % po_id, end="")
    response = session.get(query_url, headers=headers)
    if response.status_code == 200:
        j = json.loads(response.content)
        return j
    else:
        print("Error loading PO data!")

def get_po_line_id(po_data, line_num):
    # Searches through PO data to get the po_line_id of the requested PO line
    for line in po_data["order-lines"]:
        if line["line-num"] == str(line_num):
            return line["id"], line["status"]

def open_line(po_line_id, session):
    # re-opens a PO line if it is closed
    headers = {"accept": "application/json", "X-COUPA-API-KEY": str(api_key)}
    query_url = '%s/api/purchase_order_lines/%s/reopen_for_receiving' % (URL, po_line_id)
    json = {"reason-insight-code": "API"}
    response = session.put(query_url, headers=headers, json=json)

def close_line(po_id, line_ids, session):
    # soft-closes PO lines
    for line_id in line_ids:
        headers = {"accept": "application/xml", "X-COUPA-API-KEY": str(api_key)}
        query_url = '%s/api/purchase_orders/%s' % (URL, po_id)
        payload = {
                "line_id": line_id,
                "line_status": "soft_closed_for_invoicing"
            }
        put_request = line_status_template.format(**payload)
        # Puts the requested XML into a PO ID
        response = session.put(query_url, headers=headers, data=put_request)
        print(response)
        if response.status_code == 200:
            print("Line number %s of PO #%s closed successfully.\n" % (line_id, po_id))
            f.write("%s,N/A,%s,SUCCESS,Line closed successfully\n" % (po_id, line_ids.index(line_id) + 1))
        else:
            print("Line number %s of PO #%s NOT closed\n" % (line_id, po_id))
            f.write("%s,N/A,%s,FAILURE,Line NOT closed successfully\n" % (po_id, line_ids.index(line_id) + 1))
    
# Main method
if __name__ == "__main__":
    # gather parameters
    parameters = get_params()
    URL, runtime_env, file_name, sheet_name, api_key, logfile = parameters[0], parameters[1], parameters[2], \
                                                                parameters[3], parameters[4], parameters[5]

    # Read in data from XLSX file
    try:
        df = pd.read_excel(open(file_name, 'rb'),sheet_name=sheet_name, keep_default_na=False)
    except:
        print("File not found... exiting.")
        sys.exit(1)

    try:
        f = open(logfile, "a", encoding='utf-8')
    except:
        print("Unable to open %s... exiting." % logfile)
        sys.exit(1)

    # loop through each PO individually, one get and one put per PO (multiple lines)
    # Coupa requires all COAs to match, all PO lines must be updated at once
    f.write("po_id,total_lines,line_id,status,response\n")
    with requests.Session() as session:
        for po_id in df.po_id.unique():
            po_df = df.loc[df.po_id == int(po_id), ]

            # Create session object to use with all requests
            
            po_data = get_po_data(po_id, session)

            line_list = []
            line_ids = []

            # for each line of the PO, generate an order line object in XML
            for index, row in po_df.iterrows():
                try:
                    order_line_id, order_line_status = get_po_line_id(po_data, row['line_num'])
                except:
                    break

                if order_line_status == "soft_closed_for_invoicing":
                    open_line(order_line_id, session)

                if row['segment_1'] > 9:
                    segment_1 = row['segment_1']
                else: # add a '0' before the number if it's a single digit
                    segment_1 = "0" + str(row['segment_1'])

                    fields = {
                            "line_id": order_line_id,
                            "account_code": row['account_code'],
                            "chart_of_accounts": row['chart_of_accounts'],
                            "segment_1": segment_1,
                            "segment_2": row['segment_2'],
                            "segment_3": row['segment_3'],
                            "segment_4": row['segment_4'],
                            "segment_5": row['segment_5']
                        }

                output_line = order_line_template.format(**fields)
                line_list.append(output_line)  # build out a list of order lines and combine them later
                line_ids.append(order_line_id)

            # put the lines together for a single put request
            print("Putting data for %d lines...\n" % len(line_list), end="")
            full_put = order_header_template.format(all_lines=' '.join(line_list))
            r = put_request_update_po_ln(po_id, full_put, session)
            if r.status_code == 200:
                print("SUCCESS!")
                f.write("%s,%s,N/A,SUCCESS\n" % (po_id, len(line_list)))
                close_line(po_id, line_ids, session)
            else:
                # logs error and closes the PO lines
                f.write("%s,%s,N/A,FAILURE, %s, \n" % (po_id, len(line_list), r.text.split("\n")[3].strip()))
                close_line(po_id, line_ids, session)
            

f.close()
