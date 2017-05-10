__author__ = 'dbergert'

import argparse
import requests
import xml.etree.ElementTree as etree
import xlsxwriter


username = "Nsatest_api"
password = "9LqADFYYCCl4"


def main():
    
    print ("Getting Last of Applications: ")
    xmlApplist = getApplist()
    applist = etree.fromstring(xmlApplist)
    
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('SampleReport.xlsx')
    worksheet = workbook.add_worksheet()

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': 1})

    # Adjust the column width.
    worksheet.set_column(0,1,10)
    worksheet.set_column(2,3,30)
    worksheet.set_column(4,4,17)
    worksheet.set_column(5,6,30)
    worksheet.set_column(7,7,20)

    # Write some data headers.
    worksheet.write('A1', 'App_ID' ,bold )
    worksheet.write('B1', 'Build_Id' ,bold )
    worksheet.write('C1', 'App_Name',bold )
    worksheet.write('D1', 'Policy_Name',bold )
    worksheet.write('E1', 'Policy_Version',bold )
    worksheet.write('F1', 'Submitted_Date',bold )
    worksheet.write('G1', 'Published_Date',bold )
    worksheet.write('H1', 'Policy_Compliance_Status',bold )
    worksheet.write('I1', 'Rating' ,bold)
    worksheet.write('J1', 'Score' ,bold)

    row = 1
    col = 0 
    
    for appnode in applist:
        print("print appnode app_id")
        print(appnode.get('app_id'))    
        print ("Getting Last Build ID of Application: ")
        xmlBuildList = getBuildList(appnode.get('app_id'))
        root = etree.fromstring(xmlBuildList)
        print ("The last build id is :")
        print (root[-1].get('build_id'))
        print ("extracting build results below:")
        
        xmlBuildInfo = getBuildInfo(appnode.get('app_id'))
        root = etree.fromstring(xmlBuildInfo)

        app_id = ""
        build_id = ""
        app_name = ""
        policy_name = ""
        policy_version = ""
        submitted_date = ""
        published_date = ""
        policy_compliance_status = ""
        #build_status = ""
        rating = ""
        score = ""
        
        
        if (root[0].get('results_ready') == 'false'):
            print('result Ready:' + root[0].get('results_ready') + 'App_id : ' + root.get('app_id') +', Build_id : ' + root.get('build_id') + ',status : ' + root[0][0].get("status"))
            
            """app_id = root.get('app_id')
            build_id = root.get('build_id')
            app_name = appnode.get('app_name')
            policy_name = root[0].get('policy_name')
            policy_version = root[0].get('policy_version')
            submitted_date = "NA"
            published_date = "NA"
            policy_compliance_status = root[0][0].get('status')
            rating = "NA"
            score = "NA" """

            worksheet.write_string  (row, col , root.get('app_id'))
            worksheet.write_string  (row, col+1 , root.get('build_id'))
            worksheet.write_string  (row, col+2 , appnode.get('app_name'))
            worksheet.write_string  (row, col+3 , root[0].get('policy_name'))
            worksheet.write_string  (row, col+4 , root[0].get('policy_version'))
            worksheet.write_string  (row, col+5 , "NA")
            worksheet.write_string  (row, col+6 , "NA")
            worksheet.write_string  (row, col+7 , root[0][0].get('status'))
            worksheet.write_string  (row, col+8 , "NA")
            worksheet.write_string  (row, col+9 , "NA")

        else:
            print ( "getting build details" )
            xmlBuildDetail = getXMLSummaryReport(root[-1].get('build_id'))
            root = etree.fromstring(xmlBuildDetail)        
            print ('App_id : ' + root.get('app_id') +', Build_id : ' + root.get('build_id') + ', App_name : ' + root.get('app_name') + ', Policy_name : ' + root.get('policy_name') + ', Policy_version : ' + root.get('policy_version') + ', Submitted_Date : ' + root[0].get('submitted_date') + ', Published_Date : ' + root[0].get('published_date') + ', Policy_compliance_status : ' + root.get('policy_compliance_status') + ", Rating: " + root[0].get('rating') + ", Score: " + root[0].get('score'))

            worksheet.write_string  (row, col , root.get('app_id'))
            worksheet.write_string  (row, col+1 , root.get('build_id'))
            worksheet.write_string  (row, col+2 , root.get('app_name'))
            worksheet.write_string  (row, col+3 , root.get('policy_name'))
            worksheet.write_string  (row, col+4 , root.get('policy_version'))
            worksheet.write_string  (row, col+5 , root[0].get('submitted_date'))
            worksheet.write_string  (row, col+6 , root[0].get('published_date'))
            worksheet.write_string  (row, col+7 , root.get('policy_compliance_status'))
            worksheet.write_string  (row, col+8 , root[0].get('rating'))
            worksheet.write_string  (row, col+9 , root[0].get('score'))
            
        row += 1        
    workbook.close()




def getXMLSummaryReport(build_id):
    #curl --compressed --sslv3 -k -v -u username:password https://analysiscenter.veracode.com/api/2.0/summaryreport.do?build_id=111111 -o summaryreport.xml
    payload = {'build_id': build_id}
    r = requests.get("https://analysiscenter.veracode.com/api/2.0/summaryreport.do", params=payload, auth=(username, password))
    return r.content 
  
def getBuildList(app_id):
    payload = {'app_id':app_id}
    r = requests.post("https://analysiscenter.veracode.com/api/4.0/getbuildlist.do", params=payload, auth=(username, password)) 
    return r.text

def getBuildInfo(app_id):
    payload = {'app_id':app_id}
    r = requests.get("https://analysiscenter.veracode.com/api/5.0/getbuildinfo.do",params=payload, auth=(username, password))
    return r.text

def getApplist():

    r = requests.post("https://analysiscenter.veracode.com/api/5.0/getapplist.do",auth=(username, password))
    return r.text

    


if __name__ == "__main__":
    main()

    
