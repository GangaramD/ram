__author__ = 'dbergert'

import argparse
import requests
import xml.etree.ElementTree as etree
import xlsxwriter

parser = argparse.ArgumentParser(description='Retrieve Veracode Scan Results')
#parser.add_argument("username", help="Veracode API Username")
#parser.add_argument("password", help="Veracode API Password")
parser.add_argument("app_id", help="Veracode Application Id")
#parser.add_argument("projectName", help="Project Name")

args = parser.parse_args()

username = "Nsatest_api"
password = "9LqADFYYCCl4"
app_id = args.app_id
#projectName = args.projectName

list=[]
build_id = None
version = None


def main():
    
    print ("Getting Last Build ID of Application: ")
    xmlBuildList = getBuildList()
    namespace = "{https://analysiscenter.veracode.com/schema/2.0/buildlist}"
    root = etree.fromstring(xmlBuildList)
    for build in root.findall(".//{0}build".format(namespace)):
        build_id = build.get('build_id')
        version = build.get('version')
    print (build_id, version)

    print ( "getting build details" )
    xmlBuildDetail = getXMLSummaryReport(build_id)
    namespace = "{https://www.veracode.com/schema/reports/export/1.0}"
    root = etree.fromstring(xmlBuildDetail)
    str_list = []
    str_list1 = []
    #Write Rating and Score:
    for sa in root.findall(".".format(namespace)):
        result = str_list.append('App_id : ' + sa.get('app_id') +', Build_id : ' + sa.get('build_id') + ', App_name : ' + sa.get('app_name') + ', Policy_name : ' + sa.get('policy_name') + ', Policy_version : ' + sa.get('policy_version') + ', Policy_compliance_status : ' + sa.get('policy_compliance_status'))
    value = [i.split(',') for i in str_list]
    value1 = [i.split(':')for i in value[0]]
    #value2 = dict([jj[0].split("=")[0],jj[1:]] for jj in value1)  
    #print (value1)
    for sa in root.findall(".//{0}static-analysis".format(namespace)):
        result = str_list1.append('Rating : ' + sa.get('rating') + ', Score : ' + sa.get('score'))
    #print (str_list)
    final = [i.split(',') for i in str_list1]
    final2 = [i.split(':')for i in final[0]]
    #final3 = dict([jj[0].split("=")[0],jj[1:]] for jj in final2) 
    
    #print (final2)
    lastfinal=value1+final2
    print (lastfinal)

    #for t in test:
    #print(test)
    
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('Expenses03.xlsx')
    worksheet = workbook.add_worksheet()

         # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': 1})

         # Adjust the column width.
    worksheet.set_column(1,15)

         # Write some data headers.
    worksheet.write('A1', 'App_ID' ,bold )
    worksheet.write('B1', 'Build_Id' ,bold )
    worksheet.write('C1', 'App_Name',bold )
    worksheet.write('D1', 'Policy_Name',bold )
    worksheet.write('E1', 'Policy_Version',bold )
    worksheet.write('F1', 'Policy_Compliance_Status',bold )
    worksheet.write('G1', 'Rating' ,bold)
    worksheet.write('H1', 'Score' ,bold)
         
         

         # Some data we want to write to the worksheet.
    """expenses = (
             [['App_id ', ' 234454'],
              ['build_id ', ' 1310167'],
              [' app_name ', ' Isf-joinfilter'],
              [' policy_name ', ' Thomson Reuters Standard'],
              [' policy_version ', ' 12'],
              [' policy_compliance_status ', ' Pass'],
              ['Rating ', ' A'],
              [' Score ', ' 93']])"""

         # Start from the first cell below the headers.
    row = 1
    col = 0

    for item, cost in (lastfinal):
        # Convert the date string into a datetime object.
        #worksheet.write_string  (row, col,     item )             
        worksheet.write_string  (row, col , cost)
        col += 1

         # Write a total using a formula.
         #worksheet.write(row, 0, 'Total', bold)
         #worksheet.write(row, 2, '=SUM(C2:C5)', money_format)
    workbook.close()





def getXMLSummaryReport(build_id):
    #curl --compressed --sslv3 -k -v -u username:password https://analysiscenter.veracode.com/api/2.0/summaryreport.do?build_id=111111 -o summaryreport.xml
    payload = {'build_id': build_id}
    r = requests.get("https://analysiscenter.veracode.com/api/2.0/summaryreport.do", params=payload, auth=(username, password))
    return r.content 
    
def getBuildList():
    #curl --compressed -u username:password  https://analysiscenter.veracode.com/api/4.0/getbuildlist.do -F "app_id=111111"
    payload = {'app_id':app_id}
    r = requests.post("https://analysiscenter.veracode.com/api/4.0/getbuildlist.do", params=payload, auth=(username, password))
    #print (r.text);
    return r.text
    #print (r.text)




    


if __name__ == "__main__":
    main()

    

