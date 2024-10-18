import xml.etree.ElementTree as ET 
import openpyxl
import sys

class Requirement:
    def __init__(self, requimentName, reqiurementResult, proff):
        self.requimentName = requimentName
        self.proff = proff
        self.requirementResult = reqiurementResult

    def __str__(self):
        # print('Requiment Name' + self.requimentName)
        # print('Result: ' + self.requirementResult)
        # print('Proff: ' + self.proff)
        return f'{self.requimentName}\nResult: {self.requirementResult}\nProff: {self.proff}' 
        
# load report file
report_init = openpyxl.load_workbook(sys.argv[1])
report_excelfile = report_init.active
# load .nessuss file
reportTree = ET.parse(sys.argv[0])
root = reportTree.getroot()
reportElem =root.find('Report')

# extract ReportItem tag from report
reportItemList = reportElem.find('ReportHost').findall('ReportItem')
reportElem.itertext()

# extract ReportItem tag related to compliance check
complianceReportItem = []
cm = {'cm':"http://www.nessus.org/cm"} # namespace
for item in reportItemList:
    if item.get('pluginID') == '21157' and item.get('pluginName') == 'Unix Compliance Checks':
        compliance_check_name = str(item.find('cm:compliance-check-name', cm).text)
        compliance_result = str(item.find('cm:compliance-result', cm).text)
        
        # Find the cm:compliance-actual-value element
        compliance_actual_value_elem = item.find('cm:compliance-actual-value', cm)
        
        # Check if the element exists before calling itertext()
        if compliance_actual_value_elem is not None:
            compliance_actual_value = ''.join(compliance_actual_value_elem.itertext())
        else:
            compliance_actual_value = 'None'

        complianceReportItem.append(Requirement(compliance_check_name,
                                                compliance_result,
                                                compliance_actual_value))

# nonCompleteRequirement = 0
# requirementPASSED = 0
# requirementFAILED = 0
# requirementWARNING = 0
# for item in complianceReportItem:
#     if item.requimentName == None or item.requirementResult == None or item.proff == 'None':
#         nonCompleteRequirement = nonCompleteRequirement + 1
#         if item.requirementResult == 'PASSED':
#             requirementPASSED = requirementPASSED + 1
#         elif item.requirementResult == 'FAILED':
#             requirementFAILED = requirementFAILED + 1
#         else: 
#             requirementWARNING = requirementWARNING + 1
#     print(item)
#     print('-------------------------------------')

# print(f"There are {nonCompleteRequirement} requirements are missing proff. including:")
# print(f"{requirementPASSED} requirement that passed")
# print(f"{requirementFAILED} requirement that failed")
# print(f"{requirementWARNING} warining")
for row in report_excelfile.iter_rows(min_row=11, max_col=7, values_only=False):
    for item in complianceReportItem:  
        if row[4].value is not None and str(row[4].value) in item.requimentName:
            print(row[4].value)
            row[5].value = item.requirementResult
            row[6].value = item.proff
            complianceReportItem.remove(item)
report_init.save('test.xlsx')

