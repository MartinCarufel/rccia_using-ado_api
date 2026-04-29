import requests
import base64
import json
import re
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Inches, Pt

organization = "STMN-Group"
project = "Data capturing Solutions ART"
team = "SG1"
#Enter your personal ADO token
pat = ""
planid = 2935   # Use Sirios Master library ID
token = base64.b64encode(f":{pat}".encode()).decode()
headers = {
    "Authorization": f"Basic {token}",
    "Content-Type": "application/json"
}

class Rccia:

    def __init__(self):
        self.test_case_xref = {}

    def get_test_suite_doc_number(self, test_suite_id):
        """
        Get info about ADO test suite and get the text within the field Description, it also clean the text
        gather fron ADO from all HTML tag and formating.
        :param test_suite_id: test case number
        :return: clean text of desciption field
        """
        url= (f"https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/{test_suite_id}?api-version=7.0")
        resp = requests.get(url, headers=headers)
        json_response = resp.json()
        try:
            test_spec_eqt_number = (BeautifulSoup(json_response["fields"]["System.Description"], "html.parser").
                                    get_text(separator=" ", strip=True))
            return test_spec_eqt_number
        except Exception as e:
            print(f"Test suite {test_suite_id} doesn't have proper description")
            print(f"The description is: {json_response["fields"]["System.Description"]}")
            print(e)
            return "incorrect or empty description"

    def get_test_suite_id_listing(self, ver_or_val):
        """
        This is the first step to create the xref list. This function get from API get all test suite under the
        test plan 'planid'. For each test suite it check if the test suite is under the test suite name passed
        in parameter (usually verification or validation
        :param ver_or_val: string verification or validation
        :return: Only the last test suite level
        """
        # Get the list of all test suite and expand all sub level test suite under test plan
        url = (f"https://dev.azure.com/{organization}/{project}/_apis/testplan/Plans/{planid}/suites?expand={{expand}}&api-version=7.0")
        resp = requests.get(url, headers=headers)
        json_response = resp.json()
        test_suite_list = []
        parent_test_suite_list = []
        leaf_test_suite = []
        for ts in json_response["value"]:
            try:
                if ts["parentSuite"]["name"] == ver_or_val:  # Check level 1 test suite if parent is under the parameter
                    test_suite_list.append(ts["id"])         # add the test suite id to a list
            except:
                continue
            # Create a list of test suite that have child test suite
            try:
                if ts["parentSuite"]["id"]:
                    parent_test_suite_list.append(ts["parentSuite"]["id"])
            except:
                continue

        for ts in json_response["value"]:
            try:
                if ts["parentSuite"]["id"] in test_suite_list:
                    leaf_test_suite.append(ts["id"])
            except:
                continue
        for ts in test_suite_list:
            if ts not in parent_test_suite_list:
                leaf_test_suite.append(ts)
        return leaf_test_suite

    def create_xref_test_case_id_corresponding_etq_doc_number(self, list_of_test_suite):
        """
        Create a class dictionary with indice: test case id and value : table [test suite document number and title
                                                        Word Document test case ID TCxxxxx]
        :param list_of_test_suite: Table, list of test suite id
        """
        print("Processing: ", end="")
        for test_suite in list_of_test_suite:
            print(".", end="")
            test_case_list = self.get_test_case_id_list2(test_suite)
            for test_case in test_case_list:
                self.test_case_xref[test_case] = [self.get_test_suite_doc_number(test_suite),
                                                  self.get_test_case_TCid(test_case)]
        print("")

    def get_test_case_id_list2(self, test_suite):
        """
        Return WI id for the test case
        """
        url = (f"https://dev.azure.com/{organization}/{project}/_apis/testplan/Plans/{planid}/Suites/{test_suite}/TestCase?api-version=7.0")
        resp = requests.get(url, headers=headers)
        json_response = resp.json()
        test_case_list_jsonobj = json_response["value"]
        test_case_list = []
        pattern = r'\D\D\d{5,7}'
        prog = re.compile(pattern)
        for tc in test_case_list_jsonobj:
            test_case_list.append(tc["workItem"]["id"])
        return test_case_list

    def get_test_case_TCid(self, test_case):
        """
        Get form the test case title field the part that contain TCxxxx id
        :param test_case: int, test case id
        :return: String, only the TC id in format TCxxxxx
        """
        pattern = r'\D\D\d{5,7}'
        prog = re.compile(pattern)
        json_data = self.get_workitem_details(test_case)
        return prog.search(json_data["fields"]["System.Title"]).group(0)

    def get_wiql_query_result(self):
        url = (f"https://dev.azure.com/{organization}/{project}/{team}/_apis/wit/wiql?api-version=7.0")
        body = {
            "query": """
    SELECT
        [System.Id],
        [System.Title],
        [System.WorkItemType]
    FROM WorkItemLinks
    WHERE
        (
            [Source].[System.WorkItemType] IN ('Story', 'Bug')
            AND [Source].[System.IterationPath] UNDER 'Data capturing Solutions ART\\PI-2 2026'
        )
        AND
        (
            [System.Links.LinkType] = 'Microsoft.VSTS.Common.TestedBy-Forward'
        )
        AND
        (
            [Target].[System.WorkItemType] = 'Test Case'
            AND [Target].[System.Tags] CONTAINS 'Verification'
        )
    MODE (MustContain)
    """
        }

        resp = requests.post(url, headers=headers, json=body)
        resp.raise_for_status()
        data = resp.json()
        workitem_id = []
        for item in data["workItemRelations"]:
            # return only bug or story
            if item["rel"] == None:
                workitem_id.append(item["target"]["id"])
        return workitem_id

    def get_workitem_details(self, workitem):
        url = (
            f"https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/{workitem}?$expand=Relations&api-version=7.0")
        resp = requests.get(url, headers=headers)
        return resp.json()
    def get_tested_by(self, jsondata):
        """
        From the json data it extract the list of tested by test case, so it ignore all other WI relation (child, test, ...)
        :param jsondata:
        :return: list
        """
        test_by_TC_id_list = []
        try:
            for relation in jsondata["relations"]:
                if relation["rel"] == "Microsoft.VSTS.Common.TestedBy-Forward":
                    workitem_id = relation["url"].split("/")[-1]  # Get the ADO workitem ID from json object
                    workitem_detail = self.get_workitem_details(workitem_id)
                    # use previoulsy found WI ID to get detailed info if the WI have a tag verification
                    # it will be added to the tested by list
                    if workitem_detail["fields"]["System.Tags"].lower() == "verification":
                        test_by_TC_id_list.append(int(workitem_id))
                    else:
                        # print(f"test case {workitem_id} is not verification")
                        pass
        except Exception as e:
            print(e)
        return test_by_TC_id_list

    def get_impact_analysis(self, jsondata):
        return BeautifulSoup(jsondata["fields"]["Custom.ImpactAnalysis"], "html.parser").get_text(
            separator=" ", strip=True)

    def get_change_log(self, jsondata):
        return BeautifulSoup(jsondata["fields"]["Custom.ChangelogEntry_fullText"], "html.parser").get_text(
            separator=" ", strip=True)

    def get_title(self, jsondata):
        return BeautifulSoup(jsondata["fields"]["System.Title"], "html.parser").get_text(
            separator=" ", strip=True)

    def get_tc_corresponding_spec_etq_number(self, test_case_id):
        doc_number = self.test_case_xref[int(test_case_id)][0]
        return doc_number

    def formating_affecting_doc(self, tested_by_list):
        """
        From ADO WI id extracted by tested_by this function get the related Test spec (DEV-xxxxx - title)
        And also add the related TCxxxx test case numbering instead of ADO id.
        :param tested_by_list:
        :return:
        """
        doc_list = {}
        for tc in tested_by_list:
            try:
                tc_list = doc_list[self.get_tc_corresponding_spec_etq_number(tc)]
                tc_list.append(self.test_case_xref[tc][1])
                doc_list[self.get_tc_corresponding_spec_etq_number(tc)] = tc_list

            except:
                doc_list[self.get_tc_corresponding_spec_etq_number(tc)] = [self.test_case_xref[tc][1]]
        text = []
        for doc, tcs in doc_list.items():
            tcs.sort()
            text.append(f"{doc} - {", ".join(tcs)}")
        return "\n".join(text)



if __name__ == "__main__":
    x = Rccia()

    # Step 1 - Create cross-reference between TC ADO work item ID to a corresponding Test spec ETQ number
    test_suite_list = x.get_test_suite_id_listing("Verification")
    x.create_xref_test_case_id_corresponding_etq_doc_number(test_suite_list)

    # Step 2 - Get the list of all story and bug in the release
    iteration_workitem_list = x.get_wiql_query_result()
    # step 3 - From the list of workitems get the
    # get json data for the WI
    # WI title and write it in the file
    # WI change description and write it in the file
    # WI Impact analysis and write it in the file
    # WI Testby test case list ADO id
    # Associated test spec ETW doc number of each WI related TC

    with open("export_rccia.csv", mode='w', encoding='UTF-8') as f:
        f.write(f"Issue tracker,Title,Change description,Impact analysis,Affected documents\n")
        rccia_table = [["Issue tracker", "Title", "Change Description", "Impact analysis", "Affected documents"]]
        for wi in iteration_workitem_list:
            print(f"Prcessing WI: {wi}")
            rccia_line = []
            wi_json_data = x.get_workitem_details(wi)
            f.write(f"\"{wi}\",")
            rccia_line.append(str(wi))
            f.write(f"\"{x.get_title(wi_json_data)}\",")
            rccia_line.append(x.get_title(wi_json_data))
            f.write(f"\"{x.get_change_log(wi_json_data)}\",")
            rccia_line.append(x.get_change_log(wi_json_data))
            f.write(f"\"{x.get_impact_analysis(wi_json_data)}\",")
            rccia_line.append(x.get_impact_analysis(wi_json_data))
            testedby_list = x.get_tested_by(wi_json_data)
            f.write(f"\"{x.formating_affecting_doc(testedby_list)}\"\n")
            rccia_line.append(x.formating_affecting_doc(testedby_list))
            rccia_table.append(rccia_line)

    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    rows = len(rccia_table)
    cols = len(rccia_table[0])
    table = doc.add_table(rows=rows, cols=cols)
    table.style = 'Table Grid'

    for i, row in enumerate(rccia_table):
        for j, value in enumerate(row):
            table.cell(i,j).text = str(value)
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(9)

    doc.save("export_rccia.docx")







