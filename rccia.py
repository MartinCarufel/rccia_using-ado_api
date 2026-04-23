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
pat = ""
planid = 2935
token = base64.b64encode(f":{pat}".encode()).decode()
headers = {
    "Authorization": f"Basic {token}",
    "Content-Type": "application/json"
}

class Rccia:

    def __init__(self):
        self.test_case_xref = {}
        # self.test_case_xref = {14576: ['DEV-0044623', 'TC05001'], 14577: ['DEV-0044623', 'TC05002'], 14578: ['DEV-0044623', 'TC05003'], 14579: ['DEV-0044623', 'TC05004'], 14580: ['DEV-0044623', 'TC05005'], 14581: ['DEV-0044623', 'TC05006'], 14582: ['DEV-0044623', 'TC05007'], 14583: ['DEV-0044623', 'TC05008'], 14584: ['DEV-0044623', 'TC05009'], 14585: ['DEV-0044623', 'TC05010'], 14586: ['DEV-0044623', 'TC05011'], 14587: ['DEV-0044623', 'TC05012'], 14588: ['DEV-0044623', 'TC05013'], 14589: ['DEV-0044623', 'TC05014'], 14590: ['DEV-0044623', 'TC05015'], 14591: ['DEV-0044623', 'TC05016'], 11776: ['DEV-0044629', 'TC05101'], 11777: ['DEV-0044629', 'TC05102'], 11778: ['DEV-0044629', 'TC05103'], 11779: ['DEV-0044629', 'TC05104'], 11780: ['DEV-0044629', 'TC05105'], 11781: ['DEV-0044629', 'TC05106'], 13990: ['DEV-0044629', 'TC05107'], 11782: ['DEV-0044629', 'TC05117'], 11783: ['DEV-0044629', 'TC05108'], 11784: ['DEV-0044629', 'TC05109'], 11785: ['DEV-0044629', 'TC05110'], 11786: ['DEV-0044629', 'TC05111'], 11787: ['DEV-0044629', 'TC05112'], 11788: ['DEV-0044629', 'TC05113'], 11789: ['DEV-0044629', 'TC05114'], 11790: ['DEV-0044629', 'TC05115'], 11791: ['DEV-0044629', 'TC05116'], 11792: ['DEV-0044629', 'TC05118'], 11793: ['DEV-0044629', 'TC05119'], 11794: ['DEV-0044629', 'TC05120'], 11795: ['DEV-0044629', 'TC05121'], 13991: ['DEV-0044629', 'TC05122'], 11844: ['DEV-0044621', 'TC06001'], 11845: ['DEV-0044621', 'TC06002'], 11846: ['DEV-0044621', 'TC06003'], 11847: ['DEV-0044621', 'TC06004'], 11848: ['DEV-0044621', 'TC06005'], 11849: ['DEV-0044621', 'TC06006'], 13766: ['DEV-0044621', 'TC06007'], 11850: ['DEV-0044621', 'TC06008'], 11851: ['DEV-0044621', 'TC06009'], 11852: ['DEV-0044621', 'TC06010'], 11684: ['DEV-0044596', 'TC06101'], 13760: ['DEV-0044596', 'TC06118'], 11685: ['DEV-0044596', 'TC06102'], 11686: ['DEV-0044596', 'TC06103'], 11687: ['DEV-0044596', 'TC06107'], 11688: ['DEV-0044596', 'TC06108'], 11689: ['DEV-0044596', 'TC06109'], 11690: ['DEV-0044596', 'TC06111'], 11691: ['DEV-0044596', 'TC06112'], 11692: ['DEV-0044596', 'TC06113'], 11693: ['DEV-0044596', 'TC06119'], 11694: ['DEV-0044596', 'TC06114'], 11695: ['DEV-0044596', 'TC06115'], 11696: ['DEV-0044596', 'TC06116'], 11697: ['DEV-0044596', 'TC06117'], 17307: ['DEV-0046300', 'TC06301'], 17308: ['DEV-0046300', 'TC06302'], 17309: ['DEV-0046300', 'TC06303'], 17310: ['DEV-0046300', 'TC06304'], 17311: ['DEV-0046300', 'TC06305'], 17312: ['DEV-0046300', 'TC06306'], 17313: ['DEV-0046300', 'TC06307'], 17293: ['DEV-0045609', 'TC06201'], 17294: ['DEV-0045609', 'TC06202'], 17295: ['DEV-0045609', 'TC06203'], 17296: ['DEV-0045609', 'TC06204'], 17297: ['DEV-0045609', 'TC06205'], 17298: ['DEV-0045609', 'TC06206'], 17299: ['DEV-0045609', 'TC06207'], 17300: ['DEV-0045609', 'TC06208'], 17301: ['DEV-0045609', 'TC06209'], 17302: ['DEV-0045609', 'TC06210'], 17303: ['DEV-0045609', 'TC06211'], 17304: ['DEV-0045609', 'TC06212'], 17305: ['DEV-0045609', 'TC06213'], 17306: ['DEV-0045609', 'TC06214'], 11812: ['DEV-0044625', 'TC08001'], 11813: ['DEV-0044625', 'TC08002'], 11814: ['DEV-0044625', 'TC08005'], 11815: ['DEV-0044625', 'TC08007'], 11816: ['DEV-0044625', 'TC08008'], 11817: ['DEV-0044625', 'TC08009'], 11818: ['DEV-0044625', 'TC08010'], 11819: ['DEV-0044625', 'TC08011'], 11820: ['DEV-0044625', 'TC08012'], 11821: ['DEV-0044625', 'TC08013'], 11822: ['DEV-0044625', 'TC08014'], 11823: ['DEV-0044625', 'TC08015'], 11824: ['DEV-0044625', 'TC08016'], 11825: ['DEV-0044625', 'TC08017'], 17323: ['DEV-0044628', 'TC08201'], 17324: ['DEV-0044628', 'TC08202'], 17325: ['DEV-0044628', 'TC08203'], 17326: ['DEV-0044628', 'TC08209'], 17327: ['DEV-0044628', 'TC08204'], 17328: ['DEV-0044628', 'TC08205'], 17329: ['DEV-0044628', 'TC08206'], 17330: ['DEV-0044628', 'TC08207'], 17331: ['DEV-0044628', 'TC08208'], 17316: ['DEV-0044593', 'TC08101'], 17317: ['DEV-0044593', 'TC08102'], 17318: ['DEV-0044593', 'TC08103'], 17319: ['DEV-0044593', 'TC08104'], 17320: ['DEV-0044593', 'TC08105'], 17321: ['DEV-0044593', 'TC08106'], 17322: ['DEV-0044593', 'TC08107'], 17432: ['DEV-0045649', 'TC19001'], 17433: ['DEV-0045649', 'TC19002'], 17434: ['DEV-0045649', 'TC19003'], 17427: ['DEV-0045608', 'TC19101'], 17428: ['DEV-0045608', 'TC19102'], 17429: ['DEV-0045608', 'TC19103'], 17430: ['DEV-0045608', 'TC19104'], 17431: ['DEV-0045608', 'TC19105'], 11725: ['DEV-0044656', 'TC02101'], 11726: ['DEV-0044656', 'TC02102'], 11727: ['DEV-0044656', 'TC02103'], 11728: ['DEV-0044656', 'TC02104'], 11729: ['DEV-0044656', 'TC02105'], 11730: ['DEV-0044656', 'TC02106'], 11731: ['DEV-0044656', 'TC02107'], 11732: ['DEV-0044656', 'TC02108'], 11733: ['DEV-0044656', 'TC02109'], 11735: ['DEV-0044656', 'TC02110'], 11736: ['DEV-0044656', 'TC02111'], 11737: ['DEV-0044656', 'TC02112'], 11702: ['DEV-0044597', 'TC02001'], 11703: ['DEV-0044597', 'TC02002'], 11704: ['DEV-0044597', 'TC02003'], 11705: ['DEV-0044597', 'TC02004'], 11706: ['DEV-0044597', 'TC02005'], 11707: ['DEV-0044597', 'TC02006'], 11708: ['DEV-0044597', 'TC02007'], 11709: ['DEV-0044597', 'TC02008'], 11710: ['DEV-0044597', 'TC02009'], 11711: ['DEV-0044597', 'TC02010'], 11712: ['DEV-0044597', 'TC02011'], 11713: ['DEV-0044597', 'TC02012'], 11714: ['DEV-0044597', 'TC02013'], 11715: ['DEV-0044597', 'TC02014'], 11716: ['DEV-0044597', 'TC02015'], 11717: ['DEV-0044597', 'TC02016'], 11718: ['DEV-0044597', 'TC02017'], 11719: ['DEV-0044597', 'TC02018'], 11720: ['DEV-0044597', 'TC02019'], 11721: ['DEV-0044597', 'TC02020'], 11722: ['DEV-0044597', 'TC02021'], 11723: ['DEV-0044597', 'TC02023'], 11724: ['DEV-0044597', 'TC02023'], 11746: ['DEV-0044600', 'TC01001'], 11747: ['DEV-0044600', 'TC01002'], 11748: ['DEV-0044600', 'TC01003'], 11749: ['DEV-0044600', 'TC01006'], 11750: ['DEV-0044600', 'TC01007'], 11751: ['DEV-0044600', 'TC01008'], 11752: ['DEV-0044600', 'TC01009'], 15120: ['DEV-0044600', 'TC01010'], 17281: ['DEV-0044601', 'TC03001'], 17282: ['DEV-0044601', 'TC03002'], 17283: ['DEV-0044601', 'TC03003'], 17284: ['DEV-0044601', 'TC03004'], 17285: ['DEV-0044601', 'TC03005'], 17286: ['DEV-0044601', 'TC03006'], 17287: ['DEV-0044601', 'TC03007'], 17288: ['DEV-0044601', 'TC03008'], 17289: ['DEV-0044601', 'TC03009'], 11796: ['DEV-0044602', 'TC04001'], 11797: ['DEV-0044602', 'TC04002'], 11798: ['DEV-0044602', 'TC04011'], 11799: ['DEV-0044602', 'TC04003'], 11800: ['DEV-0044602', 'TC04004'], 11801: ['DEV-0044602', 'TC04005'], 11802: ['DEV-0044602', 'TC04006'], 11803: ['DEV-0044602', 'TC04007'], 11804: ['DEV-0044602', 'TC04009'], 11805: ['DEV-0044602', 'TC04010'], 11753: ['DEV-0044594', 'TC07001'], 11754: ['DEV-0044594', 'TC07002'], 11755: ['DEV-0044594', 'TC07003'], 11756: ['DEV-0044594', 'TC07004'], 11757: ['DEV-0044594', 'TC07005'], 11758: ['DEV-0044594', 'TC07006'], 11759: ['DEV-0044594', 'TC07007'], 11760: ['DEV-0044594', 'TC07008'], 11761: ['DEV-0044594', 'TC07009'], 11762: ['DEV-0044594', 'TC07010'], 11763: ['DEV-0044594', 'TC07011'], 11764: ['DEV-0044594', 'TC07012'], 11765: ['DEV-0044594', 'TC07013'], 11766: ['DEV-0044594', 'TC07014'], 11767: ['DEV-0044594', 'TC07015'], 11768: ['DEV-0044594', 'TC07016'], 11769: ['DEV-0044594', 'TC07017'], 11770: ['DEV-0044594', 'TC07018'], 11771: ['DEV-0044594', 'TC07019'], 11772: ['DEV-0044594', 'TC07020'], 11773: ['DEV-0044594', 'TC07021'], 11774: ['DEV-0044594', 'TC07022'], 11775: ['DEV-0044594', 'TC07023'], 18013: ['DEV-0044594', 'TC07024'], 11826: ['DEV-0044604', 'TC09001'], 11827: ['DEV-0044604', 'TC09002'], 11828: ['DEV-0044604', 'TC09012'], 11829: ['DEV-0044604', 'TC09003'], 11830: ['DEV-0044604', 'TC09004'], 11831: ['DEV-0044604', 'TC09005'], 11832: ['DEV-0044604', 'TC09006'], 11833: ['DEV-0044604', 'TC09007'], 11834: ['DEV-0044604', 'TC09008'], 11835: ['DEV-0044604', 'TC09009'], 11836: ['DEV-0044604', 'TC09010'], 11837: ['DEV-0044604', 'TC09011'], 11838: ['DEV-0044604', 'TC09013'], 11839: ['DEV-0044604', 'TC09014'], 11840: ['DEV-0044604', 'TC09015'], 11806: ['DEV-0044605', 'TC10001'], 11807: ['DEV-0044605', 'TC10002'], 11808: ['DEV-0044605', 'TC10003'], 11809: ['DEV-0044605', 'TC10004'], 11810: ['DEV-0044605', 'TC10005'], 11811: ['DEV-0044605', 'TC10006'], 17332: ['DEV-0044612', 'TC11001'], 17333: ['DEV-0044612', 'TC11002'], 17334: ['DEV-0044612', 'TC11003'], 17335: ['DEV-0044612', 'TC11004'], 17336: ['DEV-0044612', 'TC11005'], 11853: ['DEV-0044606', 'TC12006'], 11854: ['DEV-0044606', 'TC12014'], 11855: ['DEV-0044606', 'TC12017'], 11856: ['DEV-0044606', 'TC12018'], 11857: ['DEV-0044606', 'TC12022'], 11858: ['DEV-0044606', 'TC12023'], 11859: ['DEV-0044606', 'TC12025'], 11860: ['DEV-0044606', 'TC12026'], 11861: ['DEV-0044606', 'TC12027'], 11862: ['DEV-0044606', 'TC12028'], 11863: ['DEV-0044606', 'TC12030'], 11864: ['DEV-0044606', 'TC12031'], 11865: ['DEV-0044606', 'TC12032'], 11866: ['DEV-0044606', 'TC12034'], 11867: ['DEV-0044606', 'TC12035'], 11868: ['DEV-0044606', 'TC12036'], 11869: ['DEV-0044606', 'TC12037'], 11870: ['DEV-0044606', 'TC12038'], 11871: ['DEV-0044606', 'TC12039'], 11872: ['DEV-0044606', 'TC12040'], 11873: ['DEV-0044606', 'TC12041'], 11874: ['DEV-0044606', 'TC12042'], 11875: ['DEV-0044606', 'TC12043'], 11876: ['DEV-0044606', 'TC12044'], 11877: ['DEV-0044606', 'TC12045'], 11878: ['DEV-0044606', 'TC12046'], 11879: ['DEV-0044606', 'TC12047'], 11880: ['DEV-0044606', 'TC12048'], 17379: ['DEV-0044607', 'TC13001'], 17381: ['DEV-0044607', 'TC13002'], 17382: ['DEV-0044607', 'TC13003'], 17383: ['DEV-0044607', 'TC13004'], 17384: ['DEV-0044607', 'TC13005'], 17385: ['DEV-0044607', 'TC13006'], 17386: ['DEV-0044607', 'TC13007'], 17387: ['DEV-0044607', 'TC13008'], 17388: ['DEV-0044607', 'TC13010'], 17389: ['DEV-0044607', 'TC13011'], 17390: ['DEV-0044607', 'TC13012'], 17391: ['DEV-0044607', 'TC13014'], 17392: ['DEV-0044607', 'TC13015'], 17393: ['DEV-0044607', 'TC13016'], 17394: ['DEV-0044607', 'TC13017'], 17396: ['DEV-0044608', 'TC14001'], 17397: ['DEV-0044608', 'TC14002'], 17398: ['DEV-0044608', 'TC14003'], 17399: ['DEV-0044608', 'TC14004'], 17400: ['DEV-0044608', 'TC14005'], 17401: ['DEV-0044609', 'TC15001'], 17402: ['DEV-0044609', 'TC15002'], 17403: ['DEV-0044609', 'TC15003'], 17404: ['DEV-0044609', 'TC15004'], 17405: ['DEV-0044609', 'TC15005'], 17406: ['DEV-0044609', 'TC15006'], 17407: ['DEV-0044609', 'TC15007'], 11841: ['DEV-0044610', 'TC16001'], 11842: ['DEV-0044610', 'TC16002'], 11843: ['DEV-0044610', 'TC16003'], 17408: ['DEV-0044578', 'TC17015'], 17409: ['DEV-0044578', 'TC17033'], 17410: ['DEV-0044578', 'TC17001'], 17411: ['DEV-0044578', 'TC17028'], 17412: ['DEV-0044578', 'TC17029'], 17413: ['DEV-0044578', 'TC17030'], 17414: ['DEV-0044578', 'TC17031'], 17415: ['DEV-0044578', 'TC17032'], 17416: ['DEV-0044578', 'TC17014'], 17417: ['DEV-0044578', 'TC17034'], 17418: ['DEV-0044578', 'TC17035'], 17419: ['DEV-0044611', 'TC18001'], 17420: ['DEV-0044611', 'TC18002'], 17421: ['DEV-0044611', 'TC18003'], 17422: ['DEV-0044611', 'TC18004'], 17423: ['DEV-0044611', 'TC18005'], 17424: ['DEV-0044611', 'TC18006'], 17425: ['DEV-0044611', 'TC18007'], 17426: ['DEV-0044611', 'TC18008'], 17435: ['DEV-0045795', 'TC20001'], 17436: ['DEV-0045795', 'TC20002'], 17437: ['DEV-0044599', 'TC99001']}

    def create_document_number_reference(self):
        url = (f"https://dev.azure.com/{organization}/{project}/_apis/testplan/Plans/{planid}/suites/2938/TestCase?api-version=7.0")
        resp = requests.get(url, headers=headers)
        json_response = resp.json()

    def get_test_suite_doc_number(self, test_suite_id):
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
        url = (f"https://dev.azure.com/{organization}/{project}/_apis/testplan/Plans/{planid}/suites?expand={{expand}}&api-version=7.0")
        resp = requests.get(url, headers=headers)
        json_response = resp.json()
        test_suite_list = []
        parent_test_suite_list = []
        leaf_test_suite = []
        for ts in json_response["value"]:

            try:
                if ts["parentSuite"]["name"] == ver_or_val:
                    test_suite_list.append(ts["id"])
            except:
                continue
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
        print("Processing: ", end="")
        for test_suite in list_of_test_suite:
            print(".", end="")
            test_case_list = self.get_test_case_id_list2(test_suite)
            for test_case in test_case_list:
                self.test_case_xref[test_case] = [self.get_test_suite_doc_number(test_suite),
                                                  self.get_test_case_TCid(test_case)]
        print("")


    def get_test_case_id_list(self, test_suite):
        """
        Return the TCxxxx id for the test case
        """
        url = (f"https://dev.azure.com/{organization}/{project}/_apis/testplan/Plans/{planid}/Suites/{test_suite}/TestCase?api-version=7.0")
        resp = requests.get(url, headers=headers)
        json_response = resp.json()
        test_case_list_jsonobj = json_response["value"]
        test_case_list = []
        pattern = r'\D\D\d{5,7}'
        prog = re.compile(pattern)
        for tc in test_case_list_jsonobj:
            test_case_list.append(prog.search(tc["workItem"]["name"]).group(0))
        return test_case_list

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
            # test_case_list.append(prog.search(tc["workItem"]["name"]).group(0))
        return test_case_list

    def get_test_case_TCid(self, test_case):
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

    def get_wiql_query_result_2(self):
        url = (f"https://dev.azure.com/{organization}/{project}/{team}/_apis/wit/wiql?api-version=7.0")
        body = {
            "query": """
    SELECT
        [System.Id],
        [System.Title],
        [System.WorkItemType]
    FROM WorkItems
    WHERE
        [System.WorkItemType] In ('Bug')
        AND [Iteration Path] Under 'Data capturing Solutions ART\\PI-2 2026'
        AND [State] = 'Done'
            """
        }
        resp = requests.post(url, headers=headers, json=body)
        resp.raise_for_status()
        data = resp.json()
    def get_workitem_details(self, workitem):
        url = (
            f"https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/{workitem}?$expand=Relations&api-version=7.0")
        resp = requests.get(url, headers=headers)
        return resp.json()
    def get_tested_by(self, jsondata):
        test_by_TC_id_list = []
        try:
            for relation in jsondata["relations"]:
                if relation["rel"] == "Microsoft.VSTS.Common.TestedBy-Forward":
                    workitem_id = relation["url"].split("/")[-1]
                    workitem_detail = self.get_workitem_details(workitem_id)
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
    # WI Testby tet case list ADO id
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







