'''
File: TestCaseLinkRequirements.py
Version: 1.1
Date: 10-Oct-2024

Description: Script to automate the linking of requirements to test cases that have requirement IDs in TestSteps 
Note: Python scripts needs an external Parameters.txt file that passes on parameters such as: Polarion User, Polarion Password, Certificate Location, Project Name, Document Name, SYTS Name, and WI ID Nomenclature
Important: Make sure that in the .txt file, no extra spaces are present after each parameter
'''

from polarion import polarion
import re
import urllib3
import traceback

SyTS_Workitems = None
Polarion_Client = None
Polarion_Document = None
Polarion_Project = None
Project_Name = ""
SyTS_Path = ""
ReadText = ""
Workitems_IDs = []
individual_id_list = []
Linked_ID_List = []

def Read_File(path_inf: str):
    global ReadText
    
    try:
        urllib3.disable_warnings()

        # Opens txt file with parameters used in script execution
        parameters = open(path_inf, "r")
        ReadText = parameters.read()
        parameters.close()
    except: 
        print(traceback.format_exc())

def Init_Connection(project_name: str):
    global Polarion_Client
    global Polarion_Project
    global Project_Name
    global ReadText
   
    try:
        urllib3.disable_warnings()

        # Splits the single line in the txt to its single words and puts it into a list
        u = ReadText.splitlines()[0]
        p = ReadText.splitlines()[1]
    
        # Standardize project name
        if (project_name.upper().find("BMW") >= 0):
         server_url = Polarion_Programs["BMW"][0]
         project_name = Polarion_Programs["BMW"][1]
         certificate_verification = Polarion_Programs["BMW"][2]
         Project_Name = "BMW"
      
        elif ((project_name.upper().find("STLA") >= 0) or (project_name.upper().find("STELLANTIS") >= 0) or (project_name.upper().find("STELANTIS") >= 0)):
            server_url = Polarion_Programs["STLA"][0]
            project_name = Polarion_Programs["STLA"][1]
            certificate_verification = Polarion_Programs["STLA"][2]
            Project_Name = "STLA"
            
        elif (project_name.upper().find("GB_MY23") >= 0):
            server_url = Polarion_Programs["GB_MY23"][0]
            project_name = Polarion_Programs["GB_MY23"][1]
            certificate_verification = Polarion_Programs["GB_MY23"][2]
            Project_Name = "GB_MY23"

        elif (project_name.upper().find("GM_SDV") >= 0):
            server_url = Polarion_Programs["GM_SDV"][0]
            project_name = Polarion_Programs["GM_SDV"][1]
            certificate_verification = Polarion_Programs["GM_SDV"][2]
            Project_Name = "GM_SDV"
        else:
            raise NotImplementedError
    
        # Creating the Polarion Client and establishing connection  
        print(f"Initiating connection to url: {server_url}; project: {project_name}")
    
        Polarion_Client = polarion.Polarion(server_url, u, p, verify_certificate = certificate_verification)
        Polarion_Project = Polarion_Client.getProject(project_name)
    
        print("Connection ready")
    
    except:
        print(traceback.format_exc())
        Polarion_Client = None
        Polarion_Project = None

def Get_Work_Items(document_name: str, syts_name: str):
    global SyTS_Workitems
    global Polarion_Document 
    global Polarion_Project
    global Project_Name
    global SyTS_Path
    global Workitems_IDs 
    global individual_id_list 
    global Linked_ID_List
    linked_id_list_pre = []
    workitems_modify = []
    all_workitem_steps = []
    all_step_descriptions = []
    pattern_search_description = re.compile("'description':.[^,]+")
    pattern_search_wi = re.compile(ReadText.splitlines()[6] + r'-\d+')

    try:
        # Creates the path in which the document is present in Polarion: Ex. 44-ProductQualificationTest/PDT_PQT_1_Temp_SyTs_for_DID_4759_Copy
        SyTS_Path = r"%s/%s" % (document_name, syts_name)
        print(SyTS_Path)
        # Fetches the document from Polarion and gets all the WorkItems present in the document
        Polarion_Document = Polarion_Project.getDocument(SyTS_Path)
        SyTS_Workitems = Polarion_Document.getWorkitems()
        
        # Loop that checks all the work items that were parsed, it goes one by one and check if any of them contains teststeps. If this is the case, we are sure this WI is a Test Case
        for workitem in SyTS_Workitems:
            if (workitem.hasTestSteps()):
                workitems_modify.append(workitem)
                test_steps = workitem.getTestSteps()  
                all_workitem_steps.append(test_steps)    
        
        # From the obtained Test Case, we obtain only the test case ID and not the TC title: PGS-183023 : Ex. MEC = 0 and MEC > 0 for Field, Field Fault, Barrier, Barrier Fault Modes -> PGS-183023
        for ids in workitems_modify:
            Workitems_IDs.append(str(ids.id))
        
        print(Workitems_IDs)
    
        # From the obtained test case tables from each WI, we obtain only the step description, and ignore test action and test expectation parts 
        for individual_wi_description in all_workitem_steps:
            all_step_descriptions_pre = pattern_search_description.findall(str(individual_wi_description))
            all_step_descriptions.append(str(all_step_descriptions_pre))

        # From the obtained step description parts, we search if any requirement link is present
        for step_descriptions in all_step_descriptions:
            workitems_ids_pre = pattern_search_wi.findall(str(step_descriptions))
            linked_id_list_pre.append(workitems_ids_pre)
            
        print(linked_id_list_pre) 

        # Eliminate duplicate requirement WIs if more than one WI with the same ID exists in one TestCase
        # A for loop is used to eliminate duplicates for each one of the nested lists inside the main list: Ex. [[], ['PGS-94885'], ['PGS-94871', 'PGS-94868', 'PGS-94871']] -> [[], ['PGS-94885'], ['PGS-94871', 'PGS-94868']]
        for indiv_ID_List in linked_id_list_pre:
            indiv_ID_List = list(set(indiv_ID_List))
            Linked_ID_List.append(indiv_ID_List)

        print(Linked_ID_List)

    except:
        print(traceback.format_exc())

def Link_Work_Items():
    global Polarion_Project
    global Workitems_IDs 
    global Linked_ID_List
    number = 0
    
    # Each test case and linked WI is iterated and for each item, a linked item is created 
    for ind_wi in Workitems_IDs:
        tc_wi = Polarion_Project.getWorkitem(str(ind_wi))
        for sub_id in Linked_ID_List[number]:
            req_id = Polarion_Project.getWorkitem(str(sub_id))
            tc_wi.addLinkedItem(req_id, 'verifies')
        number += 1

if __name__ == '__main__':
    
    Read_File(path_inf = r"C:\Users\ixb3gk\Downloads\TestPolarion\MyParameters.txt")

    Polarion_Programs = {"BMW" : ("https://polarion.asux.aptiv.com/polarion", "BMW_SDPS_GEN3dot2", ReadText.splitlines()[2]),
                     "STLA": ("http://polarionprod1.aptiv.com/polarion", "10028855_FCA_MY21_WL_OCM", ReadText.splitlines()[2]),
                     "GB_MY23"  : ("http://polarionprod1.aptiv.com/polarion", "10033704_GM_Global_B_MY23", ReadText.splitlines()[2]),
                     "GM_SDV" : ("https://polarion.asux.aptiv.com/polarion", "GM_SDV", ReadText.splitlines()[2])}

    Init_Connection(ReadText.splitlines()[3])
    Get_Work_Items(ReadText.splitlines()[4],ReadText.splitlines()[5])
    Link_Work_Items()

    print("Done")
