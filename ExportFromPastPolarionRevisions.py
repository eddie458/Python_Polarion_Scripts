'''
File: ExportFromPastPolarionRevisions_v1.4.py
Version: 1.4
Date: 24th Feb 2025

Description: Script that queries Polarion for requirement attributes from past revisions and exports it into a .csv excel file.
'''

from polarion import polarion
import datetime
import urllib3
import traceback
import argparse
import itertools
import threading
import time
import sys
import csv
import re

Polarion_Client = None
Polarion_Project = None
Project_Name = ""
ReadText = ""
List_Attributes_Final = []
Attributes_Fields = []
Parameter_File_Path = ""
Finish_Export = False

def animate(holder):
    for c in itertools.cycle(['|', '/', '-', '\\']):
        if holder.done:
            break
        sys.stdout.write('\rFetching Polarion  ' + c)
        sys.stdout.flush()
        time.sleep(0.1)
    sys.stdout.write('\rDone!                         \n')

def Parse_Arguments():
    global Parameter_File_Path
    
    parser = argparse.ArgumentParser()
    parser.add_argument('--path', type=str, required=True)
    args = parser.parse_args()
    Parameter_File_Path = args.path
    print("Location of parameters file to read: " + Parameter_File_Path)

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
        print(f"Initiating connection to URL: {server_url}; Project: {project_name}")
        time.sleep(5)
        Polarion_Client = polarion.Polarion(server_url, u, p, verify_certificate = certificate_verification)
        time.sleep(5)
        Polarion_Project = Polarion_Client.getProject(project_name)
    
        print("Connected successfully to Polarion!")
    
    except:
        print(traceback.format_exc())
        Polarion_Client = None
        Polarion_Project = None
        sys.exit(1)

def Change_Team_Value(team_value: str):
    # Replace values of team_n with competency name. Exp: team_2 -> IT&V
    if team_value == "team_0" or team_value == "approvalteam_0":
        return "--"
    elif team_value == "team_1" or team_value == "approvalteam_1":
        return "EE"
    elif team_value == "team_2" or team_value == "approvalteam_2":
        return "IT&V"
    elif team_value == "team_3" or team_value == "approvalteam_3":
        return "HWQA"
    elif team_value == "team_4" or team_value == "approvalteam_4":
        return "ME"
    elif team_value == "team_5" or team_value == "approvalteam_5":
        return "Mfg"
    elif team_value == "team_6" or team_value == "approvalteam_6":
        return "Mfg_Test"
    elif team_value == "team_7" or team_value == "approvalteam_7":
        return "PM"
    elif team_value == "team_8" or team_value == "approvalteam_8":
        return "PTV"
    elif team_value == "team_9" or team_value == "approvalteam_9":
        return "CAQE"
    elif team_value == "team_10" or team_value == "approvalteam_10":
        return "Safety"
    elif team_value == "team_11" or team_value == "approvalteam_11":
        return "SW"
    elif team_value == "team_12" or team_value == "approvalteam_12":
        return "SYS"
    elif team_value == "team_13" or team_value == "approvalteam_13":
        return "Leads"
    elif team_value == "team_14" or team_value == "approvalteam_14":
        return "SYS_Architect"
    elif team_value == "team_15" or team_value == "approvalteam_15":
        return "SW_Architect"
    elif team_value == "team_16" or team_value == "approvalteam_16":
        return "PCM"
    elif team_value == "team_17" or team_value == "approvalteam_17":
        return "CyberSecurity"
    elif team_value == "team_18" or team_value == "approvalteam_18":
        return "VST"
    elif team_value == "team_19" or team_value == "approvalteam_19":
        return "Process_Config_Management"
    elif team_value == "team_20" or team_value == "approvalteam_20":
        return "Process_System"
    elif team_value == "team_21" or team_value == "approvalteam_21":
        return "Process_Others"
    elif team_value == "team_22" or team_value == "approvalteam_22":
        return "Process_SQE"
    elif team_value == "team_23" or team_value == "approvalteam_23":
        return "Process_Tests"
    elif team_value == "team_24" or team_value == "approvalteam_24":
        return "Customer"
    elif team_value == "team_25" or team_value == "approvalteam_25":
        return "Supplier"
    elif team_value == "team_26" or team_value == "approvalteam_26":
        return "CySe_Test"
    elif team_value == "team_27" or team_value == "approvalteam_27":
        return "Supplier_Vector"
    elif team_value == "team_28" or team_value == "approvalteam_28":
        return "Supplier_Renesas"
    else:
        return "N/A"
    
def Export_Past_Revision(document_name: str, revision_number: str):
    global Polarion_Project
    global List_Attributes_Final
    global Attributes_Fields
    global Finish_Export

    class Holder(object):
        done = False

    holder = Holder()
    t = threading.Thread(target=animate, args=(holder,))

    try:
        pattern_search = re.compile("}(.*?)%")
        pattern_location = re.compile("(?<=modules\\/)(.*)(?=\\/workitems)")

        # Attributes_Fields = ['ID','Type','Title','Status','Responsability To Review','Responsability To Verify','Implementation Status','Target Release']
        # Attributes_Fields = ['ID','Location','Type','Title','Status','Author','Created','Updated','Responsability To Review','Responsability To Verify','Responsability To Implement','Priority','Severity','Requirement Type','Safety Level','Implementation Status','Target Release','Derived','Test Method','Verification Strategy','Linked Work Items']
        Attributes_Fields = ['ID','Location','Type','Title','Status','Author','Created','Updated','Responsability To Review','Responsability To Verify','Responsability To Implement','Priority','Severity','Requirement Type','Safety Level','Implementation Status','Target Release','Approval - IT&V','Derived','Test Method','Linked Work Items','Linked Work Items Derived','Issues','Issue Status']

        print("Getting WIs from Polarion...")

        t.start()
        baseline_all = Polarion_Project.searchWorkitemFullItemInBaseline(baselineRevision = int(revision_number), query = document_name, sort = 'id', limit = -1) # limit = -1
        holder.done = True 
        time.sleep(2)
        # print(baseline_all)
        print("Finished getting WI from Polarion!")

        list_attributes_pre = []

        print("Starting export of Polarion WIs!")
        for baseline in baseline_all:
            # Get WI ID
            try:
                workitem_id = baseline.id
                # print(workitem_id)
                list_attributes_pre.append(str(workitem_id))
            except:
                list_attributes_pre.append("N/A")

            # Get Location
            try: 
                workitem_location_pre = pattern_location.findall(str(baseline.location))
                workitem_location = workitem_location_pre[0]
                # print(workitem_location)
                list_attributes_pre.append(str(workitem_location))
            except:
                list_attributes_pre.append("N/A")

            # Get WI Type
            try:
                workitem_type = baseline.type.id
                # print(workitem_type)
                list_attributes_pre.append(str(workitem_type))
            except:
                list_attributes_pre.append("N/A")

            # Get WI Title
            try:
                workitem_title = baseline.title
                # print(workitem_title)
                list_attributes_pre.append(str(workitem_title))
            except:
                list_attributes_pre.append("N/A")
            
            # Get WI Status
            try:
                workitem_status = baseline.status.id
                # print(workitem_status)
                list_attributes_pre.append(str(workitem_status))
            except:
                list_attributes_pre.append("N/A")

            # Get WI Author
            try:
                workitem_author_pre = pattern_search.findall(str(baseline.author.uri))
                workitem_author = workitem_author_pre[0]
                # print(workitem_author)
                list_attributes_pre.append(str(workitem_author))
            except:
                list_attributes_pre.append("N/A")

            # Get WI Created
            try:
                workitem_created = baseline.created.strftime("%d/%b/%Y %H:%M:%S")
                # print(workitem_created)
                list_attributes_pre.append(str(workitem_created))
            except:
                list_attributes_pre.append("N/A")

            # Get WI Updated
            try:
                workitem_updated = baseline.updated.strftime("%d/%b/%Y %H:%M:%S")
                # print(workitem_updated)
                list_attributes_pre.append(str(workitem_updated))
            except:
                list_attributes_pre.append("N/A")

            # Get WI Competency To Review
            try:
                list_review_final = []
                found_review = 0
                for individual_custom_review in baseline.customFields.Custom:
                    if individual_custom_review.key == 'competencyToReview':
                        found_review = 1
                        for enums in individual_custom_review.value.EnumOptionId:
                            list_review_final.append(Change_Team_Value(str(enums.id)))
                        workitem_competency_review = ', '.join(map(str, list_review_final))
                        # print(workitem_competency_review)
                        list_attributes_pre.append(str(workitem_competency_review))
                if found_review == 0:
                    list_attributes_pre.append("N/A")
            except:
                list_attributes_pre.append("N/A")

            # Get WI Competency to Verify
            try:
                list_verify_final = []
                found_verify = 0
                for individual_custom_verify in baseline.customFields.Custom:
                    if individual_custom_verify.key == 'competencyToVerify':
                        found_verify = 1
                        for enums in individual_custom_verify.value.EnumOptionId:
                            list_verify_final.append(Change_Team_Value(str(enums.id)))
                        workitem_competency_verify = ', '.join(map(str, list_verify_final))
                        # print(workitem_competency_verify)
                        list_attributes_pre.append(str(workitem_competency_verify))
                if found_verify == 0:
                    list_attributes_pre.append("N/A")
            except:
                list_attributes_pre.append("N/A")

            # Get WI Competency to Implement
            try:
                list_implement_final = []
                found_implement = 0
                for individual_custom_implement in baseline.customFields.Custom:
                    if individual_custom_implement.key == 'competencyToImplement':
                        found_implement = 1
                        for enums in individual_custom_implement.value.EnumOptionId:
                            list_implement_final.append(Change_Team_Value(str(enums.id)))
                        workitem_competency_implement = ', '.join(map(str, list_implement_final))
                        # print(workitem_competency_implement)
                        list_attributes_pre.append(str(workitem_competency_implement))
                if found_implement == 0:
                    list_attributes_pre.append("N/A")
            except:
                list_attributes_pre.append("N/A")
            
            # Get WI Priority
            try:
                workitem_priority = baseline.priority['id']
                # print(workitem_priority)
                list_attributes_pre.append(str(workitem_priority))
            except:
                list_attributes_pre.append("N/A")

            # Get WI Severity
            try:
                workitem_severity = baseline.severity['id']
                # print(workitem_severity)
                list_attributes_pre.append(str(workitem_severity))
            except:
                list_attributes_pre.append("N/A")

            # Get WI Requirement Type
            try:
                found_type = 0
                for individual_custom_type in baseline.customFields.Custom:
                    if individual_custom_type.key == 'requirementCategory':
                        found_type = 1
                        workitem_type_status = individual_custom_type.value['id']
                        # print(workitem_type_status)
                        list_attributes_pre.append(str(workitem_type_status))
                if found_type == 0: 
                    list_attributes_pre.append("N/A")
            except:
                list_attributes_pre.append("N/A")
        
            # Get WI Safety Level
            try:
                found_safety = 0
                for individual_custom_safety in baseline.customFields.Custom:
                    if individual_custom_safety.key == 'safetyLevel':
                        found_safety = 1
                        workitem_safety_status = individual_custom_safety.value['id']
                        # print(workitem_safety_status)
                        list_attributes_pre.append(str(workitem_safety_status))
                if found_safety == 0: 
                    list_attributes_pre.append("N/A")
            except:
                list_attributes_pre.append("N/A")

            # Get WI Implementation Status
            try:
                found_implementation = 0
                for individual_custom_implementation in baseline.customFields.Custom:
                    if individual_custom_implementation.key == 'implementationStatus':
                        found_implementation = 1
                        workitem_implementation_status = individual_custom_implementation.value.EnumOptionId[0].id
                        # print(workitem_implementation_status)
                        list_attributes_pre.append(str(workitem_implementation_status))
                if found_implementation == 0: 
                    list_attributes_pre.append("N/A")
            except:
                list_attributes_pre.append("N/A")

            # Get WI Target Release
            try:
                list_target_final = []
                found_target = 0
                for individual_custom_target in baseline.customFields.Custom:
                    if individual_custom_target.key == 'Targetrelease':
                        found_target = 1
                        for enums in individual_custom_target.value.EnumOptionId:
                            list_target_final.append(enums.id)
                        workitem_target_release = ', '.join(map(str, list_target_final))
                        # print(workitem_target_release)
                        list_attributes_pre.append(str(workitem_target_release))
                if found_target == 0: 
                    list_attributes_pre.append("N/A")
            except:
                list_attributes_pre.append("N/A")

            # Get WI Approved by IT&V
            try:
                found_approve_itv = 0
                for individual_custom_approve_itv in baseline.customFields.Custom:
                    if individual_custom_approve_itv.key == 'approvalteam_2':
                        found_approve_itv = 1
                        workitem_approve_itv_status = individual_custom_approve_itv.value
                        # print(workitem_approve_itv_status)
                        list_attributes_pre.append(str(workitem_approve_itv_status))
                if found_approve_itv == 0: 
                    list_attributes_pre.append("N/A")
            except:
                list_attributes_pre.append("N/A")

            # Get WI Derived
            try:
                found_derived = 0
                for individual_custom_derived in baseline.customFields.Custom:
                    if individual_custom_derived.key == 'derived':
                        found_derived = 1
                        workitem_derived_status = individual_custom_derived.value['id']
                        # print(workitem_derived_status)
                        list_attributes_pre.append(str(workitem_derived_status))
                if found_derived == 0: 
                    list_attributes_pre.append("N/A")
            except:
                list_attributes_pre.append("N/A")

            # Get WI Test Method
            try:
                found_method = 0
                for individual_custom_method in baseline.customFields.Custom:
                    if individual_custom_method.key == 'tsMethod':
                        found_method = 1
                        workitem_method_status = individual_custom_method.value
                        # print(workitem_method_status)
                        list_attributes_pre.append(str(workitem_method_status))
                if found_method == 0: 
                    list_attributes_pre.append("N/A")
            except:
                list_attributes_pre.append("N/A")

            # # Get WI Verification Strategy
            # try:
            #     list_strategy_final = []
            #     found_strategy = 0
            #     for individual_custom_strategy in baseline.customFields.Custom:
            #         if individual_custom_strategy.key == 'verificationMethod':
            #             found_strategy = 1
            #             for enums in individual_custom_strategy.value.EnumOptionId:
            #                 list_strategy_final.append(Change_Team_Value(str(enums.id)))
            #             workitem_competency_strategy = ', '.join(map(str, list_strategy_final))
            #             # print(workitem_competency_strategy)
            #             list_attributes_pre.append(str(workitem_competency_strategy))
            #     if found_strategy == 0:
            #         list_attributes_pre.append("N/A")
            # except:
            #     list_attributes_pre.append("N/A")
            
            # Get WI Linked Work Items
            list_linked_final = []
            try:
                for enums in baseline.linkedWorkItems.LinkedWorkItem:
                    linked_regex = pattern_search.findall(str(enums['workItemURI']))
                    linked_pre = str(enums['role']['id']) + ": " + linked_regex[0] 
                    list_linked_final.append(linked_pre)
                workitem_linked = ', '.join(map(str, list_linked_final))
                # print(workitem_linked)
                list_attributes_pre.append(str(workitem_linked))
            except:
                list_attributes_pre.append("N/A")

            # Get WI Linked Work Items Derived
            list_linked_derived_final = []
            try:
                for enums_derived in baseline.linkedWorkItemsDerived.LinkedWorkItem:
                    linked_derived_regex = pattern_search.findall(str(enums_derived['workItemURI']))
                    linked_derived_pre = str(enums_derived['role']['id']) + ": " + linked_derived_regex[0] 
                    list_linked_derived_final.append(linked_derived_pre)
                workitem_linked_derived = ', '.join(map(str, list_linked_derived_final))
                # print(workitem_linked_derived)
                list_attributes_pre.append(str(workitem_linked_derived))
            except:
                list_attributes_pre.append("N/A")

            # Get WI Linked Work Items that are issues
            list_issues_final = []
            list_issue_ids = []
            issues_found = 0
            try:
                for enums_issues in baseline.linkedWorkItemsDerived.LinkedWorkItem:
                    linked_issues_regex = pattern_search.findall(str(enums_issues['workItemURI']))
                    if enums_issues['role']['id'] == "issue":
                        issues_found = 1
                        list_issue_ids.append(linked_issues_regex[0])
                        linked_issues_pre = str(enums_issues['role']['id']) + ": " + linked_issues_regex[0] 
                        list_issues_final.append(linked_issues_pre)
                        workitem_issues_derived = ', '.join(map(str, list_issues_final))
                # print(list_issues_final)
                if issues_found == 1:
                    list_attributes_pre.append(str(workitem_issues_derived))
                else: 
                    list_attributes_pre.append("N/A")
            except:
                list_attributes_pre.append("N/A")

            list_issues_status_final = []
            list_issues_status_pre = []
            if issues_found == 1:
                for solo_issue in list_issue_ids:
                    issue_workitem = Polarion_Project.getWorkitem(str(solo_issue))
                    list_issues_status_pre = str(solo_issue) + ": " + str(issue_workitem.status.id)
                    list_issues_status_final.append(list_issues_status_pre)
                    workitem_issues_status = ', '.join(map(str, list_issues_status_final))
                list_attributes_pre.append(str(workitem_issues_status))
            else:
                list_attributes_pre.append("N/A")

            List_Attributes_Final.append(list(list_attributes_pre))
            list_attributes_pre.clear()
            Finish_Export = True
        
        # print(List_Attributes_Final)
        
        print("Finished exporting WI from Polarion!")
    except:
        print(traceback.format_exc())
        holder.done = True 

def Save_To_CSV(save_location: str):
    global List_Attributes_Final
    global Attributes_Fields
    global Finish_Export

    if Finish_Export == True:
        folder = save_location
        current_time = datetime.datetime.now()
        # filename = folder + "\\" + "Export_" + str(ReadText.splitlines()[5]) + "_" + str(current_time.year) + "_" + str(current_time.month) + "_" + str(current_time.day) + "_" + str(current_time.hour) + "_" + str(current_time.minute) + "_" + str(current_time.second) +  r".csv"
        filename = folder + "\\" + "Export_" + str(ReadText.splitlines()[4]) + "_" + str(ReadText.splitlines()[5]) + "_" + str(current_time.strftime("%d_%b_%Y_%H_%M_%S")) +  r".csv"
        print("File will be saved in the following directory: " + filename)
        
        # Writing to csv file
        with open(filename, 'w', newline = '', encoding = "utf-8") as csvfile:
            # Creating a csv writer object
            csvwriter = csv.writer(csvfile)
            # Writing the fields
            csvwriter.writerow(Attributes_Fields)
            # Creating the data rows
            csvwriter.writerows(List_Attributes_Final)

if __name__ == '__main__':
    
    Parse_Arguments()

    Read_File(path_inf = Parameter_File_Path)

    Polarion_Programs = {"BMW" : ("https://polarion.asux.aptiv.com/polarion", "BMW_SDPS_GEN3dot2", ReadText.splitlines()[2]),
                     "STLA": ("http://polarionprod1.aptiv.com/polarion", "10028855_FCA_MY21_WL_OCM", ReadText.splitlines()[2]),
                     "GB_MY23"  : ("http://polarionprod1.aptiv.com/polarion", "10033704_GM_Global_B_MY23", ReadText.splitlines()[2]),
                     "GM_SDV" : ("https://polarion.asux.aptiv.com/polarion", "GM_SDV", ReadText.splitlines()[2])}

    Init_Connection(ReadText.splitlines()[3])

    Export_Past_Revision(ReadText.splitlines()[4], ReadText.splitlines()[5])
    Save_To_CSV(ReadText.splitlines()[6])

    print("Done!")

