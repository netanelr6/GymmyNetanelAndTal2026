import xlsxwriter
import datetime
import Settings as s
from Joint import Joint
import os


def create_workbook():
    current_time = datetime.datetime.now()
    workbook_name = s.participant_code + ".xlsx"

    #-----netanel edit from MAYA origin ----------
    if hasattr(s, 'output_path'):
        full_path = os.path.join(s.output_path, workbook_name)
    else:
        full_path = workbook_name
    
    #save the EXCEL file in the pathe location
    s.excel_workbook = xlsxwriter.Workbook(full_path)
    # s.excel_workbook = xlsxwriter.Workbook(workbook_name) #maya original line

    # -------------------------------------------


def wf_joints(ex_name, list_joints):
    """
    Writing joints data for an exercise in Excel file in two versions
    :param ex_name:
    :param list_joints:
    :return:
    """
    current_time = datetime.datetime.now()
    name = ex_name + str(current_time.minute) + str(current_time.second)
    s.worksheet = s.excel_workbook.add_worksheet(name)
    frame_number = 1

    for l in range(1, len(list_joints)):
        row = 1
        s.worksheet.write(0, frame_number, frame_number)
        for j in list_joints[l]:
            if type(j) == Joint:
                j_ar = j.joint_to_array()
                for i in range(len(j_ar)):
                    s.worksheet.write(row, frame_number, str(j_ar[i]))
                    row += 1
            else:
                s.worksheet.write(row, frame_number, j)
                row += 1
        frame_number += 1


def success_worksheet():
    row = 1
    col = 0
    s.worksheet = s.excel_workbook.add_worksheet("success")
    for ex in s.ex_list:
        s.worksheet.write(row, col, ex[0])
        s.worksheet.write(row, col+1, ex[1])
        row += 1
        col = 0

    row = 1
    col = 0
    s.worksheet = s.excel_workbook.add_worksheet("performance_class")
    for ex in s.performance_class:
        s.worksheet.write(row, col, ex)
        if s.performance_class[ex]['right'] is not None:
            s.worksheet.write(row, col+1, s.performance_class[ex]['right'])
        else:
            s.worksheet.write(row, col+1, "nan")
        if s.performance_class[ex]['left'] is not None:
            s.worksheet.write(row, col+2, s.performance_class[ex]['left'])
        else:
            s.worksheet.write(row, col+2, "nan")
        row += 1
        col = 0

    s.worksheet = s.excel_workbook.add_worksheet("Experiment_Results")
    s.worksheet.write(0, 0, "Exercise")
    s.worksheet.write(0, 1, "Handled")
    s.worksheet.write(0, 2, "Handled In Exercise")
    s.worksheet.write(0, 3, "Reaction Time (Sec)")
    
    fault_ex = getattr(s, 'fault_exercise', '')
    if fault_ex == "":
        s.worksheet.write(1, 0, "Ideal Mode - No Faults")
        s.worksheet.write(1, 1, "N/A")
        s.worksheet.write(1, 2, "N/A")
        s.worksheet.write(1, 3, "N/A")
    else:
        s.worksheet.write(1, 0, fault_ex)
        s.worksheet.write(1, 1, str(getattr(s, 'fault_handled', False)))
        s.worksheet.write(1, 2, getattr(s, 'fault_handled_exercise', ''))
        s.worksheet.write(1, 3, round(getattr(s, 'reaction_time', 0.0), 2))

def close_workbook():
    s.excel_workbook.close()
