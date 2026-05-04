def __init__():

    #Netanel&tal - 2024-06-11 - Global variables for the project. These variables are used in all the files and are defined here to avoid circular imports.
    global  project_folder = "netanel&tal_2026A_" # fulder name in dats
    global  output_path = ""    # bild in main
    global  save_outputs = True  # Set to False to disable folder creation, Excel saving, and logging

    global show_reboot_button = True
    global reboot_flag = False
    global exercise_completed = False


    global experiment_started #TODO whay its existe and wher we use it
    #TO MAIN--> global experiment_started = False

    
    global WORKFLOW_MODE
    #to main global WORKFLOW_MODE = 1        # 1=Normal, 2=Hardware, 3=Interactio
    #TODO ADD VALU TO MAIN

    global inter_aff
    global hardwere_aff
    global Team_Number
    global reboot_flag
    # TODO ADD TO MAIN--->  reboot_flag = False


    global RUN_MODE
    # RUN_MODE = os.getenv('GYMMY_MODE', 'SIM') #TODO MYBE FO LATER



    #--------------------------------------------------------------------------------------
    #FORM HEAR DOWN IS THE ORIGINAL CODE

    # classes pointers
    global training
    global camera
    global robot
    global screen

    global participant_code
    global excel_workbook
    global ex_list

    # training variables
    global exercise_amount
    global rep
    global req_exercise
    global finish_workout
    global waved
    global success_exercise
    global calibration
    global poppy_done
    global camera_done
    global robot_count
    global try_again # Adaptive scenario - successful performance
    global robot_rep # number of repetition of the robot

    # audio variables
    global audio_path

    # screen variables
    global picture_path

    global camera_num

    # adaptation
    global adaptation_model
    global adaptive
    global performance_class
    global corrective_feedback
    global one_hand