import time
import Settings as s
import Excel
from Camera import Camera
from Poppy import Poppy
from Audio import Audio
from Training import Training
from Screen import Screen, FullScreenApp
from PIL import Image, ImageTk
import pickle
import datetime
import os
import sys


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.
# TODO add more exercises
# TODO GUI
# delay between exercises


if __name__ == '__main__':


    #-----------------------participant settinga -----------------------
    participant_number = 42 #exp: 42
    participant_round = 'A' #exp: 'A'  --->'A'= firs week, 'B'= second week
    participant_numberANDround = f"{participant_number}{participant_round}"   #TODO always set who is partificate(impurtent for file name)
    #numberANDround  --->  example: {participant number = '42'}{round(WEEK) = 'A'} = '42A'

    s.WORKFLOW_MODE = None        # 0=Normal, 1=Hardware, 2=Interactio ,None = select page #TODO --> update the value----->If you don't want a selection screen


    s.experiment_is_waiting_to_start = False #true =we after prase the "tham letter" button
    s.experiment_started = False #true =we after prase the "start expirement" button    


    if s.WORKFLOW_MODE != None:
        s.experiment_is_waiting_to_start = True


    print (participant_numberANDround)
    #--------------------------------------------------------------------




    #-------------------------------------------------
    #NETANEL&tal SETTINGS

    #otput file setting
    s.project_folder = "netanel&tal_2026A_" # fulder name in dats
    s.output_path = ""    # bild in main
    s.save_outputs = True  # Set to False to disable folder creation, Excel saving, and logging



    s.reboot_flag = False
    s.exercise_completed = False




    s.inter_aff = False
    s.hardwere_aff = False
    # s.Team_Number = 1                  # s.WORKFLOW_MODE #todo --->verifide if work welll  #TODO --> update the value
    # print(f"Team_Number is: {s.Team_Number}")
    s.reboot_flag = False

    # RUN_MODE = os.getenv('GYMMY_MODE', 'SIM') #TODO MYBE FOr LATER







    #---------------------------------------------------
    #ORIGINAL SRTTINGS
    s.camera_num = 1  # 0 - webcam, 2 - second USB in maya's computer

    # Audio variables initialization
    language = 'Hebrew'
    gender = 'Male'
    s.audio_path = 'audio files/' + language + '/' + gender + '/'
    s.picture_path = 'audio files/' + language + '/' + gender + '/'
    # s.str_to_say = ""
    current_time = datetime.datetime.now()


    # Training variables initialization
    s.exercise_amount = 6
    s.rep = 8
    s.req_exercise = ""
    s.finish_workout = False
    s.waved = False
    s.success_exercise = False
    s.calibration = False # False to have calibration session, True to not have
    s.training_done = False
    s.poppy_done = False
    s.camera_done = False
    s.robot_count = False #True
    s.try_again = False
    # # Excel variable ##move after we select the grop in screen
    # Excel.create_workbook()
    s.ex_list = []

    # Create all components
    s.camera = Camera()
    s.training = Training()
    s.robot = Poppy()

    # Adaptation variables
    s.adaptive = False #True
    s.corrective_feedback = False
    s.one_hand = False
    s.performance_class = {}
    if s.adaptive:
        s.adaptation_model_name = 'performance_evaluation_model'
        # s.adaptation_model = pickle.load(open(f'{adaptation_model_name}.sav', 'rb'))


    s.screen = Screen()
    #---------------------------------------

    print("Waiting for researcher selection on screen...")
    while not s.experiment_is_waiting_to_start:
        try:
            s.screen.update_idletasks()
            s.screen.update()
            time.sleep(0.01)


        except Exception as e:
            print(f"GUI Interaction Error: {e}")
            break

    #------------File system go up -----------------------
    MOD_code = getattr(s, 'WORKFLOW_MODE', None)
    if not MOD_code:
        s.WORKFLOW_MOD = "INVALID_MODE_CODE"

    s.participant_code = participant_numberANDround + "_"+ "Mode" + f"{s.WORKFLOW_MODE}" + "_"+ str(current_time.day) + "." + str(current_time.month) + " " + str(current_time.hour) + "." + \
                         str(current_time.minute) + "." + str(current_time.second) 

    if s.save_outputs:
        s.output_path = os.path.join("DATS", s.project_folder, s.participant_code)
        os.makedirs(s.output_path, exist_ok=True)

        # Initialize logging
        log_file_path = os.path.join(s.output_path, "code_output.txt")
        log_file = open(log_file_path, "w", encoding="utf-8")

        class Logger(object):
            def __init__(self, terminal, logfile):
                self.terminal = terminal
                self.logfile = logfile

            def write(self, message):
                self.terminal.write(message)
                self.logfile.write(message)
                self.logfile.flush()

            def flush(self):
                pass

        sys.stdout = Logger(sys.stdout, log_file)
        sys.stderr = Logger(sys.stderr, log_file)



        # Excel variable 
        Excel.create_workbook()
        # s.ex_list = []


        print(f"--- Session Initialized. Saving to: {s.output_path} ---")
    else:
        print("--- Debug Mode: Outputs and Logging are DISABLED ---")

    #------------------------------------------------------



    print("Waiting for researcher whitin to start experiment screen...")
    while not s.experiment_started:
        try:
            s.screen.update_idletasks()
            s.screen.update()
            time.sleep(0.01)
        except Exception as e:
            print(f"GUI Interaction Error: {e}")
            break


    #---------------------------------------


    try:
        # Start all threads
        s.camera.start()
        s.training.start()
        s.robot.start()
        image1 = Image.open('Pictures//icon.jpg')
        s.screen.tk.call('wm', 'iconphoto', s.screen._w, ImageTk.PhotoImage(image1))
        app = FullScreenApp(s.screen)
        s.screen.mainloop()
        print("ALL SYSTEMS GO. Workout session in progress.")


    except Exception as e:
        print(f"!!! CRITICAL RUNTIME ERROR: {e}")
        
    finally:
        # Finalize and close resources only if they were initialized
        if s.save_outputs:
            Excel.close_workbook()
            if log_file:
                log_file.close()
            print("--- Data securely saved. ---")
        print("--- System shutdown complete. ---")
    # ==========================================================


