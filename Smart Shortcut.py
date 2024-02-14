import psutil
from AppOpener import open
import time
import threading

programs = [
    ("SrvSurvey.exe", "Srv Survey"),
    ("Elite Dangerous Odyssey Materials Helper.exe", "EDOMH"),
    ("EDMarketConnector.exe", "elite dangerous market connector"),
    ("EDDiscovery.exe", "EDDiscovery"),
    ("EDDI.exe", "EDDI"),
]

def isProcessRunning(process, running_processes):
    '''
    Check if the given process is running.
    '''
    first_letter = process[0].lower()
    for running_process in running_processes:
        if running_process and running_process[0] == first_letter and process.lower() in running_process:
            return True
    return False

def openProgram(process, program):
    '''
    Open a program if it's not already running.
    '''
    if not isProcessRunning(process, running_processes):
        open(program, output=debug)     
    else:
        if debug:
            print(program + " already open")

# Get the list of running processes
running_processes = [proc.name().lower() for proc in psutil.process_iter()]

debug = True
if debug:
    starttime = time.time()

# Create and start a new thread for each program
threads = []
for process, program in programs:
    thread = threading.Thread(target=openProgram, args=(process, program))
    thread.start()
    threads.append(thread)

# Wait for all threads to finish
for thread in threads:
    thread.join()

if debug:
    TimeElapsed = time.time() - starttime
    TimeElapsedSTR = str(TimeElapsed)
    print("runtime " + TimeElapsedSTR )
