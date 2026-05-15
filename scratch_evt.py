import sys
import os

PATH_APP = r"C:\Program Files\DIgSILENT\PowerFactory 2024"
PATH_API = fr"{PATH_APP}\Python\3.12"
PROJECT_NAME = "Import(3)"

if PATH_API not in sys.path: sys.path.append(PATH_API)
os.environ['PATH'] = PATH_APP + ";" + os.environ['PATH']

import powerfactory
app = powerfactory.GetApplication()
app.ActivateProject(PROJECT_NAME)

evt_folder = app.GetFromStudyCase("IntEvt")
if evt_folder is None:
    print("IntEvt folder not found. Creating one...")
    sc = app.GetActiveStudyCase()
    evt_folder = sc.CreateObject("IntEvt", "Events")

evt = evt_folder.CreateObject("EvtSwitch", "TestSwitch")
print("Event created.")
try:
    evt.i_switch = 0
    print("i_switch exists")
except Exception as e:
    print("i_switch error:", e)

try:
    evt.i_all = 0
    print("i_all exists")
except Exception as e:
    print("i_all error:", e)

evt.Delete()
