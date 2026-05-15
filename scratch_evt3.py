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
    sc = app.GetActiveStudyCase()
    evt_folder = sc.CreateObject("IntEvt", "Events")

evt1 = evt_folder.CreateObject("EvtOutage", "TestOutage")
print("EvtOutage created.")
evt1.Delete()

evt2 = evt_folder.CreateObject("EvtSwitch", "TestSwitch")
print("EvtSwitch created.")
evt2.Delete()
