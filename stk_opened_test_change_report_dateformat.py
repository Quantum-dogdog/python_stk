import time
from tqdm import tqdm

startTime = time.time()
from comtypes.gen import STKObjects, STKUtil, AgStkGatorLib
from comtypes.client import CreateObject, GetActiveObject, GetEvents, CoGetObject, ShowEvents
from ctypes import *
import comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0
from comtypes import GUID
from comtypes import helpstring
from comtypes import COMMETHOD
from comtypes import dispid
from ctypes.wintypes import VARIANT_BOOL
from ctypes import HRESULT
from comtypes import BSTR
from comtypes.automation import VARIANT
from comtypes.automation import _midlSAFEARRAY
from comtypes import CoClass
from comtypes import IUnknown
import comtypes.gen._00DD7BD4_53D5_4870_996B_8ADB8AF904FA_0_1_0
import comtypes.gen._8B49F426_4BF0_49F7_A59B_93961D83CB5D_0_1_0
from comtypes.automation import IDispatch
import comtypes.gen._42D2781B_8A06_4DB2_9969_72D6ABF01A72_0_1_0
from comtypes import DISPMETHOD, DISPPROPERTY, helpstring

# Get reference to running STK instance using comtypes
from comtypes.client import GetActiveObject
uiApplication = GetActiveObject('STK11.Application')
uiApplication.Visible = True
root = uiApplication.Personality2
#获取现在的场景
scenario = root.CurrentScenario


satellite = scenario.Children.New(STKObjects.eSatellite, "qq")
satellite2 = satellite.QueryInterface(STKObjects.IAgSatellite)

# Select Propagator


'''
satellite2.SetPropagatorType(STKObjects.ePropagatorAstrogator)
root.ExecuteCommand('Propagate */Satellite/qq "2 Nov 2000 00:00:00.00" "2 Nov 2000 08:00:00.00"')
root.ExecuteCommand('ReportCreate */Satellite/qq Style "Segment Summary" Type Display')
'''
satellite2.SetPropagatorType(STKObjects.ePropagatorTwoBody)

root.ExecuteCommand('Propagate */Satellite/qq "2 Nov 2000 00:00:00.00" "2 Nov 2000 08:00:00.00"')

root.ExecuteCommand('ReportCreate */Satellite/qq Style "Classical Orbit Elements" Type Display')
root.ExecuteCommand('Units_set * Report DateFormat "UTCJ"')


