# -*- coding: utf-8 -*-

import comtypes.client
# from comtypes import COMError
from input import STA_FILE, XYZ_FILE, GRD_FILE

# # Generate modules of necessary typelibs (AutoCAD Civil 3D 2008)
# comtypes.client.GetModule("C:\\Program Files\\Common Files\\Autodesk Shared\\acax17enu.tlb")
# comtypes.client.GetModule("C:\\Program Files\\AutoCAD Civil 3D 2008\\AecXBase.tlb")
# comtypes.client.GetModule("C:\\Program Files\\AutoCAD Civil 3D 2008\\AecXUIBase.tlb")
# comtypes.client.GetModule("C:\\Program Files\\AutoCAD Civil 3D 2008\\Civil\\AeccXLand.tlb")
# comtypes.client.GetModule("C:\\Program Files\\AutoCAD Civil 3D 2008\\Civil\\AeccXUiLand.tlb")
# raise SystemExit


def get_closest_section(station):
    min_dist = station
    for sec, sta in sec_sta:
        dist = abs(station - sta)
        if dist <= min_dist:
            min_dist = dist
            closest_section = sec

    return closest_section

sta_f = open(STA_FILE, "r")
sec_sta = []
for line in sta_f:
    sec, sta = line.strip().split(",")
    sec_sta.append((sec, float(sta)))
sta_f.close()

inp_f = open(XYZ_FILE, "r")

# Get running instance of the AutoCAD application
acadApp = comtypes.client.GetActiveObject("AutoCAD.Application")
aeccApp = acadApp.GetInterfaceObject("AeccXUiLand.AeccApplication.5.0")

# Document object
doc = aeccApp.ActiveDocument
alignment, point_clicked = doc.Utility.GetEntity("Select an alignment:")

sec_xh = {}
for line in inp_f:
    Id, x, y, z = line.strip().split(",")
    x = float(x)
    y = float(y)
    z = float(z)
    station, offset = alignment.StationOffset(x, y)
    section = get_closest_section(station)
    try:
        sec_xh[section].append((offset, z))
    except KeyError:
        sec_xh[section] = [(offset, z)]
inp_f.close()

for value in sec_xh.values():
    value.sort()

out_f_lines = []
for sec, sta in sec_sta:
    try:
        xh = sec_xh[sec]
        out_f_lines.append("*")
        out_f_lines.append("%s      %.2f" % (sec, sta))
        for x, h in xh:
            s = "%.2f  %.2f" % (x, h)
            out_f_lines.append(s.replace("-0.00", "0.00"))
    except KeyError:
        pass
out_f_lines.append("*")

out_f = open(GRD_FILE, "w")
out_f.write("\n".join(out_f_lines))
out_f.close()
