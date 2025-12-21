import ezdxf
import os
import pandas as pd
import numpy as np
from shapely.geometry import Polygon
from tkinter import Tk, filedialog, messagebox

def polygon_from_entity(entity):
    if entity.dxftype() == "LWPOLYLINE" and entity.closed:
        return Polygon(entity.get_points("xy"))
    return None

def circumscribed_circle(points):
    center = np.mean(points, axis=0)
    radius = max(np.linalg.norm(p - center) for p in points)
    return 2 * radius

# Hide Tk window
root = Tk()
root.withdraw()

folder = filedialog.askdirectory(title="Select DXF Folder")

if not folder:
    messagebox.showerror("Error", "No folder selected")
    exit()

results = []

for file in os.listdir(folder):
    if not file.lower().endswith(".dxf"):
        continue

    try:
        doc = ezdxf.readfile(os.path.join(folder, file))
        msp = doc.modelspace()

        polygons = []
        for e in msp:
            poly = polygon_from_entity(e)
            if poly and poly.is_valid:
                polygons.append(poly)

        if not polygons:
            raise ValueError("No closed contours")

        polygons.sort(key=lambda p: p.area, reverse=True)
        outer = polygons[0]
        inners = polygons[1:]

        net_area = outer.area - sum(p.area for p in inners)
        chambers = len(inners)

        points = np.array(outer.exterior.coords)
        circle_dia = circumscribed_circle(points)

        results.append([
            file,
            round(net_area, 3),
            round(circle_dia, 3),
            chambers,
            "OK"
        ])

    except:
        results.append([file, "", "", "", "Error"])

df = pd.DataFrame(results, columns=[
    "File Name",
    "Cross Section Area (mm2)",
    "Circumscribed Circle Diameter (mm)",
    "Chambers",
    "Status"
])

output_path = os.path.join(folder, "DXF_Analysis_Output.xlsx")
df.to_excel(output_path, index=False)

messagebox.showinfo("Done", "Excel file generated successfully!")

