import ezdxf
import os
import math
import pandas as pd
import numpy as np
from shapely.geometry import Polygon
from tkinter import Tk, filedialog, messagebox


def circle_to_polygon(center, radius, segments=180):
    return Polygon([
        (
            center[0] + radius * math.cos(2 * math.pi * i / segments),
            center[1] + radius * math.sin(2 * math.pi * i / segments)
        )
        for i in range(segments)
    ])


def polygon_from_entity(entity):
    if entity.dxftype() == "LWPOLYLINE" and entity.closed:
        return Polygon(entity.get_points("xy"))

    if entity.dxftype() == "CIRCLE":
        return circle_to_polygon(entity.center, entity.radius)

    return None


def circumscribed_circle_diameter(polygon):
    coords = np.array(polygon.exterior.coords)
    center = coords.mean(axis=0)
    radius = max(np.linalg.norm(p - center) for p in coords)
    return 2 * radius


def analyze_dxf(file_path):
    doc = ezdxf.readfile(file_path)
    msp = doc.modelspace()

    polygons = []
    for e in msp:
        poly = polygon_from_entity(e)
        if poly and poly.is_valid and poly.area > 0:
            polygons.append(poly)

    if not polygons:
        raise ValueError("No valid closed geometry found")

    polygons.sort(key=lambda p: p.area, reverse=True)

    outer = polygons[0]
    inners = polygons[1:]

    net_area = outer.area - sum(p.area for p in inners)
    chambers = len(inners)
    circ_dia = circumscribed_circle_diameter(outer)

    return net_area, circ_dia, chambers


def main():
    root = Tk()
    root.withdraw()

    folder = filedialog.askdirectory(title="Select DXF Folder")
    if not folder:
        return

    results = []

    for file in os.listdir(folder):
        if file.lower().endswith(".dxf"):
            path = os.path.join(folder, file)
            try:
                area, dia, chambers = analyze_dxf(path)
                results.append([
                    file,
                    round(area, 3),
                    round(dia, 3),
                    chambers,
                    "OK"
                ])
            except Exception:
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


if __name__ == "__main__":
    main()
