import ezdxf
import os
import math
import pandas as pd
from shapely.geometry import Polygon
from tkinter import Tk, filedialog, messagebox


def entity_to_polygon(entity):
    try:
        if entity.dxftype() == "LWPOLYLINE" and entity.closed:
            return Polygon(entity.get_points("xy"))

        if entity.dxftype() == "CIRCLE":
            center = entity.dxf.center
            r = entity.dxf.radius
            points = [
                (
                    center.x + r * math.cos(a),
                    center.y + r * math.sin(a)
                )
                for a in [i * 2 * math.pi / 180 for i in range(180)]
            ]
            return Polygon(points)

        if entity.dxftype() == "ARC":
            if abs(entity.dxf.end_angle - entity.dxf.start_angle) >= 360:
                center = entity.dxf.center
                r = entity.dxf.radius
                points = [
                    (
                        center.x + r * math.cos(a),
                        center.y + r * math.sin(a)
                    )
                    for a in [i * 2 * math.pi / 180 for i in range(180)]
                ]
                return Polygon(points)

    except:
        return None

    return None


def circumscribed_circle_diameter(polygons):
    all_points = []
    for poly in polygons:
        all_points.extend(list(poly.exterior.coords))

    cx = sum(p[0] for p in all_points) / len(all_points)
    cy = sum(p[1] for p in all_points) / len(all_points)

    max_r = max(math.dist((cx, cy), p) for p in all_points)
    return 2 * max_r


def analyze_dxf(file_path):
    doc = ezdxf.readfile(file_path)
    msp = doc.modelspace()

    polygons = []
    for e in msp:
        poly = entity_to_polygon(e)
        if poly and poly.is_valid and poly.area > 0:
            polygons.append(poly)

    if not polygons:
        raise Exception("No valid closed contours found")

    polygons.sort(key=lambda p: p.area, reverse=True)

    outer = polygons[0]
    inners = polygons[1:]

    net_area = outer.area - sum(p.area for p in inners)
    chambers = len(inners)
    circle_dia = circumscribed_circle_diameter(polygons)

    return round(net_area, 3), round(circle_dia, 3), chambers


def main():
    Tk().withdraw()
    folder = filedialog.askdirectory(title="Select DXF Folder")

    if not folder:
        return

    results = []

    for file in os.listdir(folder):
        if file.lower().endswith(".dxf"):
            path = os.path.join(folder, file)
            try:
                area, dia, chambers = analyze_dxf(path)
                results.append([file, area, dia, chambers, "OK"])
            except Exception as e:
                results.append([file, "", "", "", "Error"])

    df = pd.DataFrame(
        results,
        columns=[
            "File Name",
            "Cross Section Area (mm2)",
            "Circumscribed Circle Diameter (mm)",
            "Chambers",
            "Status",
        ],
    )

    output = os.path.join(folder, "DXF_Analysis_Output.xlsx")
    df.to_excel(output, index=False)

    messagebox.showinfo("Done", "DXF Analysis Completed Successfully!")


if __name__ == "__main__":
    main()
