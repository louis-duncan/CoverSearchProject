import datetime
import os
import time

import coord
import main

head_text = """<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2" xmlns:kml="http://www.opengis.net/kml/2.2" xmlns:atom="http://www.w3.org/2005/Atom">
<Folder>
\t<name>OpenArchive Pins</name>"""

tail_text = """</Folder>
</kml>"""

point_base = """\t<Placemark>
\t\t<name>name-here</name>
\t\t<description>description-here</description>
\t\t<LookAt>
\t\t\t<longitude>lon-here</longitude>
\t\t\t<latitude>lat-here</latitude>
\t\t\t<altitude>0</altitude>
\t\t\t<heading>0</heading>
\t\t\t<tilt>0</tilt>
\t\t\t<range>2000</range>
\t\t\t<gx:altitudeMode>relativeToSeaFloor</gx:altitudeMode>
\t\t</LookAt>
\t\t<Point>
\t\t\t<coordinates>lon-here, lat-here</coordinates>
\t\t</Point>
\t</Placemark>"""

line_base = """
\t<Placemark>
\t\t<name>title-here</name>
\t\t<description>description-here</description>
\t\t<styleUrl>#inline</styleUrl>
\t\t<LineString>
\t\t\t<tessellate>1</tessellate>
\t\t\t<coordinates>
\t\t\t\tstart-lon-here,start-lat-here,0 end-lon-here,end-lat-here,0 
\t\t\t</coordinates>
\t\t</LineString>
\t</Placemark>"""


def create_kml_point(title, description, lon, lat):
    assert type(title) == type(description) == str
    assert type(lon) == type(lat) == float
    new_point = point_base
    new_point = new_point.replace("name-here", title)
    new_point = new_point.replace("description-here", description)
    new_point = new_point.replace("lon-here", str(lon))
    new_point = new_point.replace("lat-here", str(lat))
    new_point = new_point.replace("&", "&amp;")
    return new_point


def create_kml_line(title, description, start_lon, start_lat, end_lon, end_lat):
    new_line = line_base
    new_line = new_line.replace("name-here", title)
    new_line = new_line.replace("description-here", description)
    new_line = new_line.replace("start-lon-here", str(start_lon))
    new_line = new_line.replace("start-lat-here", str(start_lat))
    new_line = new_line.replace("end-lon-here", str(end_lon))
    new_line = new_line.replace("end-lat-here", str(end_lat))
    new_line = new_line.replace("&", "&amp;")
    return new_line


def create_batch(points):
    kml_text = ""
    kml_text += head_text
    n = len(points)
    start_time = time.time()
    for i, p in enumerate(points):
        if str(type(p)) == "<class '__main__.Point'>":
            kml_text += create_kml_point(p.title, p.description, p.longitude, p.latitude) + "\n"
        elif str(type(p)) == "<class '__main__.Line'>":
            kml_text += create_kml_line(p.title, p.description,
                                        p.start_longitude, p.start_latitude,
                                        p.end_longitude, p.end_latitude,) + "\n"
        else:
            print("Point {} is unknown type {}!".format(p.title, type(p)))
        if (i + 1) % 1000 == 0:
            now = time.time()
            total_time = ((now - start_time) / i) * n
            eta = datetime.datetime.fromtimestamp(start_time) + datetime.timedelta(seconds=total_time)
            print("Processed {} of {} points... eta: {:%H:%M:%Shrs}\n".format(i + 1, n, eta))
    kml_text += tail_text
    return kml_text
