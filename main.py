import datetime
import os
import pickle
import random
import time

import kml_load
import OSGridConverter
import easygui
import coord
import winsound
from excel import OpenExcel

ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
EXCEL_EPOCH = datetime.datetime(1900, 1, 1)


class UnknownFormatError(Exception):
    pass


class AerialImage:
    def __init__(self,
                 sortie_number,
                 library_number,
                 frame_number,
                 center_point_grid,
                 run,
                 date,
                 sortie_quality,
                 scale,
                 focal_length,
                 film_details,
                 ):
        self.sortie_number = sortie_number
        self.library_number = library_number
        self.frame_number = frame_number
        self.center_point_grid = center_point_grid
        self.run = run
        self.date = datetime.datetime.strptime(date, "%d %b %Y")
        self.sortie_quality = sortie_quality
        self.scale = scale
        self.focal_length = focal_length
        self.film_details = film_details

    def lon_lat(self):
        lon_lat_object = OSGridConverter.grid2latlong(self.center_point_grid)
        return lon_lat_object.longitude, lon_lat_object.latitude


class Point:
    def __init__(self, title, description, longitude, latitude):
        self.title = title
        self.description = description
        self.longitude = longitude
        self.latitude = latitude


class Line:
    def __init__(self, title, description, start_longitude, start_latitude, end_longitude, end_latitude):
        self.title = title
        self.description = description
        self.start_longitude = start_longitude
        self.start_latitude = start_latitude
        self.end_longitude = end_longitude
        self.end_latitude = end_latitude


def search_files(root_dir=None, key_word=None):
    if root_dir is None:
        root_dir = easygui.diropenbox()
    if root_dir is None:
        return None
    if key_word is None:
        key_word = easygui.enterbox("Enter search keyword:")
    if key_word is None:
        return None
    else:
        key_word = key_word.upper()

    found_files = []
    searched_files = 0
    file_tree = os.walk(root_dir)
    for path in file_tree:
        for file_name in path[2]:
            searched_files += 1
            if key_word in file_name.upper():
                found_files.append(os.path.join(path[0], file_name))
            else:
                pass
            if searched_files % 5000 == 0:
                print("Searched {} files...".format(searched_files))
    winsound.MessageBeep()
    return found_files


# "ENGLISH HERITAGE - NATIONAL MONUMENTS RECORD" at A1
def type_one_cover_search_2_images(path):
    cover_search = OpenExcel(path)
    starts = []
    for r in range(cover_search.getRows()):
        pass
        # Todo
    row = 15
    done = False
    images = []
    while not done:
        row_data = cover_search.read(row)
        if row_data[1] == "":
            done = True
        else:
            new_image = AerialImage(row_data[1],
                                    int(row_data[2]),
                                    int(row_data[4]),
                                    row_data[6],
                                    int(row_data[8]),
                                    row_data[9],
                                    row_data[12],
                                    int(row_data[14]),
                                    int(row_data[15]),
                                    row_data[16])
            images.append(new_image)
        row += 1
    return images


# "Sortie number" at B13
def type_two_cover_search_2_images(path):
    cover_search = OpenExcel(path)
    row = 15
    done = False
    images = []
    while not done:
        row_data = cover_search.read(row)
        if row_data[1] == "":
            done = True
        else:
            new_image = AerialImage(row_data[1],
                                    int(row_data[2]),
                                    int(row_data[4]),
                                    row_data[6],
                                    int(row_data[8]),
                                    row_data[9],
                                    row_data[12],
                                    int(row_data[14]),
                                    int(row_data[15]),
                                    row_data[16])
            images.append(new_image)
        row += 1
    return images


def analise_cover_search(path):
    random_ref = hex(random.randint(1, 4294967296)).upper().strip("0X")
    cover_search = OpenExcel(path)
    data = cover_search.read()
    # Find Table Starts
    found_start = False
    table_starts = []
    sheet_width = cover_search.getCols()
    sheet_height = cover_search.getRows()
    for c in range(sheet_width):
        for r in range(sheet_height):
            cell_data = cover_search.read("{}{}".format(ALPHABET[c], r + 1))
            if ("SORTIE" in str(cell_data).upper()) and ("TOTAL" not in str(cell_data).upper()):
                found_start = True
                table_starts.append((c, r))
        if found_start:
            break
        else:
            pass

    items = []
    for start in table_starts:
        headers = []  # To be list of (column, "header text")

        current_row = start[1]
        data_started = False
        while not data_started:
            current_row += 1
            first_cell_address = "{}{}".format(ALPHABET[start[0]], current_row + 1)
            try:
                cell_data = cover_search.read(first_cell_address)
            except IndexError:
                print("Failed at: {}".format(first_cell_address))
                raise
            if cell_data.count("/") > 0:
                data_started = True
            else:
                pass

        empty_cols = []
        for c in range(sheet_width):
            cell_data = cover_search.read("{}{}".format(ALPHABET[c], current_row + 1))
            if cell_data == "":
                empty_cols.append(c)
            else:
                headers.append([ALPHABET[c], ""])

        for c, h in enumerate(headers):
            cell_data = cover_search.read("{}{}".format(h[0], start[1] + 1))
            if cell_data == "":
                headers[c][1] = "{}:{}".format(random_ref, [c][0])
            else:
                headers[c][1] = cell_data.strip()

        data_started = False
        data_finished = False
        current_row = start[1] + 1

        while not data_finished:
            first_cell_address = "{}{}".format(headers[0][0], current_row + 1)
            try:
                cell_data = cover_search.read(first_cell_address)
            except IndexError:
                data_finished = True
                break

            if cell_data.count("/") > 0:
                data_started = True
            else:
                pass

            new_item = dict()
            for i, h in enumerate(headers):
                cell_address = "{}{}".format(h[0], current_row + 1)
                cell_data = cover_search.read(cell_address)
                if not data_started:
                    if cell_data != "":
                        headers[i][1] += " {}".format(cell_data)
                    else:
                        pass
                else:
                    if (i == 0) and (cell_data.count("/") < 1):
                        data_finished = True
                        break
                    else:
                        new_item[h[1]] = cell_data

            if new_item == dict():
                pass
            else:
                new_item["originating_file"] = path
                items.append(new_item)

            current_row += 1
    return items


def cover_search_2_images(path):
    cover_search = OpenExcel(path)
    if cover_search.read("A1") == "ENGLISH HERITAGE - NATIONAL MONUMENTS RECORD":
        return type_one_cover_search_2_images(path)
    elif cover_search.read("B13") == "Sortie number":
        return type_two_cover_search_2_images(path)
    elif cover_search.read("B10") == "Sortie number":
        return type_two_cover_search_2_images(path)
    else:
        raise UnknownFormatError


def images_2_points(images):
    points = []
    i: AerialImage
    for i in images:
        title = "{}/{}".format(i.sortie_number, i.frame_number)
        description = """Library Number: {}
Frame Number: {}
Centre Point: {}
Run: {}
Date: {:%d/%m/%Y}
Sortie Quality: {}
Scale: 1:{}
Focal Length (in inches): {}
Film Details (in inches): {}""".format(i.library_number,
                                       i.frame_number,
                                       i.center_point_grid,
                                       i.run,
                                       i.date,
                                       i.sortie_quality,
                                       i.scale,
                                       i.focal_length,
                                       i.film_details)
        longitude, latitude = i.lon_lat()
        points.append(Point(title, description, longitude, latitude))
    return points


def batch_load(file_list=None):
    if file_list is None:
        file_list = easygui.fileopenbox("Select file list:", filetypes="*.dat")
    if file_list is None:
        return None

    with open(file_list, "br") as fh:
        file_names = pickle.load(fh)

    all_items = []
    for i, f in enumerate(file_names):
        try:
            results = analise_cover_search(f)
        except:
            print("Error with {}".format(f))
        else:
            for r in results:
                all_items.append(r)
            if (i + 1) % 200 == 0:
                print("{} files processed, {} items found...".format(i, len(all_items)))
    print("Done!")
    winsound.MessageBeep()
    return all_items


def get_headers_in_use(results):
    """Takes list of dicts and returns all used keys without repeats"""
    headers = []
    for r in results:
        for h in r:
            if h not in headers:
                headers.append(h)
            else:
                pass
    headers.sort()
    return headers


def results_to_objects(results):
    objects = []
    date_formats = ["%d-%m-%Y",
                    "%d/%m/%Y",
                    "%d/%m/%y",
                    "%d %b %Y",
                    "%d-%b-%Y",
                    ]
    try:
        for r in results:
            is_point = False
            try:
                p = r["Grid"]
                is_point = True
            except KeyError:
                pass

            title = ""
            try:
                title = "{}/{:0.0f}".format(r["Sortie Number"],
                                            r["Frame Number"])
            except KeyError:
                try:
                    title = "{}/{:0.0f}-{:0.0f}".format(r["Sortie Number"],
                                                        r["Start Frame Number"],
                                                        r["End Frame Number"])
                except KeyError:
                    try:
                        title = "{}/{:0.0f}".format(r["Sortie Number"],
                                                    r["Start Frame Number"])
                    except KeyError:
                        pass
            except ValueError:
                title = "{}".format(r["Sortie Number"])

            while title == "":
                title = easygui.enterbox("No title could be made for record\n{}".format(r))
                if title is None:
                    title = ""
                else:
                    pass

            description = ""

            # Date
            try:
                if type(r["Date"]) == float:
                    try:
                        description += "Date:{:%d/%m/%Y}\n".format(EXCEL_EPOCH + datetime.timedelta(days=r["Date"]))
                    except:
                        print("Failed to add date to results:\{}".format(r))
                else:
                    date_text = ""
                    date_data = r["Date"]
                    tried_adding_zero = False
                    while date_text == "":
                        key_error = False
                        for df in date_formats:
                            try:
                                date_text = "Date: {:%d/%m/%Y}\n".format(datetime.datetime.strptime(date_data, df))
                            except ValueError:
                                pass
                            except KeyError:
                                key_error = True
                        if key_error:
                            break
                        if date_text == "":
                            if tried_adding_zero:
                                date_data = date_data[0:-1]
                                resp = easygui.multenterbox("No date format worked for \'{}\'\n"
                                                            "in {} from file:\n{}\n"
                                                            "Enter new format:".format(date_data,
                                                                                       title,
                                                                                       r["originating_file"]),
                                                            fields=["Data:", "Format:"], values=[date_data, ""])
                                if resp is None:
                                    pass
                                else:
                                    date_data, new_format = resp
                                    if new_format == "":
                                        pass
                                    else:
                                        date_formats.append(new_format)
                            else:
                                date_data += "0"
                                tried_adding_zero = True
                        else:
                            pass
                    if date_text != "":
                        description += "Date: {}\n".format(date_text)
                    else:
                        pass
            except KeyError:
                print("No date on {}".format(title))

            try:
                description += "Scale: 1:{:0.0f}\n".format(r["Scale"])
            except KeyError:
                pass

            try:
                description += "Library Number: {:0.0f}\n".format(r["Library Number"])
            except KeyError:
                pass
            except ValueError:
                description += "Library Number: {}\n".format(r["Library Number"])

            try:
                description += "Quality: {}\n".format(r["Quality"])
            except KeyError:
                pass

            try:
                description += "Run Number: {:0.0f}\n".format(r["Run"])
            except KeyError:
                pass
            except ValueError:
                description += "Run Number: {}\n".format(r["Run"])

            description += "\nFound in:{}".format(r["originating_file"])

            key_error = False
            if is_point:
                lat_lon_obj: OSGridConverter.LatLong = OSGridConverter.grid2latlong(r["Grid"])
                latitude, longitude = lat_lon_obj.latitude, lat_lon_obj.longitude
                new_obj = Point(title, description, longitude, latitude)
            else:
                try:
                    start_lat_lon_obj: OSGridConverter.LatLong = OSGridConverter.grid2latlong(r["Grid Start"])
                    try:
                        end_lat_lon_obj: OSGridConverter.LatLong = OSGridConverter.grid2latlong(r["Grid End"])
                    except KeyError:
                        pass

                except KeyError:
                    try:
                        grid_start, grid_end = r["Grid Start-End"].split(" ", 1)
                        start_lat_lon_obj: OSGridConverter.LatLong = OSGridConverter.grid2latlong(grid_start)
                        end_lat_lon_obj: OSGridConverter.LatLong = OSGridConverter.grid2latlong(grid_end)
                    except KeyError:
                        key_error = True
                if not key_error:
                    start_lat, start_lon = start_lat_lon_obj.latitude, start_lat_lon_obj.longitude
                    end_lon = end_lat = None
                    try:
                        end_lat, end_lon = end_lat_lon_obj.latitude, end_lat_lon_obj.longitude
                    except NameError:
                        pass

                    new_obj = Line(title, description, start_lon, start_lat, end_lon, end_lat)
                else:
                    print("{} has no location!".format(title))
                    new_obj = None
            if new_obj is None:
                pass
            else:
                objects.append(new_obj)
        return objects
    except:
        print("Failed @ {}".format(r))
        raise


def load_results(path=None):
    if path is None:
        path = easygui.fileopenbox("Select file list:", filetypes="*.dat")
    if path is None:
        return None

    with open(path, "br") as fh:
        data = pickle.load(fh)

    return data


def remove_headers(results, headers_to_remove):
    for t in headers_to_remove:
        for r in range(len(results)):
            try:
                p = results[r].pop(t)
            except KeyError:
                pass
    return results


def look_for_duplicates():
    objects = load_results("C:\\Users\\louis\\GitKraken\\CoverSearchProject\\point_line_objects.dat")
    base_list = []
    working_list = []
    for o in objects:
        base_list.append(o)
        working_list.append(o.title)
    duplicates = []
    tot = len(base_list)
    start_time = time.time()
    t = 0
    while len(base_list) > 0:
        title_to_find = base_list[0].title
        title_dupes = []
        w = 0
        reached_end = False
        while not reached_end:
            if base_list[w].title == title_to_find:
                title_dupes.append(base_list.pop(w))
            else:
                pass
            w += 1
            if w >= len(base_list):
                reached_end = True
        assert len(title_dupes) > 0
        duplicates.append(title_dupes)

        if (t + 1) % 500 == 0:
            eta = ((time.time() - start_time) / (t / len(base_list))) - (time.time() - start_time)
            unit = "seconds"
            if eta > 60:
                eta = eta / 60
                unit = "minutes"
            if eta > 60:
                eta = eta / 60
                unit = "hours"
            print("Processed {} of {} points... eta: {:0.2f} {}\n".format(t + 1, len(base_list), eta, unit))
        t += 1
    return duplicates


def consolidate(dupes):
    new_objects = []
    for d in dupes:
        new_object = d[0]
        for i in range(len(d) - 1):
            new_object.description += "\n" + d[i].description.split("\n")[-1]
        new_objects.append(new_object)
    return new_objects