import datetime
import os
import pickle

import OSGridConverter
import easygui
import coord
import winsound
from excel import OpenExcel

ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

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
                table_starts.append(r)
        if found_start:
            break
        else:
            pass

    items = []
    for start in table_starts:
        headers = [] # To be list of (column, "header text")

        empty_cols = []
        for c in range(sheet_width):
            cell_data = cover_search.read("{}{}".format(ALPHABET[c], start + 1))
            if cell_data == "":
                empty_cols.append(c)
            else:
                headers.append([ALPHABET[c], cell_data.strip()])

        data_started = False
        data_finished = False
        current_row = start + 1

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


def load_results(path=None):
    if path is None:
        path = easygui.fileopenbox("Select file list:", filetypes="*.dat")
    if path is None:
        return None

    with open(path, "br") as fh:
        data = pickle.load(fh)

    return data