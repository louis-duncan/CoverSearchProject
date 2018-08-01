import csv
import os
import unicodedata


class Item(object):
    lvl = 0
    path = ""
    name = ""

    def __init__(self, lvl, path, name):
        self.lvl = lvl
        self.path = path
        self.name = name


def join_dirs(root, d):
    output = root
    if root.endswith("\\"):
        output += d
    else:
        output += "\\" + d
    if output.endswith("\\"):
        pass
    else:
        output += "\\"
    return output


def pretty_walk(root, level=0, tab_format=False, show_files=True):
    dir_list = []
    file_list = []
    if tab_format:
        lvl = "│  " * level
    else:
        lvl = str(level)
    for i in os.scandir(root):
        if i.is_dir():
            dir_list.append(i.name)
        elif i.is_file():
            file_list.append(i.name)
        else:
            pass
    text = lvl + (" " * (not tab_format)) + root
    if show_files:
        for f in file_list:
            text += "\n" + lvl + (" " * (not tab_format)) + f
    else:
        pass
    for d in dir_list:
        text += "\n" + pretty_walk(join_dirs(root, d), level + 1, tab_format, show_files)
    return text


def walk(root, level=0):
    dir_list = []
    file_list = []
    try:
        for i in os.scandir(root):
            if i.is_dir():
                dir_list.append(i.name)
            elif i.is_file():
                file_list.append(Item(level, (i.path[0:len(i.path) - len(i.name)]), i.name))
            else:
                pass
        for d in dir_list:
            for f in walk(join_dirs(root, d), level + 1):
                file_list.append(f)
    except PermissionError:
        file_list.append(Item(level, root, "PERMISSION DENIED"))
    except FileNotFoundError:
        file_list.append(Item(level, root, "COULD NOT FIND"))
    return file_list


def main():
    input("Press <ENTER> to begin...")
    root = input("Enter root dir: ")
    output_dir = input("Enter output file: ")

    print("Traversing files...")
    file_list = walk(root)

    print("%s items found." % len(file_list))

    print("Formatting data...")
    output_lines = []
    current_dir = ""
    for f in file_list:
        if f.path != current_dir:
            current_dir = f.path
            output_lines.append("")
            output_lines.append(current_dir)
            output_lines.append("")
        else:
            pass
        output_lines.append(f.name)

    print("Writing output...")
    with open(output_dir, "w", newline="") as file:
        csv_writer = csv.writer(file, dialect="excel")
        for r in output_lines:
            try:
                f_r = unicodedata.normalize('NFKD', r).encode("ascii", "replace").decode()
                csv_writer.writerow([f_r])
            except UnicodeEncodeError:
                csv_writer.writerow(["Encoding Error"])

if __name__ == "__main__":
    main()
