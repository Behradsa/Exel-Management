from openpyxl import load_workbook

workbook = load_workbook(filename="BOM.xlsx")
sheet = workbook.active
SD = sheet.calculate_dimension()
SDF = SD.partition(":")

B = sheet["B"][1::]


def get_level(cell):
    row = cell.row
    level = sheet["D"][row - 1]
    return int(level.value)


def get_material_count(cell):
    row = cell.row
    count = sheet["E"][row - 1]
    return count.value if count.value != None else 0


def get_children_cell(cell):
    output = []
    for i in B:
        i_level = get_level(i)
        i_value = i.value.strip()
        if i_level - 1 == get_level(cell):
            if i_value[: i_level + 1] == cell.value.strip()[: i_level + 1]:
                output.append(i)
    return output


def get_parent_cell(cell):
    for i in B:
        i_level = get_level(i)
        i_value = i.value.strip()
        if i_level + 1 == get_level(cell):
            if (
                i_value[: get_level(cell) + 1]
                == cell.value.strip()[: get_level(cell) + 1]
            ):
                return i


def get_raw_materials(cell):
    output = []
    for i in B:
        i_value = i.value.strip()
        if (
            get_children_cell(i) == []
            and get_level(i) > get_level(cell)
            and cell.value.strip()[: get_level(cell) + 2]
            == i_value[: get_level(cell) + 2]
        ):
            output.append(i)
    return output


def print_parents():
    for cell in B:
        cell_value = cell.value.strip()
        parent = get_parent_cell(cell)
        if parent:
            print(f"{cell_value} Parent is: {parent.value.strip()}")
        else:
            print(f"{cell_value} Parent is: None")


def print_children():
    for cell in B:
        cell_value = cell.value.strip()
        children = get_children_cell(cell)
        if children:
            print(f"{cell_value} Children are:", end=" ")
            for i in children:
                print(i.value.strip(), end=" ")
            print()
        else:
            print(f"{cell_value} Children are: None")


def print_raw_materials_need(count):

    for cell in B:
        sum = 0
        if get_children_cell(cell) != []:
            print(f"For building {count}*{cell.value.strip()} we need:", end=" ")
            raw_materials = get_raw_materials(cell)
            if raw_materials != []:
                for material in raw_materials:
                    print(
                        f"{count*get_material_count(material)}*{material.value.strip()}",
                        end=" ",
                    )
                    sum += count * get_material_count(material)

            print(f"//In total:{sum} materials")


print_children()
print_parents()
print_raw_materials_need(50)
