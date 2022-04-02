from curses.ascii import isdigit
import datetime
import random
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell


begin_time = datetime.datetime.now()

# Variable init
sacs = []
tubes = []
tubes_PA_PB = []
tubes_PC_PD = []
pilluliers = []
sacs_zip = []
col_start = 0
row_start = 1
boite_BS = []
boite_RO = []
boite_RH = []
boite_LF = []

random_PAB_BS = []
random_PAB_RH = []
random_PAB_LF = []
random_PAB_RO = []

random_PCD_BS = []
random_PCD_RH = []
random_PCD_LF = []
random_PCD_RO = []

sample_coordinate_box = {}
sample_coordinate_random = {}

sample_coordinate_box.update({"Echantillons":"Coordonnée"})
sample_coordinate_random.update({"Echantillons":"Coordonnée"})

tab1 = ["PA","PB","PC","PD"]
tab1bis = ["PA","PB","PC","PD","Culturomique","Culturomique"]
tab2 = ["BS","RH","LF","RO"]

# Find the longest word in a list
# Param => list : List
# Return => int : nomber of char in the longest word
def find_longest_word(list):
    max_length = 0
    for x in list:
        if len(x)>max_length:
            max_length = len(x)
    return max_length

# Write in the excel a list from top to bot at the column index given by the user
# Param => col : index of the column to write the list
# Param => title : title to write in the first row
# Param => list : list to write
def write_list_in_excel(col,title,list):
    worksheet.set_column(col, col, find_longest_word(list))
    worksheet.write(0,col,title,title_cell)
    for row in range(len(list)):
        worksheet.write(row+1, col, list[row])

# Write in the excel a box from a list with the given size 
# Param => length : length of the box
# Param => width : width of the box
# Param => global_col : index of the col to start to write
# Param => global_row : index of the row to start to write
# Param => list : list to write
# Param => index : index of the element in the list to start
# Param => box_number : number of the box 
def box(length,width,global_col,global_row,list,index,box_number):
    worksheet.set_column(global_col, global_col+width-1, find_longest_word(list))
    worksheet.merge_range(global_row-1,global_col,global_row-1, global_col+width-1, 'Merged Range')
    worksheet.write(global_row-1,global_col,"Back up " +str(box_number),title_cell)
    

    for row in range(0,length):
        for col in range(0,width):
            if index < len(list):
                worksheet.write(row+global_row, col+global_col, list[index])
                index = index+1
    return index

# Write in the excel a random box from a list with the given size 
# Param => length : length of the box
# Param => width : width of the box
# Param => global_col : index of the col to start to write
# Param => global_row : index of the row to start to write
# Param => list : list to write
# Param => index : index of the element in the list to start
# Param => box_number : number of the box
def random_box(title, length, width, global_col, global_row, list, index, box_number):
    worksheet.set_column(global_col, global_col+width-1, find_longest_word(list))
    worksheet.merge_range(global_row-1,global_col,global_row-1, global_col+width-1, 'Merged Range')
    worksheet.write(global_row-1,global_col,"Plaque "+title + " " +str(box_number), title_cell)
    
    tab_coor = tab_2D_coordinate(length,width)
    
    if len(list) < len(tab_coor):
        shortest_list = len(list)
    else: 
        shortest_list = len(tab_coor)
    
    for x in range(shortest_list):
        if index >= len(list):
            break
        coor = random_coor(tab_coor)
        # print("Coor 0 : " + str(coor[0]+global_row))
        # print("Coor 1 : " + str(coor[1]+global_col))
        # print("Sample : " + list[index] +  "chess coor : "+ xl_rowcol_to_cell(coor[0]+global_row,coor[1]+global_col))

        # Add to dict the coordinate translate to chess coordinate
        sample_coordinate_random.update({list[index]: xl_rowcol_to_cell(coor[0]+global_row,coor[1]+global_col)})
        worksheet.write(coor[0]+global_row, coor[1]+global_col, list[index])
        index=index+1

    yellow_cell = workbook.add_format()
    yellow_cell.set_bg_color('yellow')
    dead_columns = ["Truc : "]
    for row in range(length):
        worksheet.write(global_row+row,width+global_col, dead_columns[0]+str(row),yellow_cell)

    return index
    
# Generate a 2D table containing is own index as element
def tab_2D_coordinate(x,y):
    res = []
    for i in range(x):
        for j in range(y):
            res.append([i,j])
    return res
    
# Will use box function until there are enought element in the list
# Param => length : length of the box
# Param => width : width of the box
# Param => global_col : index of the col to start to write
# Param => global_row : index of the row to start to write
# Param => list : list to write
def boxs(length,width,global_col,list, title):
    index = 0
    box_number = 0
    global_row = 1
    while index < len(list):
        index = box(length,width,global_col,global_row+box_number,list,index,str(box_number)+ " " +title)
        global_row = global_row + length
        box_number = box_number + 1
        
# Will use ranom_box function until there are enought element in the list
# Param => length : length of the box
# Param => width : width of the box
# Param => global_col : index of the col to start to write
# Param => global_row : index of the row to start to write
# Param => list : list to write
def random_boxes(title, length,width,global_col,list):
    index = 0
    box_number = 0
    global_row = 1

    while index < len(list):
        index = random_box(title, length,width,global_col, global_row+box_number,list,index,box_number)
        global_row = global_row+length
        box_number = box_number + 1
        
# Return a random element from the list and remove this element from the list
def random_coor(list):
    if list:
        r = random.randint(0,len(list)-1)
        return list.pop(r)
        
# Prefixe generator  = AfXXX
# Param => prefix : string, the prefix to write
# Param => number : int, the number to adapte
# Return => string : the correcte association between the prefix and the number
# Af001 / Af010
def prefix_count(prefix,number):
    if number < 10 :
        return prefix + "00" + str(number)
    if 10 <= number < 100:
        return prefix + "0"  + str(number)
    if 100 <= number < 1000:
        return prefix + str(number)
    if number <= 1000:
        return "Count > 1000 is not implement."

# Create the main list to be used in the program
def list_generator():
    for x in range(countMin,countMax+1):
    
        line = prefix_count(prefix,x)

        line = line+'-'+name+'-'+year+'-'+season
    
        sacs_zip.append(line)

        for y in tab1bis:
            sacs.append(line +'-' + y)

        for y in tab1:
            pilluliers.append(line + '-' + y)
            for z in tab2: 
                tubes.append(line + '-' + y + '-' + z)
                if z == "BS":
                    boite_BS.append(line + '-' + y + '-' + z)
                if z == "RH":
                    boite_RH.append(line + '-' + y + '-' + z)
                if z == "LF":
                    boite_LF.append(line + '-' + y + '-' + z)
                if z == "RO":
                    boite_RO.append(line + '-' + y + '-' + z)

                if y == "PA" or y == "PB":
                    tubes_PA_PB.append(line + '-' + y + '-' + z)
                    if z == "BS":
                        random_PAB_BS.append(line + '-' + y + '-' + z)
                    if z == "RH":
                        random_PAB_RH.append(line + '-' + y + '-' + z)
                    if z == "LF":
                        random_PAB_LF.append(line + '-' + y + '-' + z)
                    if z == "RO":
                        random_PAB_RO.append(line + '-' + y + '-' + z)
                # if y == "PC" or y == "PD":
                else:
                    tubes_PC_PD.append(line + '-' + y + '-' + z)
                    if z == "BS":
                        random_PCD_BS.append(line + '-' + y + '-' + z)
                    if z == "RH":
                        random_PCD_RH.append(line + '-' + y + '-' + z)
                    if z == "LF":
                        random_PCD_LF.append(line + '-' + y + '-' + z)
                    if z == "RO":
                        random_PCD_RO.append(line + '-' + y + '-' + z)
                



# Main
workbook = xlsxwriter.Workbook('projetB.xlsx')
worksheet = workbook.add_worksheet()


title_cell = workbook.add_format()
title_cell.set_bg_color('green')
title_cell.set_font_size(18)
title_cell.set_align('center')


def non_negative_input(output):
    while True:
        try:
            value = int(input(output))
        except ValueError:
            print("Ce n'est pas un nobmre.")
            continue

        if value < 0 or value >= 1000:
            print("Votre réponse doit être entre 0 et 1000.")
        elif not value: 
            print("Votre réponse ne doit pas être vide.")
        else:
            break
    return value

# Generate an input for user and check if it's a word and not empty
def non_empty_string_input(output):
    while True:
        try:
            value = str(input(output))
        except ValueError:
            print("Ce n'est pas une chaîne de caractère.")
            continue

        if not value:
            print("Votre réponse ne doit pas être vide.")
        elif value.isdigit():
            print("Votre réponse ne doit pas être une nombre.")
        else:
            break
    return value

# Main
workbook = xlsxwriter.Workbook('projetB.xlsx')
worksheet = workbook.add_worksheet()


title_cell = workbook.add_format()
title_cell.set_bg_color('green')
title_cell.set_font_size(18)
title_cell.set_align('center')


# print("Pour quitter l'opération en cours, faite CTRL+C.")
# prefix = non_empty_string_input("Veuillez entrer le prefix : ")
# countMin = non_negative_input("Veuillez entrer la borne min (>0) : ") 
# countMax = non_negative_input("Veuillez entrer la borne max (>0) : ")
# name = non_empty_string_input("Veuillez entrer le nom de votre truc (ex: Bn): ")
# year = non_empty_string_input("Veuillez entrer l'année (ex: Y21) : ")
# season = non_empty_string_input("Veuillez entrer la saison (ex: Au) : ")
# length_box= non_negative_input("Veuillez entrer la longueur de votre boite normal : ")
# width_box= non_negative_input("Veuillez entrer la largeur de votre boite normal : ")
# length_randombox= non_negative_input("Veuillez entrer la longueur de votre boite aléatoire : ")
# width_randombox= non_negative_input("Veuillez entrer la largeur de votre boite aléatoire : ")

prefix = "TestM"
countMin = 0
countMax = 25
name = "Bn"
year = "Y22"
season = "Sp"
length_box= 6
width_box= 6
length_randombox= 5
width_randombox= 5



# Créer les listes
list_generator()

# #Sacs column
# write_list_in_excel(0,"Sacs",sacs)

# # Sacs Zip column
# write_list_in_excel(1,"Sacs ZIP",sacs_zip)

# # Pilluliers
# write_list_in_excel(2,"Pilluliers",pilluliers)

# # Tubes 
# write_list_in_excel(3,"Tubes", tubes)


# # Box 
# boxs(length_box,width_box, 5, tubes, "")

# boxs(length_box,width_box, 7+width_box, boite_RO, "RO")
# boxs(length_box,width_box, 9+width_box*2, boite_BS, "BS")
# boxs(length_box,width_box, 11+width_box*3, boite_LF, "LF")
# boxs(length_box,width_box, 13+width_box*4, boite_RH, "RH")


# Random Box PA PB
#random_boxes("PA_PB",length_randombox,width_randombox, 15 + width_box*5, tubes_PA_PB)

# Random Box PC PD

#random_boxes("PC_PD",length_randombox,width_randombox, 17 + width_box*5+width_randombox, tubes_PC_PD)

#Random Box

# print("Nombre de tube = " + str(len(random_PAB_BS)))
random_boxes("PA_PB_BS",length_randombox,width_randombox, 15 + width_box*5+width_randombox*0, random_PAB_BS)
# random_boxes("PA_PB_RH",length_randombox,width_randombox, 17 + width_box*5+width_randombox*1, random_PAB_RH)
# random_boxes("PA_PB_LF",length_randombox,width_randombox, 19 + width_box*5+width_randombox*2, random_PAB_LF)
# random_boxes("PA_PB_RO",length_randombox,width_randombox, 21 + width_box*5+width_randombox*3, random_PAB_RO)
# random_boxes("PC_PD_BS",length_randombox,width_randombox, 23 + width_box*5+width_randombox*4, random_PCD_BS)
# random_boxes("PC_PD_RH",length_randombox,width_randombox, 25 + width_box*5+width_randombox*5, random_PCD_RH)
# random_boxes("PC_PD_LF",length_randombox,width_randombox, 27 + width_box*5+width_randombox*6, random_PCD_LF)
# random_boxes("PC_PD_RO",length_randombox,width_randombox, 29 + width_box*5+width_randombox*7, random_PCD_RO)
for i in sample_coordinate_random:
    print("Key : " +i +"| Value : " +sample_coordinate_random[i])

workbook.close()

print(datetime.datetime.now() - begin_time)
print("minute:second:microsecond")
print("Nombre de tube = " + str(len(tubes)))
