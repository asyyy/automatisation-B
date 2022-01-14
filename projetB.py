from curses.ascii import isdigit
import datetime
import random
import xlsxwriter


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

# FUNCTIONS (TODO les commenter parce que la c'est chaud)

def find_longest_word(list):
    max_length = 0
    for x in list:
        if len(x)>max_length:
            max_length = len(x)
    return max_length

def write_list_in_excel(col,title,list):
    worksheet.set_column(col, col, find_longest_word(list))
    worksheet.write(0,col,title,title_cell)
    for row in range(len(list)):
        worksheet.write(row+1, col, list[row])

def box(length,width,global_col,global_row,list,index,box_number):
    worksheet.set_column(global_col, global_col+width-1, find_longest_word(list))
    worksheet.merge_range(global_row-1,global_col,global_row-1, global_col+width-1, 'Merged Range')
    worksheet.write(global_row-1,global_col,"Box " +str(box_number),title_cell)
    

    for row in range(0,length):
        for col in range(0,width):
            if index < len(list):
                worksheet.write(row+global_row, col+global_col, list[index])
                index = index+1
    return index

def boxs(length,width,global_col,list):
    index = 0
    box_number = 0
    global_row = 1
    while index < len(list):
        index = box(length,width,global_col,global_row,list,index,box_number)
        global_row = global_row + length
        box_number = box_number + 1

def random_coor(list):
    if list:
        r = random.randint(0,len(list)-1)
        return list.pop(r)

def random_box(title, length, width, global_col, list):
    worksheet.set_column(global_col, global_col+width, find_longest_word(list))
    worksheet.merge_range(0,global_col,0, global_col+width, 'Merged Range')
    worksheet.write(0,global_col,"Box Random "+title, title_cell)
    random_box_coordinate = []

    for i in range(0,length):
        for j in range(0,width):
            random_box_coordinate.append([i,j])
   
    for x in range(len(random_box_coordinate)):
        
        coor = random_coor(random_box_coordinate)
        worksheet.write(coor[0]+1, coor[1]+global_col, list[x])

    yellow_cell = workbook.add_format()
    yellow_cell.set_bg_color('yellow')
    dead_columns = ["Truc : "]
    for row in range(length):
        worksheet.write(row+1,width+global_col, dead_columns[0]+str(row),yellow_cell)
# Prefixe generator  = AfXXX
def prefix_count(prefix,index):
    if index < 10 :
        return prefix + "00" + str(index)
    if 10 <= index < 100:
        return prefix + "0"  + str(index)
    if 100 <= index < 1000:
        return prefix + str(index)

# List generator (TODO A mettre dans une fonction)
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
                if y == "PA" or y == "PB":
                    tubes_PA_PB.append(line + '-' + y + '-' + z)
                # if y == "PC" or y == "PD":
                else:
                    tubes_PC_PD.append(line + '-' + y + '-' + z)


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
            print("Sorry, I didn't understand that.")
            continue

        if value < 0:
            print("Sorry, your response must not be negative.")
        else:
            break
    return value

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

#prefix = digit_input("Prefix: ")
## use check_is_digit()
print("Pour quitter l'opération en cours, faite CTRL+C.")
prefix = non_empty_string_input("Veuillez entrer le prefix : ")
countMin = non_negative_input("Veuillez entrer la borne min (>0) : ") 
countMax = non_negative_input("Veuillez entrer la borne max (>0) : ")
name = non_empty_string_input("Veuillez entrer le nom de votre truc (ex: Bn): ")
year = non_empty_string_input("Veuillez entrer l'année (ex: Y21) : ")
season = non_empty_string_input("Veuillez entrer la saison (ex: Au) : ")

tab1 = ["PA","PB","PC","PD"]
tab1bis = ["PA","PB","PC","PD","Culturomique"]
tab2 = ["BS","RH","LF","RO"]
length_box= non_negative_input("Veuillez entrer la longueur de votre boite : ")
width_box= non_negative_input("Veuillez entrer la largeur de votre boite : ")
length_randombox= length_box
width_randombox= width_box

# Créer les listes
list_generator()

# Sacs column
write_list_in_excel(0,"Sacs",sacs)

# Sacs Zip column
write_list_in_excel(1,"Sacs ZIP",sacs_zip)

# Pilluliers
write_list_in_excel(2,"Pilluliers",pilluliers)

# Tubes 
write_list_in_excel(3,"Tubes", tubes)


# Box 
boxs(length_box,width_box, 5, tubes)

# Random Box PA PB

random_box("PA_PB",length_randombox,width_randombox, 7 + width_box, tubes_PA_PB)

# Random Box PC PD

random_box("PC_PD",length_randombox,width_randombox, 7 + width_box+width_randombox+2, tubes_PC_PD)

#Random Box

workbook.close()

print(datetime.datetime.now() - begin_time)
print("minute:second:microsecond")
print("Nombre de tube = " + str(len(tubes)))
