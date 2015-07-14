                                            ##########################
                                            # Certificates Generator # 
                                            #       By- sd5869       #
                                            ##########################

#####################
# Importing Modules #
#####################

from PIL import ImageFont, ImageDraw, Image
from openpyxl import load_workbook

###############################################################
# The Dictionary storing all information about each parameter #
###############################################################

parameter_info=dict()

##################################################################
# The function that is used for reading data from data.xlsx file #
##################################################################

def get_data(Start,End):
    wb = load_workbook(filename = 'data.xlsx')
    sheet= wb['Sheet1']
    A=[]
    t=[]
    for rowOfCellObjects in sheet[Start:End]:
        for cellObj in rowOfCellObjects:
            t.append(cellObj.value)
        A.append(t)
        t=[]
    return A

#########################################################
# The function that is used for generating certificates #
#########################################################

def generate_certi(img,data,parameter_info):
        for info in data:
                im = Image.open(img)
                for x in range(len(info)):
                        [coordinates,font,size,color]=parameter_info[x]
                        pilfont = ImageFont.truetype(font,size)
                        draw = ImageDraw.Draw(im)
                        draw.text(coordinates,info[x],color,font=pilfont)
                im.save("certificates/"+info[0]+".png")
        print("Successfully generated all certificates")

#####################################################################################################
# Getting basic information i.e name of template and number of parameters to write in a certificate #
#####################################################################################################

print("Enter Certificate Template name with extension")
img=input()
print("Enter number of parameters to write on the certificate")
num_par=int(input())

######################################################################################
# This Loop Handles all the input part and collects the details about each parameter #
######################################################################################

for x in range(num_par):
    data=list()
    print("Enter x and y coordinate of parameter "+str(x+1)+" in the template sperate coordinates by Comma")
    coordinates=tuple(map(int,input().split(",")))
    print("Enter the name of font to use for parameter "+str(x+1)+" NOTE: only *.ttf fonts are allowed,to use default font enter NULL")
    font=input()
    if(font=="NULL"):
        font="micross.ttf"
    print("Enter size of the font to use")
    size=int(input())
    print("Enter decimal value of colour of the font to use seperate RGB values by Comma")
    color=tuple(map(int,input().split(",")))

    #######################################
    # Packing all details into dictionary #
    #######################################

    parameter_info[x]=[coordinates,font,size,color]
print("Are details of each parameter present in data.xlsx file y/n")
ans=input()
if(ans=="y"):

    ###########################################
    # Extracting all data from the excel file #
    ###########################################
        
    print("Enter address of starting cell")
    Start=input()
    print("Enter address of ending cell")
    End=input()
    data=get_data(Start,End)
else:
    print("Write data in excel file then restart the program")

###################################
# Generating all the Certificates #
###################################

generate_certi(img,data,parameter_info)
                                            ##########################################
                                            # For any bugs or feature request please #
                                            #       Contact: sd5869@gmail.com        #
                                            ##########################################
