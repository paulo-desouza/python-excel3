from excel_methods import *  # NOQA

                        ###########################
                        ###   COMMENCE CODING   ###
                        ###########################



file_list = get_file_names()


                            ###################
                            #### TABLE ONE ####

table1_cols = ["Total", "Facebook", "FranchiseGator", "LinkedIn", "Website", "PPC", "BizBuySell", "franchise.com", "IFA"]
table1_rows = ["Week Total", "Day1", "Day2", "Day3", "Day4", "Day5", "Day6", "Day7"]

##############     WE ARE WAITING ON FRANCONNECT DATA FOR THIS PART OF THE REPORT     ########################
##############     WE ARE WAITING ON FRANCONNECT DATA FOR THIS PART OF THE REPORT     ########################
##############     WE ARE WAITING ON FRANCONNECT DATA FOR THIS PART OF THE REPORT     ########################
##############     WE ARE WAITING ON FRANCONNECT DATA FOR THIS PART OF THE REPORT     ########################
##############     WE ARE WAITING ON FRANCONNECT DATA FOR THIS PART OF THE REPORT     ########################


for file in file_list:
    if "daily lead count and source" in file:
        if "xlsx" in file:
            conv_file = file
        else:    
            conv_file = convert_xls(file)
            delete_file(file)


wb = load_workbook(conv_file)

ws = wb.active


# This has to be dynamic, count the amount of data entries (rows) before scraping.

rowcount = count_rows(ws)
    
table_1_dict = {}
    

for row in range(28, rowcount+24):
    table_1_dict[ws["B"+ str(row)].value] = ws["C"+ str(row)].value
    
# We need to add the logic to add 0s to the empty data entries. 


if "yesterdays_report.xlsx" in file_list:
    wb = load_workbook("yesterdays_report.xlsx")
    ws = wb.active
    
   
else:
    wb = Workbook()  
    ws = wb.active
    
    


# Title

ws.merge_cells("B2:K3")
ws["B2"].value = "DAILY LEAD AND SOURCE"



# We are able to get the data and output it correctly. 
# Now we need the logic for it to be recursive. 

# 1 - Before writing the current day table, check for data in the
# previous days, and determine what day of the week we are currently in.

# 2 - Knowing what line we are writing on, we can sum this line's values 
# with the values of the previous days, and output it in the "week total"
# row. 

# For this logic to work, we have to know if this is the first report of the week or not


weekday = None

if ws["C7"].value == None:
    weekday = (1, "C7")
    
elif ws["C8"].value == None:
    weekday = (2, "C8")
    
elif ws["C9"].value == None:
    weekday = (3, "C9")
    
elif ws["C10"].value == None:
    weekday = (4, "C10")
    
elif ws["C11"].value == None:
    weekday = (5, "C11")
    
elif ws["C12"].value == None:
    weekday = (6, "C12")
    
elif ws["C13"].value == None:
    weekday = (7, "C13")



    
    

write_table(ws, table1_rows,              
            "B6:B"+str(len(table1_rows)+5),
            table1_cols,
            "C5:"+get_column_letter(len(table1_cols)+2)+"5")


# write_table(ws, table_1_data, weekday[1] + ":" + get_column_letter(len(table_1_data)) + str(weekday[0]+6)) 





wb.save("output.xlsx")


##############     WE ARE WAITING ON FRANCONNECT DATA FOR THIS PART OF THE REPORT     ########################
##############     WE ARE WAITING ON FRANCONNECT DATA FOR THIS PART OF THE REPORT     ########################
##############     WE ARE WAITING ON FRANCONNECT DATA FOR THIS PART OF THE REPORT     ########################
##############     WE ARE WAITING ON FRANCONNECT DATA FOR THIS PART OF THE REPORT     ########################


                    #### TABLE ONE FINISH ####
                    ##########################
                    
                    
                    
                    
                        ###################
                        #### TABLE TWO ####


table2_cols = ["Total Leads", "<1 Hour", "1-2 hours", "2-3 hours", "3-4 hours", "4-5 hours", "5+ hours"]
table2_rows = ["Day1", "Day2", "Day3", "Day4", "Day5", "Day6", "Day7"]
# rows are defined by the dates



for file in file_list:
    if "daily inquiry response time" in file:
        if "xlsx" in file:
            conv_file = file
        else:    
            conv_file = convert_xls(file)
            delete_file(file)

# load the to-be-scraped xlsx file

wb = load_workbook(conv_file)

ws = wb.active


rowcount = count_rows(ws)

table_2_raw = scrape_table(ws, "B4:B" + str(rowcount+1))

# convert table2 raw data into time values, compare them to the time values of
# each column, and make a list with all of the numbers to be written. 




table_2_data = {
        "Total": 0,
        "<1 Hour":   0,
        "1-2 Hours": 0,
        "2-3 Hours": 0,
        "3-4 Hours": 0,
        "4-5 Hours": 0,
        "5+ Hours":  0
    }


for item in table_2_raw:
    if "Minute" in item or "Seconds" in item:
        table_2_data["<1 Hour"] += 1
        table_2_data["Total"] += 1
    
    elif "Hour" in item:
        
        if item[1].isnumeric():
        # has 2 digits        
            table_2_data["5+ Hours"] += 1
            table_2_data["Total"] += 1
            
            
        else:
            if int(item[0]) <= 2:
                table_2_data["1-2 Hours"] += 1
                table_2_data["Total"] += 1
            
            elif int(item[0]) <= 3:
                table_2_data["2-3 Hours"] += 1
                table_2_data["Total"] += 1
            
            
            elif int(item[0]) <= 4:
                table_2_data["3-4 Hours"] += 1
                table_2_data["Total"] += 1
                
            elif int(item[0]) <= 5:
                table_2_data["4-5 Hours"] += 1
                table_2_data["Total"] += 1
            
    
    elif "Day" in item:    
        table_2_data["5+ Hours"] += 1
        table_2_data["Total"] += 1
    


# load our to-be-written xlsx file

wb = load_workbook("output.xlsx")

ws = wb.active


weekday = None

if ws["Q6"].value == None:
    weekday = (1, "Q6")
    
elif ws["Q7"].value == None:
    weekday = (2, "Q7")
    
elif ws["Q8"].value == None:
    weekday = (3, "Q8")
    
elif ws["Q9"].value == None:
    weekday = (4, "Q9")
    
elif ws["Q10"].value == None:
    weekday = (5, "Q10")
    
elif ws["Q11"].value == None:
    weekday = (6, "Q11")
    
elif ws["Q12"].value == None:
    weekday = (7, "Q12")


# Title

ws.merge_cells("P2:W3")
ws["P2"].value = "DAILY INQUIRY RESPONSE TIME"



write_table(ws, 
            table2_rows, "P6:P"+str(len(table2_rows)+5),
            table2_cols, "Q5:"+get_column_letter(len(table2_cols)+16)+"5")

table_2_list=list(table_2_data.values())


write_table(ws, table_2_list, str(weekday[1])+":W"+str(weekday[1][1:]))


wb.save("output.xlsx")


                ## TABLE 2 FINISH ##
                ####################
                
                
                
                
                ####################
                ## TABLE  3 BEGIN ##

table3_cols = ["Goal", "Week3", "Week2", "Week1"]
table3_rows = ["Leads", "Connected", "Intro Call Scheduled"]
table3_goals = ["50", "60%", "20%"]

# scrape the 3 week files

table_3_data = []
j = 0
for file in file_list:
    if "last 3 weeks" in file:
    
        if "xlsx" in file:
            conv_file = file
        else:    
            conv_file = convert_xls(file)
            delete_file(file)
        
                
        wb = load_workbook(conv_file)

        ws = wb.active

        rowcount = count_rows(ws)
        
        table_3_data.append(scrape_table(ws, "c5:f5"))
        table_3_data[j].pop(1)
        
     
        
        j += 1
     
wb = load_workbook("output.xlsx")

ws = wb.active

# Title

ws.merge_cells("R16:U17")
ws["R16"].value = "Last 3 weeks' connected & scheduled rates"

write_table(ws, 
            table3_rows, "Q19:Q"+str(len(table3_rows)+18),
            table3_cols, "R18:"+get_column_letter(len(table3_cols)+17)+"18")
write_table(ws, table3_goals, "R19:R21")

write_table(ws, table_3_data[0], "S19:S21")
            
write_table(ws, table_3_data[1], "T19:T21")

write_table(ws, table_3_data[2], "U19:U21")



wb.save("output.xlsx")

                ## TABLE 3 FINISH ##
                ####################
                
                
                
                ####################
                ## TABLE 4 START  ##
                
                
                
                
for file in file_list:
    if "currentpipelinestatus" in file.lower():
        
    
        if "xlsx" in file:
            conv_file = file
            
        else:    
            conv_file = convert_xls(file)
            delete_file(file)
            
        
        wb = load_workbook(conv_file)
        
        ws = wb.active

#    A2: rowcount + colcount

colcount4 = count_cols(ws, char=False)
col_char4 = count_cols(ws)

rowcount4 = count_rows(ws)


scraped_4 = scrape_table(ws, "A2:"+col_char4+str(rowcount4))
        

     
wb = load_workbook("output.xlsx")

ws = wb.active


                # Title
ws['M60'].value = "Current Pipeline Status"
ws.merge_cells("M60:W61")


write_table(ws, scraped_4, "M63:"+get_column_letter(colcount4+12)+str(rowcount4+61))


wb.save("output.xlsx")


                
                ## TABLE 4 FINISH ##
                ####################
                
                
                
                
                ####################
                ## TABLE 5 START  ##
                
table5_cols = ['Total', 'Intro Call Scheduled', 'Intro Call Completed', 'Intro Call Not Scheduled',
               'Bad Contact Info', 'Insufficient Capital', 'International', 'Market Not Available',
               "Accidentally Submitted", "Looking for Childcare", "Looking for Employment",
               "Not Interested."]    


# scrape here

data_3pt1 = []
data_3pt2 = []

j = 0
for file in file_list:
    if "rolling 7 day inquiry" in file:
        
    
        if "xlsx" in file:
            conv_file = file
            
        else:    
            conv_file = convert_xls(file)
            delete_file(file)
            
        
        wb = load_workbook(conv_file)
        
        ws = wb.active
            
        if "pt 1" in file:
            data_3pt1.append(scrape_table(ws, "C5:H5"))
            data_3pt1[0].pop(1)
            data_3pt1[0].pop(4)
            
            
        elif "pt 2" in file:
            data_3pt2.append(scrape_table(ws, "B43:C" + str(43 + (count_rows(ws)-5))))
            

     
wb = load_workbook("output.xlsx")

ws = wb.active



# Title

ws.merge_cells("Q27:V28")
ws["Q27"].value = "Rolling 7 Day Inquiry Lead Status"


write_table(ws, table5_cols, "N30:"+get_column_letter(len(table5_cols)+13)+"30")

            
# write scraped table here






values_dict = {
    'Bad Contact Info':  0,
    'Insufficient Capital': 0,
    'International Interest Only': 0,
    'Market Not Currently Available': 0,
    'Accidentally Submitted Inquiry': 0,
    'Looking for Childcare, not Franchise': 0,
    'Looking for employment': 0,
    'Not interested': 0
    }

for i, item in enumerate(data_3pt2[0]):
    if i%2 == 0:
       values_dict[item] += data_3pt2[0][i+1]
       

ws["N31"].value = data_3pt1[0][0]
ws["O31"].value = data_3pt1[0][2]
ws["P31"].value = data_3pt1[0][3]
ws["Q31"].value = data_3pt1[0][1]

table_5_list=list(values_dict.values())


write_table(ws, table_5_list, "R31:Y31")

          
wb.save("output.xlsx")


                ## TABLE 5 FINISH ##
                ####################
                
                
                
                ####################
                ## TABLE 6 START  ##
                
table6_rows = ["Day1", "Day2", "Day3", "Day4", "Day5", "Day6", "Day7"]
  # to be pulled from excel sheet




j = 0
for file in file_list:
    if "daily # of intro calls" in file:
        
    
        if "xlsx" in file:
            conv_file = file
            
        else:    
            conv_file = convert_xls(file)
            delete_file(file)
            
        
        wb = load_workbook(conv_file)
        
        ws = wb.active


rowcount = count_rows(ws)

table6_cols = scrape_table(ws, "A6:A"+str(rowcount+1))   #salespeople

salespeople = []

for person in table6_cols:
    salespeople.append(person.split(" ")[0])
    

data_6 = scrape_table(ws, "F6:F"+str(rowcount+1))





     
wb = load_workbook("output.xlsx")



ws = wb.active


weekday = None

if ws["R50"].value == None:
    weekday = (1, "R50")
    
elif ws["R51"].value == None:
    weekday = (2, "R51")
    
elif ws["R52"].value == None:
    weekday = (3, "R52")
    
elif ws["R53"].value == None:
    weekday = (4, "R53")
    
elif ws["R54"].value == None:
    weekday = (5, "R54")
    
elif ws["R55"].value == None:
    weekday = (6, "R55")
    
elif ws["R56"].value == None:
    weekday = (7, "R56")


# Title

ws.merge_cells("R46:U47")
ws["R46"].value = "Daily # of Intro Calls Scheduled"




write_table(ws, table6_rows, "Q50:Q"+str(len(table6_rows)+49),
            salespeople, "R49:"+get_column_letter(len(salespeople)+17)+"49")

            
# write scraped table here

write_table(ws, data_6, weekday[1]+":U"+weekday[1][1:])

wb.save("output.xlsx")
                
                ## TABLE 6 FINISH ##
                ####################
                
                
                
                
                ####################
                ## TABLE 7 START  ##
                

table7_cols = ["Goal Actual", "Goal %", "Actual", "Actual %"]
table7_rows = ["Leads", "Connected", "Intro Call", "Business Overview Webinar",
               "Operations & Marketing Webinar", "FDD Review", "Competency Call",
               "Executive Call", "Meet the Team Scheduled", "Decision Day",
               "Awarded"]


table7_goals = [[200,120,40,30,20,12,8,4,4,4,2],
                ["100%", "60%", "20%", "15%", "10%", "6%",
                 "4%", "2%", "2%", "2%", "1%"]]


j = 0
for file in file_list:
    if "rolling 30 day new inquiry funnel" in file:
        
    
        if "xlsx" in file:
            conv_file = file
            
        else:    
            conv_file = convert_xls(file)
            delete_file(file)
            
        
        wb = load_workbook(conv_file)
        
        ws = wb.active
        
        
table_7_data = scrape_table(ws, "C5:P5")

# POPPIN

table_7_data.pop(1)
table_7_data.pop(3)
table_7_data.pop(10)




new_7 = []

for item in table_7_data:
    if item == "0":
        new_7.append(item)
        new_7.append(item+"%")
    else:
        new_7.append(item.split("(")[0])
        new_7.append(item.split("(")[1].replace(")",""))



wb = load_workbook("output.xlsx")

ws = wb.active


# Title

ws.merge_cells("E60:H61")
ws["E60"].value = "Rolling 30 Day New Inquiry Funnel"



write_table(ws, table7_rows, "D65:D"+str(len(table7_rows)+64),
            table7_cols, "E64:"+get_column_letter(len(table7_rows)+4)+"64")

            
# write scraped table here

write_table(ws, new_7, "G65:H75")

write_table(ws, table7_goals[0], "E65:E75")

write_table(ws, table7_goals[1], "F65:F75")

          
wb.save("output.xlsx")

                
                ## TABLE 7 FINISH ##
                ####################
                
                ####################
                ## TABLE 8 START  ##

table8_rows = ["Welcome", "EBITDA", "Support from the top down",
               "Day in the Life w/ Katie Young", "Still Interested?"]

table8_cols = ["Email", "# Read"]


j = 0
for file in file_list:
    if "emails read " in file:
        
    
        if "xlsx" in file:
            conv_file = file
            
        else:    
            conv_file = convert_xls(file)
            delete_file(file)
            
        
        wb = load_workbook(conv_file)
        
        ws = wb.active



scraped_info = []

rowcount = count_rows(ws)

for row in range(4, rowcount+2):

    scraped_info.append([ws["B"+str(row)].value, ws["G"+str(row)].value, ws["C"+str(row)].value])

    

prospects = []
email_count = {
        "EBITDA":0,
        "a day in the life": 0,
        "from the top down": 0,
        "welcome to the celebree school franchise":0,
        "still interested": 0
    }
    
for email in scraped_info:
    if "EBITDA" in email[0] and len(email[1]) > 3:
        
        email_count['EBITDA'] += 1
        
    elif "a day in the life" in email[0].lower() and len(email[1]) > 3:
        email_count['a day in the life'] += 1
    
    elif "from the top down" in email[0].lower() and len(email[1]) > 3:
        email_count['from the top down'] += 1
    
    elif "welcome to the celebree school franchise" in email[0].lower() and len(email[1]) > 3:
        email_count['welcome to the celebree school franchise'] += 1
    
    elif "still interested" in email[0] and len(email[1]) > 3:
        email_count['still interested'] += 1
        
    if email[2] not in prospects:
        prospects.append(email[2])
    



wb = load_workbook("output.xlsx")

ws = wb.active


# Title

ws.merge_cells("F46:I47")
ws["F46"].value = "Emails Read from 1/22  -  1/28 (weekly)"


write_table(ws, table8_rows, "E50:e54",
            table8_cols, "E49:F49")

            
# write scraped table here
ws.merge_cells("E55:F55")
ws["E55"].value = "Out of "+ str(len(prospects))+ " leads"


j = 0
#for i, row in enumerate():
 #   for cell in row:
 #       cell.value = 
  #      j += 1



          
wb.save("output.xlsx")

                
                ## TABLE 8 FINISH ##
                ####################
                
                ####################
                ## TABLE 9 START  ##
                

table9_rows = ["Leads", "Connected", "Intro Call", "Business Overview Webinar",
               "Operations & Marketing Webinar", "FDD Review", "Competency Call",
               "Executive Call", "Meet the Team Scheduled", "Decision Day",
               "Awarded"]
table9_cols = ["Goal Actual", "Goal %", "Actual", "Actual %"]


for person in salespeople:
    table9_cols.append(person+" Actual")
    table9_cols.append(person+" %")

table9_goals = [[200,120,40,30,20,12,8,4,4,4,2],
                ["100%", "60%", "20%", "15%", "10%", "6%",
                 "4%", "2%", "2%", "2%", "1%"]]

j = 0
for file in file_list:
    if "rolling 30 day activity funnel" in file:
        
    
        if "xlsx" in file:
            conv_file = file
            
        else:    
            conv_file = convert_xls(file)
            delete_file(file)
            
        
        wb = load_workbook(conv_file)
        
        ws = wb.active

table_9_raw = []

rowcount = count_rows(ws)


for row in range(5, rowcount+2):
    table_9_raw.append(scrape_table(ws, "C"+ str(row) +":P"+ str(row)))

# POPPIN

for table in table_9_raw:
    table.pop(1)
    table.pop(3)
    table.pop(10)

table_9_data = []

for k, item in enumerate(table_9_raw):
    for j, i in enumerate(item):
        
        if j == 0:
            new_leads = i.split("(")[0]
            table_9_data.append([])
            
            
        table_9_data[k].append(i.split("(")[0])
        
        perc_float = (int(i.split("(")[0])*100)/int(new_leads)
        table_9_data[k].append(f'{perc_float:.0f}%')
    
     
wb = load_workbook("output.xlsx")

ws = wb.active


# Title

ws.merge_cells("F79:T80")
ws["F79"].value = "Rolling 30 Day Activity Funnel (All inquiry date)"



write_table(ws, table9_rows, "F83:F"+str(len(table9_rows)+82),
            table9_cols, "G82:"+get_column_letter(len(table9_cols)+6)+"82")

            
# write scraped table here

write_table(ws, table7_goals[0], "G83:G93")

write_table(ws, table7_goals[1], "H83:H93")


i = 0
for j, table in enumerate(table_9_data):
    write_table(ws, table_9_data[j], get_column_letter(i+9)+"83:"+ get_column_letter(i+10) +"93")
    i += 2

                
                ## TABLE 9 FINISH ##
                ####################
                
                
                
                
                ####################
                ## SHHHHTYLE   !  ##

                
# Base Style for ALL CELLS
    # ADD Conditional statements for the rest of the configs?
    # either this or make new methods to apply styles based on received range

for row in range(1, 100):
    ws.row_dimensions[row].height = 32
    
    for col in range(1, 60):
        char = get_column_letter(col)
    
        if row == 1:
            ws.column_dimensions[char].width = 17
        
        
        ws[char+str(row)].font = Font(size = 15)
        ws[char+str(row)].alignment = Alignment(horizontal = "center", vertical = "center", wrap_text=True)
        
# TITLES

ws["F79"].font = Font(size = 22, bold = True)
ws["F46"].font = Font(size = 22, bold = True)
ws["E60"].font = Font(size = 22, bold = True)
ws["R46"].font = Font(size = 22, bold = True)
ws["Q27"].font = Font(size = 22, bold = True)
ws["B27"].font = Font(size = 22, bold = True)
ws["R16"].font = Font(size = 22, bold = True)
ws["P2"].font = Font(size = 22, bold = True)
ws["B2"].font = Font(size = 22, bold = True)
ws["M60"].font = Font(size = 22, bold = True)


# Table Header Cols and Rows

# 1

# bold
mod_font(ws,"C5:K5")
mod_font(ws,"B6:B13")

#border
set_border(ws, "B5:K13")
set_border(ws, "B2:K3", _medium=True)

# color
mod_color(ws, "C5:K5")
mod_color(ws, "B6:B13")
mod_color(ws, "B2:K3")
                
# 2

mod_font(ws,"Q5:W5")
mod_font(ws,"P6:P12")

set_border(ws, "P5:W12")
set_border(ws, "P2:W3", _medium=True)

mod_color(ws, "P2:W3")
mod_color(ws, "Q5:W5")
mod_color(ws, "P6:P12")      
 
# 3
         
mod_font(ws,"R18:U18")
mod_font(ws,"Q19:Q21")

set_border(ws, "Q18:U21")
set_border(ws, "R16:U17", _medium=True)

mod_color(ws, "R16:U17")
mod_color(ws, "R18:U18")
mod_color(ws, "Q19:Q21") 

# 4

mod_font(ws,"N63:R63")
mod_font(ws,"M64:M76")

set_border(ws, "M63:R76")
set_border(ws, "M60:W61", _medium=True)

mod_color(ws, "N63:R63")
mod_color(ws, "M64:M76")
mod_color(ws, "M60:W61") 
                
# 5

ws["O37"].value = 'Connected Rate:'

mod_font(ws,"N30:Y30")
mod_font(ws,"O37:P37")

set_border(ws, "N30:Y31")
set_border(ws, "Q27:V28", _medium=True)
                
set_border(ws,"O37:Q37")
ws.merge_cells("O37:P37")

mod_color(ws, "Q27:V28")
mod_color(ws, "N30:Y30")
mod_color(ws, "O37:P37") 

# 6

#  dynamic

mod_font(ws,"R49:U49") ##
mod_font(ws,"Q50:Q56")

set_border(ws, "Q49:U56")  ##
set_border(ws, "R46:U47", _medium=True)

mod_color(ws, "R46:U47")
mod_color(ws, "R49:U49") ##
mod_color(ws, "Q50:Q56") 

# 7

mod_font(ws,"E64:H64")
mod_font(ws,"D65:D75")

set_border(ws, "D64:H75")
set_border(ws, "E60:H61", _medium=True)

mod_color(ws, "E60:H61")
mod_color(ws, "E64:H64")
mod_color(ws, "D65:D75")
       
# 8

mod_font(ws, "E49:F49")
mod_font(ws, "E50:E54")
mod_font(ws, "E55:F55")

set_border(ws, "E49:F54")
set_border(ws, "F46:I47", _medium=True)
set_border(ws, "E55:F55", _medium=True)

mod_color(ws, "F46:I47")
mod_color(ws, "E49:F49")
mod_color(ws, "E50:E54")


# 9

# dynamic

mod_font(ws,"G82:R82") ##
mod_font(ws,"F83:F93")

set_border(ws, "F82:R93") ##
set_border(ws, "F79:T80", _medium=True)

mod_color(ws, "F79:T80")
mod_color(ws, "G82:R82") ##
mod_color(ws, "F83:F93")
       

### CELL SPACING MODS

# LIST WITH 2 LINERS

twoliners = [5,21,30,72,73,75,89,90,92,82]

# LIST WITH 3 LINERS
threeliners = [71,70,69,74,86,87,91]


for row in twoliners:
    ws.row_dimensions[row].height = 43

for row in threeliners:
    ws.row_dimensions[row].height = 68

ws.column_dimensions["J"].width = 18


ws.sheet_view.zoomScale = 60


wb.save("output.xlsx")                