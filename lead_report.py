from excel_methods import *  # NOQA

                        ###########################
                        ###   COMMENCE CODING   ###
                        ###########################


file_list = get_file_names()


#FROM file names, get TODAY'S DATE


                            ###################
                            #### TABLE ONE ####

table1_cols = ["Total", "Facebook", "FranchiseGator", "LinkedIn", "Website", "PPC", "BizBuySell", "franchise.com", "IFA"]


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
    
table_1_dict = {
        "Total": 0,
        "Facebook": 0,
        "FranchiseGator": 0,
        "LinkedIn InMail": 0, 
        "Website": 0,
        "PPC": 0, 
        "BizBuySell": 0,
        "franchise.com": 0,
        "IFA": 0
}
    
for key in table_1_dict.keys():
    for row in range(27, rowcount+24):
        if key == ws["B"+str(row)].value:
            table_1_dict[key] = ws["D"+str(row)].value
    
    
# We need to add the logic to add 0s to the empty data entries. 


if "yesterdays_report.xlsx" in file_list:
    wb = load_workbook("yesterdays_report.xlsx")
    ws = wb.active
    
   
else:
    wb = Workbook()  
    ws = wb.active
    
    # FROM file names, get TODAY'S DATE
    for file in file_list:
        if 'daily lead count and source' in file: 
            
            current_day = file.split(" ")[0]
            
            current_day = current_day.split('.')
            
            current_year = datetime.now().year

            current_date = datetime.strptime(f'{current_year}-{current_day[1]}-{current_day[0]}', '%Y-%d-%m')
            
            first = current_date.strftime("%d/%b")
            
            seven_weekdays = [current_date.strftime("%d-%b")]
            
            for i in range(0,6):
                time_change = timedelta(hours=24)
                current_date = current_date + time_change
                seven_weekdays.append(current_date.strftime("%d-%b"))
                if i == 5:
                    last = current_date.strftime("%d/%b")
                    
            current_week = f'WEEK {first.upper()} - {last.upper()}'
            
            
                
table1_rows = seven_weekdays[:]
table1_rows.insert(0, "Week Total")
    
            
    


# Title

ws.merge_cells("B2:K3")
ws["B2"].value = f"DAILY LEAD AND SOURCE | {current_week}"



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


table_1_data = list(table_1_dict.values())


write_table(ws, table_1_data, weekday[1] + ":" + get_column_letter(len(table_1_data)+2) + str(weekday[0]+6)) 


## GIVE ME THE WEEK TOTAL SO FAR::::

for column in range(3, 12):
    sum = 0
    char = get_column_letter(column)
    
    for row in range(7, 14):
        if ws[char+str(row)].value == None:
            break
        
        sum += int(ws[char+str(row)].value)

    
    ws[char+str(6)].value = sum

### DRAW CHART

pie = PieChart()                                # Initiate the chart; 

                                                # Declare label coordinates
                                                # min_col, min_row, max_row
                                                # Declare data coordinates

labels = Reference(ws, min_row=5, min_col=4, max_col=11) 
data = Reference(ws, min_col=4, min_row=6, max_col=11)   

pie.add_data(data, from_rows=6, titles_from_data=False)       # Add data to pie chart. 
pie.set_categories(labels)                                   # Add labels to pie chart. 
pie.title = "Daily Lead and Source"





ws.add_chart(pie, "C15")

wb.save("output.xlsx")

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

if ws["O6"].value == None:
    weekday = (1, "O6")
    
elif ws["O7"].value == None:
    weekday = (2, "O7")
    
elif ws["O8"].value == None:
    weekday = (3, "O8")
    
elif ws["O9"].value == None:
    weekday = (4, "O9")
    
elif ws["O10"].value == None:
    weekday = (5, "O10")
    
elif ws["O11"].value == None:
    weekday = (6, "O11")
    
elif ws["O12"].value == None:
    weekday = (7, "O12")


# Title

ws.merge_cells("N2:U3")
ws["N2"].value = f"DAILY INQUIRY RESPONSE TIME | {current_week}"



write_table(ws, 
            seven_weekdays, "N6:N"+str(len(table2_rows)+5),
            table2_cols, "O5:"+get_column_letter(len(table2_cols)+15)+"5")

table_2_list=list(table_2_data.values())


write_table(ws, table_2_list, str(weekday[1])+":U"+str(weekday[1][1:]))


wb.save("output.xlsx")


                ## TABLE 2 FINISH ##
                ####################
                
                
                
                
                ####################
                ## TABLE  3 BEGIN ##


table3_rows = ["Leads", "Connected", "Intro Call Scheduled"]
table3_goals = ["50", "60%", "20%"]

# scrape the 3 week files

table_3_data = []
dates_3 = []
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
        dates_3.append(ws["A1"].value.split("(")[1][:-1].split(" - "))
        
        j += 1
        
table3_cols = ["Goal", dates_3[0][0][:-5] + ' - ' + dates_3[0][1][:-5], dates_3[1][0][:-5] + ' - ' + dates_3[1][1][:-5], dates_3[2][0][:-5] + ' - ' + dates_3[2][1][:-5]]

wb = load_workbook("output.xlsx")

ws = wb.active

# Title

ws.merge_cells("I35:L36")
ws["I35"].value = "Last 3 weeks' connected & scheduled rates"

write_table(ws, 
            table3_rows, "H39:H41",
            table3_cols, "I38:"+get_column_letter(len(table3_cols)+8)+"38")
write_table(ws, table3_goals, "I39:I41")

write_table(ws, table_3_data[0], "J39:J41")
            
write_table(ws, table_3_data[1], "K39:K41")

write_table(ws, table_3_data[2], "L39:L41")



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
ws['K49'].value = "Current Pipeline Status"
ws.merge_cells("K49:U50")


write_table(ws, scraped_4, "K52:"+get_column_letter(colcount4+10)+str(rowcount4+50))


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

ws.merge_cells("O19:T20")
ws["O19"].value = "Rolling 7 Day Inquiry Lead Status"


write_table(ws, table5_cols, "L22:W22")

            
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
       

ws["L23"].value = data_3pt1[0][0]
ws["M23"].value = "dontknow"
ws["N23"].value = data_3pt1[0][2]
ws["O23"].value = data_3pt1[0][1]

table_5_list=list(values_dict.values())


write_table(ws, table_5_list, "P23:W23")



ws["M29"].value = 'Connected Rate:'
ws["O29"].value = data_3pt1[0][2].split("(")[1][:-1]

          
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

if ws["P39"].value == None:
    weekday = (1, "P39")
    
elif ws["P40"].value == None:
    weekday = (2, "P40")
    
elif ws["P41"].value == None:
    weekday = (3, "P41")
    
elif ws["P42"].value == None:
    weekday = (4, "P42")
    
elif ws["P43"].value == None:
    weekday = (5, "P43")
    
elif ws["P44"].value == None:
    weekday = (6, "P44")
    
elif ws["P45"].value == None:
    weekday = (7, "P45")


# Title

ws.merge_cells("P35:S36")
ws["P35"].value = "Daily # of Intro Calls Scheduled"




write_table(ws, seven_weekdays, "O39:O45",
            salespeople, "P38:"+get_column_letter(len(salespeople)+17)+"38")

            
# write scraped table here


write_table(ws, data_6, weekday[1]+":"+get_column_letter(15+len(salespeople))+weekday[1][1:])   # the S in here needs to be dynamic. 



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

ws.merge_cells("C49:F50")
ws["C49"].value = "Rolling 30 Day New Inquiry Funnel"



write_table(ws, table7_rows, "B54:B64",
            table7_cols, "C53:F53")

            
# write scraped table here

write_table(ws, new_7, "E54:F64")

write_table(ws, table7_goals[0], "C54:C64")

write_table(ws, table7_goals[1], "D54:D64")

          
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
        "welcome to the celebree school franchise":0,
        "EBITDA":0,
        "from the top down": 0,
        "a day in the life": 0,
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

ws.merge_cells("C35:F36")
ws["C35"].value = "Emails Read (Weekly)"


write_table(ws, table8_cols, "B38:C38")

            
# write scraped table here
ws.merge_cells("B44:C44")
ws["B44"].value = "Out of "+ str(len(prospects))+ " leads"

count_list = list(email_count.values())



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

ws.merge_cells("B68:U69")
ws["B68"].value = "Rolling 30 Day Activity Funnel (All inquiry date)"



write_table(ws, table9_rows, "B72:B82",
            table9_cols, "C71:"+get_column_letter(len(table9_cols)+2)+"71")

            
# write scraped table here

write_table(ws, table7_goals[0], "C72:C82")

write_table(ws, table7_goals[1], "D72:D82")


i = 0
for j, table in enumerate(table_9_data):
    write_table(ws, table_9_data[j], get_column_letter(i+5)+"72:"+ get_column_letter(i+6) +"82")
    i += 2

                
                ## TABLE 9 FINISH ##
                ####################
                
                ####################
                ## TABLE 10       ##


# no scraping, just draw an empty table with the weeks' data. 

ws.merge_cells("B86:L87")

ws["B86"].value = f"Outbound Campaigns, Daily Calls, and Weekly Rolling Conversion Funnel | {current_week}"             

    

ws["B89"].value = "Relocation Campaign"
ws["B90"].value = "Disenrollment Campaign"
ws["B91"].value = "Prime Market Campaign"
ws["B95"].value = "Total Daily Outbound Calls"
ws["B96"].value = "Total Contacted Daily Calls"
ws["B99"].value = "Bad Contact #"

ws["K88"].value = "Weekly"
ws["L88"].value = "Intro Calls"
ws["J92"].value = "Total"
ws["K93"].value = "Rolling Total Intro Calls from Campaigns"
ws["K94"].value = "Total"

write_table(ws, seven_weekdays, "D88:J88",
            seven_weekdays, "D94:J94")

                





                ## TABLE 10 FINISH ##
                #####################

                
                
                
                ####################
                ## SHHHHTYLE   !  ##

                
# Base Style for ALL CELLS
    # ADD Conditional statements for the rest of the configs?
    # either this or make new methods to apply styles based on received range

for row in range(1, 100):
    ws.row_dimensions[row].height = 41
    
    for col in range(1, 60):
        char = get_column_letter(col)
    
        if row == 1:
            ws.column_dimensions[char].width = 20
        
        
        ws[char+str(row)].font = Font(size = 15)
        ws[char+str(row)].alignment = Alignment(horizontal = "center", vertical = "center", wrap_text=True)
        

ws.column_dimensions["A"].width = 8
# TITLES

ws["B68"].font = Font(size = 22, bold = True)
ws["C35"].font = Font(size = 22, bold = True)
ws["C49"].font = Font(size = 22, bold = True)
ws["P35"].font = Font(size = 22, bold = True)
ws["O19"].font = Font(size = 22, bold = True)
ws["B27"].font = Font(size = 22, bold = True)
ws["I35"].font = Font(size = 22, bold = True)
ws["N2"].font = Font(size = 22, bold = True)
ws["B2"].font = Font(size = 22, bold = True)
ws["k49"].font = Font(size = 22, bold = True)
ws["B86"].font = Font(size = 22, bold = True)



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

mod_font(ws,"O5:U5")
mod_font(ws,"N6:N12")

set_border(ws, "n5:u12")
set_border(ws, "N2:U3", _medium=True)

mod_color(ws, "N2:U3")
mod_color(ws, "O5:U5")
mod_color(ws, "N6:N12")      
 
# 3
         
mod_font(ws,"I38:L38")
mod_font(ws,"H39:H41")

set_border(ws, "H38:L41")
set_border(ws, "I35:L36", _medium=True)

mod_color(ws, "I35:L36")
mod_color(ws, "I38:L38")
mod_color(ws, "H39:H41") 

# 4

mod_font(ws,"L52:"+get_column_letter(colcount4+10)+"52") ##
mod_font(ws,"K53:K65")

set_border(ws, "K52:"+get_column_letter(colcount4+10)+str(rowcount4+50))##
set_border(ws, "K49:U50", _medium=True)

mod_color(ws, "L52:"+get_column_letter(colcount4+10)+"52")##
mod_color(ws, "K53:K65")
mod_color(ws, "K49:U50") 
                
# 5



mod_font(ws, "L22:W22")
mod_font(ws, "M29:N29")

set_border(ws, "L22:W23")
set_border(ws, "O19:T20", _medium=True)
                
set_border(ws,"M29:O29")
ws.merge_cells("M29:N29")

mod_color(ws, "O19:T20")
mod_color(ws, "L22:W22")
mod_color(ws, "M29:N29") 

# 6

#  dynamic

mod_font(ws,"P38:"+get_column_letter(len(salespeople)+15)+"38") ##
mod_font(ws,"O39:O45")

set_border(ws, "O38:"+ get_column_letter(len(salespeople)+15) +"45")  ##
set_border(ws, "P35:S36", _medium=True)

mod_color(ws, "P35:S36")
mod_color(ws, "P38:"+get_column_letter(len(salespeople)+15)+"38") ##
mod_color(ws, "O39:O45") 

# 7

mod_font(ws,"C53:F53")
mod_font(ws,"B54:B64")

set_border(ws, "B53:F64")
set_border(ws, "C49:F50", _medium=True)

mod_color(ws, "C49:F50")
mod_color(ws, "C53:F53")
mod_color(ws, "B54:B64")
       
# 8

mod_font(ws, "B38:C38")
mod_font(ws, "B39:B43")
mod_font(ws, "B44:C34")

set_border(ws, "B38:C43")
set_border(ws, "C35:F36", _medium=True)
set_border(ws, "B44:C44", _medium=True)

mod_color(ws, "C35:F36")
mod_color(ws, "B38:C38")
mod_color(ws, "B39:B43")


# 9

# dynamic

mod_font(ws,"C71:"+get_column_letter(len(table9_cols)+2)+"71") ##
mod_font(ws,"B72:B82")

set_border(ws, "B71:"+get_column_letter(len(table9_cols)+2)+"82") ## 
set_border(ws, "B68:U69", _medium=True)

mod_color(ws, "B68:U69")
mod_color(ws, "C71:"+get_column_letter(len(table9_cols)+2)+"71") ## 
mod_color(ws, "B72:B82")
       

#10

mod_font(ws, "D88:L88")
mod_font(ws, "B89:B91")
mod_font(ws, "D94:K94")
mod_font(ws, "B95:B99")
mod_font(ws, "J92:J92")
mod_font(ws, "K93:K93")

set_border(ws, "B86:L87", _medium=True)
set_border(ws, "B88:L91")
set_border(ws, "J92:L92")
set_border(ws, "K93:L93")
set_border(ws, "B94:K96")
set_border(ws, "B99:I99")

mod_color(ws, "B86:L87")
mod_color(ws, "D88:L88")
mod_color(ws, "D94:K94")
mod_color(ws, "B99:B99")
mod_color(ws, "J92:J92")
mod_color(ws, "B89:B91")
mod_color(ws, "B95:B96")
mod_color(ws, "K93:L93", "EBE600")


### CELL SPACING MODS


# LIST WITH 3 LINERS
threeliners = [76,75,58,57,54,42]



for row in threeliners:
    ws.row_dimensions[row].height = 68

ws.column_dimensions["J"].width = 18

ws.row_dimensions[93].height = 82


ws.sheet_view.zoomScale = 60


wb.save("output.xlsx")                


