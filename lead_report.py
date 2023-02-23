from repgen import *  # NOQA

                        ###########################
                        ###   COMMENCE CODING   ###
                        ###########################



file_list = get_file_names()


                            ###################
                            #### TABLE ONE ####

table1_cols = ["Total", "Facebook", "Franchise Gator", "LinkedIn", "Website", "PPC", "BizBuySell", "franchise.com", "IFA"]
table1_rows = ["Week Total", "Day1", "Day2", "Day3", "Day4", "Day5", "Day6", "Day7"]



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

table_1_data = scrape_table(ws, "C27:C" + str(24 + rowcount))  # 24 empty lines plus the N of populated lines.


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

if ws["B3"].value == None:
    weekday = (1, "C7")
    
elif ws["B4"].value == None:
    weekday = (2, "C8")
    
elif ws["B5"].value == None:
    weekday = (3, "C9")
    
elif ws["B6"].value == None:
    weekday = (4, "C10")
    
elif ws["B7"].value == None:
    weekday = (5, "C11")
    
elif ws["B8"].value == None:
    weekday = (6, "C12")
    
elif ws["B9"].value == None:
    weekday = (7, "C13")




write_table(ws, table1_rows,              
            "B6:B"+str(len(table1_rows)+5),
            table1_cols,
            "C5:"+get_column_letter(len(table1_cols)+2)+"5")


write_table(ws, table_1_data, weekday[1] + ":" + get_column_letter(len(table_1_data)) + str(weekday[0]+6)) 



wb.save("result.xlsx")


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


# load our to-be-written xlsx file

wb = load_workbook("result.xlsx")

ws = wb.active

# Title

ws.merge_cells("P2:W3")
ws["P2"].value = "DAILY INQUIRY RESPONSE TIME"



write_table(ws, 
            table2_rows, "P6:P"+str(len(table2_rows)+5),
            table2_cols, "Q5:"+get_column_letter(len(table2_cols)+16)+"5")




# write_table(ws, table_2_data, "N2:T2")


wb.save("result.xlsx")


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
        
        table_3_data.append( scrape_table(ws, "c5:f5") )
        table_3_data[j].pop(1)
        
     
        
        j += 1
     
wb = load_workbook("result.xlsx")

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



wb.save("result.xlsx")

                ## TABLE 3 FINISH ##
                ####################
                
                
                
                ####################
                ## TABLE 4 START  ##
                
                # Title

     
wb = load_workbook("result.xlsx")

ws = wb.active

ws.merge_cells("B27:K28")
ws["B27"].value = "CURRENT PIPELINE STATUS"



wb.save("result.xlsx")


                
                ## TABLE 4 FINISH ##
                ####################
                
                
                
                
                ####################
                ## TABLE 5 START  ##
                
table5_cols = ['Total', 'Intro Call Scheduled', 'Intro Call Completed', 'Intro Call Not Scheduled',
               'Bad Contact Info', 'Insufficient Capital', 'International', 'Market Not Available',
               "Accidentally Submitted", "Looking for Childcare", "Looking for Employment",
               "Not Interested."]    


# scrape here

     
wb = load_workbook("result.xlsx")

ws = wb.active

# Title

ws.merge_cells("Q27:V28")
ws["Q27"].value = "Rolling 7 Day Inquiry Lead Status"


write_table(ws, table5_cols, "N30:"+get_column_letter(len(table5_cols)+13)+"30")

            
# write scraped table here

          
wb.save("result.xlsx")


                ## TABLE 5 FINISH ##
                ####################
                
                
                
                ####################
                ## TABLE 6 START  ##
                
table6_rows = ["Day1", "Day2", "Day3", "Day4", "Day5", "Day6", "Day7"]

table6_cols = ["Janet", "Jackie", "Sales", "Sales"]  # to be pulled from excel sheet


     
wb = load_workbook("result.xlsx")

ws = wb.active



# Title

ws.merge_cells("R46:U47")
ws["R46"].value = "Daily # of Intro Calls Scheduled"




write_table(ws, table6_rows, "Q50:Q"+str(len(table6_rows)+49),
            table6_cols, "R49:"+get_column_letter(len(table6_cols)+17)+"49")

            
# write scraped table here

          

wb.save("result.xlsx")
                
                ## TABLE 6 FINISH ##
                ####################
                
                
                
                
                ####################
                ## TABLE 7 START  ##
                

table7_cols = ["Goal Actual", "Goal %", "Actual", "Actual %"]
table7_rows = ["Leads", "Connected", "Intro Call", "Business Overview Webinar",
               "Operations and Marketing Webinar", "FDD Review", "Competency Call",
               "Executive Call", "Meet the Team Schaduled", "Decision Day",
               "Awarded"]


table7_goals = [[200,120,40,30,20,12,8,4,4,4,2],
                ["100%", "60%", "20%", "15%", "10%", "6%",
                 "4%", "2%", "2%", "2%", "1%"]]


     
wb = load_workbook("result.xlsx")

ws = wb.active


# Title

ws.merge_cells("E62:H63")
ws["E62"].value = "Rolling 30 Day New Inquiry Funnel"



write_table(ws, table7_rows, "D66:D"+str(len(table7_rows)+65),
            table7_cols, "E65:"+get_column_letter(len(table7_rows)+4)+"65")

            
# write scraped table here

          
wb.save("result.xlsx")

                
                ## TABLE 7 FINISH ##
                ####################
                
                ####################
                ## TABLE 8 START  ##

table8_rows = ["Email", "Welcome", "EBITDA", "Support from the top down",
               "Day in the Life w/ Katie Young", "Still Interested?", "Out of X Total Leads"]
table8_cols = ["Email", "# Read"]


     
wb = load_workbook("result.xlsx")

ws = wb.active


# Title

ws.merge_cells("R62:U63")
ws["R62"].value = "Emails Read from 1/22  -  1/28 (weekly)"


write_table(ws, table8_rows, "P67:P"+str(len(table8_rows)+66),
            table8_cols, "Q66:"+get_column_letter(len(table8_cols)+16)+"66")

            
# write scraped table here

          
wb.save("result.xlsx")

                
                ## TABLE 8 FINISH ##
                ####################
                
                ####################
                ## TABLE 9 START  ##
                

table9_rows = ["Leads", "Connected", "Intro Call", "Business Overview Webinar",
               "Operations and Marketing Webinar", "FDD Review", "Competency Call",
               "Executive Call", "Meet the Team Schaduled", "Decision Day",
               "Awarded"]
table9_cols = ["Goal Actual", "Goal %", "Actual", "Actual %", "Jackie Actual", "Jackie %", "Janet Actual", "Janet %"]


     
wb = load_workbook("result.xlsx")

ws = wb.active


# Title

ws.merge_cells("F79:T80")
ws["F79"].value = "Rolling 30 Day Activity Funnel (All inquiry date)"






write_table(ws, table9_rows, "F83:F"+str(len(table9_rows)+82),
            table9_cols, "G82:"+get_column_letter(len(table9_cols)+6)+"82")

            
# write scraped table here


                
                
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
        
        
        ws[char+str(row)].font = Font(size = 16)
        ws[char+str(row)].alignment = Alignment(horizontal = "center", vertical = "center", wrap_text=True)
        
# TITLES

ws["F79"].font = Font(size = 22, bold = True)
ws["R62"].font = Font(size = 22, bold = True)
ws["E62"].font = Font(size = 22, bold = True)
ws["R46"].font = Font(size = 22, bold = True)
ws["Q27"].font = Font(size = 22, bold = True)
ws["B27"].font = Font(size = 22, bold = True)
ws["R16"].font = Font(size = 22, bold = True)
ws["P2"].font = Font(size = 22, bold = True)
ws["B2"].font = Font(size = 22, bold = True)


# Table Header Cols and Rows

# mod_font(ws, )

              
# set borders

# set_border()                
                

wb.save("result2.xlsx")                
                
  
                
                
                