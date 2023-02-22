from repgen import *  # NOQA



### TABLE HEADERS AND DATA THAT ARE STATIC AND NOT AVAILABLE ON THE DATABASES ###

table1_cols = ["Total", "Facebook", "Franchise Gator", "LinkedIn", "Website", "PPC", "BizBuySell", "franchise.com", "IFA"]
table1_rows = ["Week Total", "Day1", "Day2", "Day3", "Day4", "Day5", "Day6", "Day7"]

table2_cols = ["Total Leads", "<1 Hour", "1-2 hours", "2-3 hours", "3-4 hours", "4-5 hours", "5+ hours"]
table2_rows = ["Day1", "Day2", "Day3", "Day4", "Day5", "Day6", "Day7"]
# rows are defined by the dates

table3_cols = ["Goal", "Week3", "Week2", "Week1"]
table3_rows = ["Leads", "Connected", "Intro Call Scheduled"]
table3_goals = ["50", "60%", "20%"]
#columns are defined by dates

#scrape table 4 entirely 

table5_cols = ['Intro Call Scheduled', 'Intro Call Completed', 'Intro Call Not Scheduled',
               'Bad Contact Info', 'Insufficient Capital', 'International', 'Market Not Available',
               "Accidentally Submitted", "Looking for Childcare", "Looking for Employment",
               "Not Interested."]

table6_rows = ["Day1", "Day2", "Day3", "Day4", "Day5", "Day6", "Day7"]
table6_cols = ["Janet", "Jackie"]


table7_cols = ["Goal Actual", "Goal %", "Actual", "Actual %"]
table7_rows = ["Leads", "Connected", "Intro Call", "Business Overview Webinar",
               "Operations and Marketing Webinar", "FDD Review", "Competency Call",
               "Executive Call", "Meet the Team Schaduled", "Decision Day",
               "Awarded"]

table7_goals = [[200,120,40,30,20,12,8,4,4,4,2],
                ["100%", "60%", "20%", "15%", "10%", "6%",
                 "4%", "2%", "2%", "2%", "1%"]]


table8_rows = ["Email", "Welcome", "EBITDA", "Support from the top down",
               "Day in the Life w/ Katie Young", "Still Interested?", "Out of X Total Leads"]
table8_cols = ["Email", "# Read"]




table9_rows = ["Leads", "Connected", "Intro Call", "Business Overview Webinar",
               "Operations and Marketing Webinar", "FDD Review", "Competency Call",
               "Executive Call", "Meet the Team Schaduled", "Decision Day",
               "Awarded"]
table9_cols = ["Goal Actual", "Goal %", "Actual", "Actual %", "Jackie Actual", "Jackie %", "Janet Actual", "Janet %"]



###   COMMENCE CODING   ###

convert_xls("P:\\dev\\Sales Reporting\\source_sheets\\1.30 daily lead count and source.xls")


wb = load_workbook("P:\\dev\\Sales Reporting\\source_sheets\\1.30 daily lead count and source.xlsx")

ws = wb.active


#This has to be dynamic, count the amount of data entries (rows) before scraping.

rowcount = count_rows(ws)

scraped = scrape_table(ws, "C27:C" + str(24 + rowcount))  #24 empty lines plus the N of populated lines.

file_list = get_file_names()

print(file_list)


if "yesterdays_report.xlsx" in file_list:
    wb = load_workbook("yesterdays_report.xlsx")
    ws = wb.active
    
   
else:
    wb = Workbook()  
    ws = wb.active
    
 

# We are able to get the data and output it correctly. 
# Now we need the logic for it to be recursively. 

# 1 - Before writing the current day table, check for data in the
# previous days, and determine what day of the week we are currently in.

# 2 - Knowing what line we are writing on, we can sum this line's values 
# with the values of the previous days, and output it in the "week total"
# row. 

# For this logic to work, we have to know if this is the first report of the week or not


weekday = None

if ws["B3"].value == None:
    weekday = (1, "B3")
    
elif ws["B4"].value == None:
    weekday = (2, "B4")
    
elif ws["B5"].value == None:
    weekday = (3, "B5")
    
elif ws["B6"].value == None:
    weekday = (4, "B6")
    
elif ws["B7"].value == None:
    weekday = (5, "B7")
    
elif ws["B8"].value == None:
    weekday = (6, "B8")
    
elif ws["B9"].value == None:
    weekday = (7, "B9")


print(weekday)
write_table(ws, table1_rows,              
            "A2:A"+str(len(table1_rows)+1),
            table1_cols,
            "B1:"+get_column_letter(len(table1_cols)+1)+"1")


write_table(ws, scraped, weekday[1] + ":" + get_column_letter(len(scraped)) + str(weekday[0]+2)) 









wb.save("TEST14.xlsx")


























