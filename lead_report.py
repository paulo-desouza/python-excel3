from repgen import *  # NOQA


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














convert_xls("P:\\dev\\Sales Reporting\\source_sheets\\1.30 daily lead count and source.xls")


wb = load_workbook("P:\\dev\\Sales Reporting\\source_sheets\\1.30 daily lead count and source.xlsx")

ws = wb.active

scraped_shit = scrape_table(ws, "B28:D31")

print(scraped_shit)









wb = Workbook()  # NOQA
ws = wb.active
 




write_table(ws, table1_rows, table1_cols,              
            "A2:A"+str(len(table1_rows)+1),
            "B1:"+get_column_letter(len(table1_cols)+1)+"1")
















































wb.save("TEST7.xlsx")









