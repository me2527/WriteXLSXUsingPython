import xlsxwriter
workbook=xlsxwriter.Workbook('analysis.xlsx')
worksheet=workbook.add_worksheet()
chartsheet = workbook.add_chartsheet()
# Add a format for the headings.
bold = workbook.add_format({'bold': 1})
# Add the worksheet data that the charts will refer to.
headings = ['APPLICATIONS', 'USAGE-RANGE 1', 'USAGE-RANGE 2']
data = [
["SEARCH ENGINE","EMAIL","MAPS","VIDEOS","E-COMMERCE","SOCIAL
MEDIA","MUSIC","MESSAGING"],
[15000,14000,12000,10000,8000,6000,4000,2000],
[14000,12000,10000,8000,6000,4000,2000,0]
]
worksheet.write_row('A1', headings, bold)
worksheet.write_column('A2', data[0])
worksheet.write_column('B2', data[1])
worksheet.write_column('C2', data[2])
# Create a new bar chart.
chart1 = workbook.add_chart({'type': 'bar'})
# Configure the first series.
chart1.add_series({
'name': '=Sheet1!$B$1',
'categories': '=Sheet1!$A$2:$A$7',
'values': '=Sheet1!$B$2:$B$7',
})
# Configure a second series. Note use of alternative syntax to define ranges.
chart1.add_series({
'name': ['Sheet1', 0, 2],
'categories': ['Sheet1', 1, 0, 6, 0],
'values': ['Sheet1', 1, 2, 6, 2],
})
# Add a chart title and some axis labels.
chart1.set_title ({'name': 'Results of application usage analysis'})
chart1.set_x_axis({'name': 'USAGE-RANGE'})
chart1.set_y_axis({'name': 'APPLICATION'})
# Set an Excel chart style.
chart1.set_style(11)
# Add the chart to the chartsheet.
chartsheet.set_chart(chart1)
# Display the chartsheet as the active sheet when the workbook is opened.
chartsheet.activate();
workbook.close()