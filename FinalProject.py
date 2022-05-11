
import xlsxwriter
 

workbook = xlsxwriter.Workbook('CrashReport.xlsx')

worksheet = workbook.add_worksheet()
 


bold = workbook.add_format({'bold': 1})
 

headings = ['Year', 'Atlantic', 'Bergen', 'Burlington', 'Camden', 'Cape May', 'Cumberland', 'Essex', 'Gloucester', 'Hudson', 'Hunterdon', 'Mercer', 'Middlesex', 'Monmouth', 'Morris', 'Ocean', 'Passaic', 'Salem', 'Somerset', 'Sussex', 'Union', 'Warren']

 
data = [
    [2018,2019,2020],
    [7494,29459,12238,15755,2532,3706,30078,7713,19627,4033,10473,28965,18164,14742,14848,19110,1738,10748,3153,20377,3460],
    [8140,29722,11172,14950,2478,3726,30287,7121,19729,3921,11576,28932,17507,14690,14986,17921,1723,10758,3034,21171,3317],
    [5591,19472,8888,11002,2023,3237,20071,6022,12711,2741,7495,18132,13058,8701,11970,12705,1401,6766,2409,13779,2609],
]

worksheet.write_row('A1', headings, bold)
 

worksheet.write_column('A2', data[0])
worksheet.write_column('B2', data[1])
worksheet.write_column('C2', data[2])
worksheet.write_column('D2', data[3])
 

 
# here we create a bar chart object .
chart1 = workbook.add_chart({'type': 'bar'})


chart1.add_series({
    'name':       '= Sheet1 !$B$1',
    'categories': '= Sheet1 !$A$2:$A$18',
    'values':     '= Sheet1 !$B$2:$B$18',
})
 

chart1.add_series({
    'name':       ['Sheet1', 2, 22],
    'categories': ['Sheet1', 1, 2, 0, 2],
    'values':     ['Sheet1', 1, 4, 0, 4],
})
 

chart1.set_title ({'name': 'Crash Reports in NJ Counties'})
 

chart1.set_x_axis({'name': 'Amount of Crashes'})
 

chart1.set_y_axis({'name': 'Years'})
 

chart1.set_style(11)
 

worksheet.insert_chart('E2', chart1)
 

workbook.close()