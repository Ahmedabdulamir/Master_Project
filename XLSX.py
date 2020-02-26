import pandas as pd
import os
import xlsxwriter
import shutil



def _chkfile(Path, ParserIn):
    Direction = ParserIn.get('Direction')
    if not os.path.exists('xlsx/' + Path):
        os.mkdir('xlsx/' + Path)

    if os.path.isfile('xlsx/' + Path + '/'+ Path + Direction + '.xlsx'):
        overwirte = input("File exists, overwirte?(Y,N) (Default = overwirte)")
        if overwirte == 'N':
            print('Please rename and close the file and retry again')
        else:
            while True:
                try:
                    os.remove('xlsx/' + Path + '/'+ Path + Direction + '.xlsx')
                except IOError:
                    input('Please close Excelfile to continue')
                    continue
                break

def _no2lt(col):
    col = xlsxwriter.utility.xl_col_to_name(col)
    return (col)

def _color(workbook, No, bold ):
    No_dict = { '01' :'cell_format01','02' : 'cell_format02', '03' : 'cell_format03'}
    format_dict = {'01': 'E0F2F6' ,'02' : 'yellow' ,'03' : 'E1E1E1'}
    No_dict[No] = workbook.add_format()
    No_dict[No].set_pattern(1)
    No_dict[No].set_bg_color(format_dict[No])
    if bold == True: No_dict[No].set_bold()
    return No_dict[No]


def _fill_data(ParserIn, Para_dict, workbook, worksheet, format_dict):
    col = 0
    col1 = 1
    for Para in ParserIn.get('Parameter'):
        sec_dict = Para_dict.get(Para)
        for Sec in sec_dict.keys():
            w_dict = sec_dict.get(Sec)
            worksheet.set_column(_no2lt(col1+1) + ':' + _no2lt(col1+9), None, None, {'hidden': True})
            worksheet.write(0, col1, Para + '_avg',format_dict['head'])
            worksheet.write(7, col1, w_dict[ParserIn['W'][0]][Para + '_h'],format_dict[Para])
            row = 1
            for W in w_dict.keys():
                data_dict = w_dict.get(W)
                worksheet.write_column(1, 0, ('W=' + str(x) for x in ParserIn['W']),format_dict['head'])
                worksheet.write_column(10, 0, ('W=' + str(x) for x in ParserIn['W']), format_dict['head'])
                worksheet.write(26, col, 'Y_w=' + str(W),format_dict['head'])
                worksheet.write_column(27, col, data_dict['Y'],format_dict['head'])
                worksheet.write(26, col+1, Para +'_X=' + str(Sec),format_dict['head'])

                worksheet.write_column(27, col+1, data_dict[Para],format_dict[Para])
                worksheet.write(row, col1, data_dict[Para + '_avg'],format_dict[Para])
                worksheet.write(row+9, col1, data_dict[Para + '_h']/data_dict[Para + '_avg'],format_dict[Para])
                worksheet.write(row+14, col1, data_dict[Para + '_h']/2/data_dict[Para + '_avg'],format_dict[Para])
                worksheet.write(row+19, col1, data_dict[Para + '_h']/(-2)/data_dict[Para + '_avg'],format_dict[Para])
                row+=1
                col+=2
            worksheet.write(9, col1, Sec, format_dict['head'])
            col1 = col+1
    return(col, col1)


def _chart(chart, chart_dict, ParserIn):
    chart.show_hidden_data()
    chart.set_style(chart_dict['style'])
    axis = {       'name': chart_dict['x_label'],
             'minor_unit': chart_dict['minor_unit'],
             'major_unit': chart_dict['major_unit'],
        'major_gridlines': {    'visible': True,
                                   'line': {'width': 1.0, 'dash_type': 'solid'}},
        'minor_gridlines': {    'visible': True,
                                   'line': {'width': 1.0, 'dash_type': 'solid'}},
        'major_unit_type': chart_dict['x_unit']}

    chart.set_x_axis(axis)
    axis['name'] = chart_dict['y_label']
    axis['major_unit_type'] = chart_dict['y_unit']
    chart.set_y_axis(axis)
    
    chart.set_size({'x_scale': chart_dict['size'] , 'y_scale': ParserIn['aspect_ratio']*chart_dict['size']})
    chart.set_title({'name': chart_dict['Tiltle']})
    return(chart)


def xlsExcel(Path, Para_dict, ParserIn):
    Direction = ParserIn.get('Direction')
    _chkfile(Path, ParserIn)

    workbook = xlsxwriter.Workbook('xlsx/' + Path + '/'+ Path + Direction + '.xlsx', {'strings_to_numbers': True})
    worksheet = workbook.add_worksheet()
    
    format_dict = {'M2' : _color(workbook, '01', False)
                  ,'Vy' : _color(workbook, '02', False)
                 ,'head' : _color(workbook, '03', True)}
    
    (col, col1) = _fill_data(ParserIn, Para_dict, workbook, worksheet, format_dict)

    Tamplate_dict = {'A1' : 'w',
                     'A7' : 'Hand Calculation',
                     'A8' : 'Moment/Shear',
                    'A10' : 'Eq.width (m)' }
    
    for key in Tamplate_dict.keys():
        worksheet.write(key, Tamplate_dict[key], format_dict['head'])
    
    for row in [1 , 10] + list(range(15,25)):
        worksheet.set_row(row, None, None, {'hidden': True})

    _fill_data(ParserIn, Para_dict, workbook, worksheet, format_dict)
    
    #chart5 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
    chart_dict = {'Tiltle' : 'Equavalant Width',
                    'size' : 1,
                   'style' : 14,
                  'x_unit' : 'm',
                 'x_label' : 'Width (m)',
                  'y_unit' : 'm',
                 'y_label' : 'Length (m)',
               'minor_unit': ParserIn['Mesh'],
               'major_unit': 1 }

    chart2 = _chart(workbook.add_chart({'type': 'scatter'}), chart_dict, ParserIn)
    
    for row in range(17,21):
        chart2.add_series({'categories': '=(Sheet1!$B$10:$'+  _no2lt(col1) +'$10,' + 'Sheet1!$B$10:$'+  _no2lt(col1) +'$10)',
                               'values': '=(Sheet1!$B$' + str(row)  + ':$'+  _no2lt(col1) +'$'+ str(row) + ','
                                        + 'Sheet1!$B$' + str(row+5)  + ':$'+  _no2lt(col1) +'$'+ str(row+5) + ')'  ,
                                 'name': '=Sheet1!$A$'+ str(row - 5)})
    
    worksheet.insert_chart( _no2lt(col1+1) + '2', chart2)

    workbook.close()
    # Add the worksheet data to be plotted.
    
    # Create a new chart object.
   

    # Add a series to the chart.
   

    # Insert the chart into the worksheet.
    
