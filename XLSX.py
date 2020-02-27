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


def _fill_data(ParserIn, Para_dict, worksheet, format_dict, Para, col):
    col1 = col + 1
    w_col = len(ParserIn['W'])
    sec_dict = Para_dict.get(Para)
    for Sec in sec_dict.keys():
        w_dict = sec_dict.get(Sec)
        worksheet.set_column(_no2lt(col1+1) + ':' + _no2lt(col1+w_col*2-1), None, None, {'hidden': True})
        worksheet.write(0, col1, Para + '_avg',format_dict['head'])
        worksheet.write(w_col+2, col1, w_dict[ParserIn['W'][0]][Para + '_h'],format_dict[Para])
        row = 1
        for W in w_dict.keys():
            data_dict = w_dict.get(W)
            worksheet.write_column(1, 0, ('W=' + str(x) for x in ParserIn['W']),format_dict['head'])
            worksheet.write_column(w_col + 5, 0, ('W=' + str(x) for x in ParserIn['W']), format_dict['head'])
            worksheet.write(w_col*4 + 6, col, 'X_w=' + str(W),format_dict['head'])
            worksheet.write_column(w_col*4 + 7, col, data_dict['X'],format_dict['head'])
            worksheet.write(w_col*4 + 6, col+1, Para +'_X=' + str(Sec),format_dict['head'])

            worksheet.write_column(w_col*4 + 7, col+1, data_dict[Para],format_dict[Para])
            worksheet.write(row, col1, data_dict[Para + '_avg'],format_dict[Para])
            worksheet.write(row+w_col+4, col1, data_dict[Para + '_h']/data_dict[Para + '_avg'],format_dict[Para])
            worksheet.write(row+w_col*2+4, col1, data_dict[Para + '_h']/2/data_dict[Para + '_avg'],format_dict[Para])
            worksheet.write(row+w_col*3+4, col1, data_dict[Para + '_h']/(-2)/data_dict[Para + '_avg'],format_dict[Para])
            row+=1
            col+=2
        worksheet.write(w_col+4, col1, Sec, format_dict['head'])
        col1 = col+1
        worksheet.write('A'+ str(w_col*2+8), '=MAX(B'+ str(w_col*2+7) +':'+_no2lt(col) + str(w_col*3+5) + ')')
        worksheet.write('A'+ str(w_col*2+9), '=-A'+ str(w_col*2+8))
        
    return(col1)


def _chartformat(chart, chart_dict, ParserIn):
    chart.show_hidden_data()
    #chart.set_style(chart_dict['style'])
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


def _Graph(chart, Path, FEMIn, ParserIn):
    w_col = len(ParserIn['W'])
    chart.add_series({'categories': '{' + str(FEMIn['yload']) +'}',
                           'values': '{0}',
                             'name': 'Load',
                           'marker': {
                             'type': 'square',
                             'size': 20,
                           'border': {'color': 'black'},
                             'fill':   {'color': 'red'}}})


    chart.add_series({'categories': '{' + str(FEMIn['ysupp'][0])  + ',' + str(FEMIn['ysupp'][0]) +'}',
                           'values': '=' + Path + '!A' + str(w_col*2+8)+ ':' + 'A'+ str(w_col*2+9),
                             'name': 'Support 1',
                           'marker': {
                             'type': 'square',
                             'size': 2,
                           'border': {'color': 'red'},
                             'fill':   {'color': 'red'}} ,
                             'line': {'color': 'red','width': 2}})


    chart.add_series({'categories': '{' + str(FEMIn['ysupp'][1])  + ',' + str(FEMIn['ysupp'][1]) +'}',
                           'values': '=' + Path + '!A' + str(w_col*2+8)+ ':' + 'A'+ str(w_col*2+9),
                             'name': 'Support 2',
                           'marker': {
                             'type': 'square',
                             'size': 2,
                           'border': {'color': 'red'},
                             'fill':   {'color': 'red'}} ,
                             'line': {'color': 'red', 'width': 2}}) 

                             #chart1.add_series({'categories':'{' + str(FEMIn['yload'] + FEMIn['dia']) + ',' + str(FEMIn['yload'] + FEMIn['dia']) + ',' +  
   #                                       str(FEMIn['yload'] - FEMIn['dia']) + ',' + str(FEMIn['yload'] - FEMIn['dia']) +'}',
    #                       'values':'{' + str(- FEMIn['dia']) + ',' + str(  FEMIn['dia']) + ',' + 
    #                                      str(  FEMIn['dia']) + ',' + str(- FEMIn['dia']) +'}',
     #                        'name': 'Load',

    #                         'line': {'color': 'red'}})
    return

def _eqchart(ParserIn, Path, Para, workbook, col1, col2):
    #chart5 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
    w_col = len(ParserIn['W'])
    chart_dict = {'Tiltle' : 'Equivalent Width for '+ Para,
                    'size' : 1.1,
                   'style' : 14,
                  'x_unit' : 'm',
                 'x_label' : 'Width (m)',
                  'y_unit' : 'm',
                 'y_label' : 'Length (m)',
               'minor_unit': ParserIn['Mesh'],
               'major_unit': 1 }

    chart = _chartformat(workbook.add_chart({'type': 'scatter'}), chart_dict, ParserIn)
    
    scol1 = _no2lt(col1)
    scol2 = _no2lt(col2-1)
    for row in range(w_col*2+7,w_col*3+6):
        chart.add_series({'categories': '=(' + Path + '!$' + scol1 + '$'+ str(w_col+5) +':$'+  scol2 +'$'+ str(w_col+5) +',' 
                                        + '' + Path + '!$' + scol1 + '$'+ str(w_col+5) +':$'+  scol2 +'$'+ str(w_col+5) +')',
                               'values': '=(' + Path + '!$' + scol1 + '$' + str(row)  + ':$'+  scol2 +'$'+ str(row) + ','
                                        + '' + Path + '!$' + scol1 + '$' + str(row+w_col)  + ':$'+  scol2 +'$'+ str(row+w_col) + ')'  ,
                                 'name': '=' + Path + '!$A$'+ str(row - w_col),
                                 'marker': {'type': 'automatic'},})
    return(chart)

def _secchart(ParserIn, Path, workbook, col1, col2):
    #chart5 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
    w_col = len(ParserIn['W'])
    chart_dict = {'Tiltle' : 'Section',
                    'size' : 1.1,
                   'style' : 14,
                  'x_unit' : 'm',
                 'x_label' : 'Width (m)',
                  'y_unit' : 'm',
                 'y_label' : 'Length (m)',
               'minor_unit': ParserIn['Mesh'],
               'major_unit': 1 }

    chart = _chartformat(workbook.add_chart({'type': 'scatter'}), chart_dict, ParserIn)
    
    scol1 = _no2lt(col1)
    scol2 = _no2lt(col2-1)
    for row in range(w_col*2+7,w_col*3+6):
        chart.add_series({'categories': '=(' + Path + '!$' + scol1 + '$'+ str(w_col+5) +':$'+  scol2 +'$'+ str(w_col+5) +',' 
                                        + '' + Path + '!$' + scol1 + '$'+ str(w_col+5) +':$'+  scol2 +'$'+ str(w_col+5) +')',
                               'values': '=(' + Path + '!$' + scol1 + '$' + str(row)  + ':$'+  scol2 +'$'+ str(row) + ','
                                        + '' + Path + '!$' + scol1 + '$' + str(row+w_col)  + ':$'+  scol2 +'$'+ str(row+w_col) + ')'  ,
                                 'name': '=' + Path + '!$A$'+ str(row - w_col),
                                 'marker': {'type': 'automatic'},})
    return(chart)


def _fill_formula(ParserIn, Para_dict, worksheet, col, refcell):    
    w_col = len(ParserIn['W'])    
    len_itr = len(Para_dict['M2'][ParserIn['Section_M2'][0]][max(ParserIn['W'])]['X'])      #fixed to M2
    
    worksheet.write_column( w_col*4 + 7 , col + 2, list(range(1,len_itr)))

    for j in range (w_col*4 + 7, len_itr + w_col*4 + 7):
        for i in range(-1, w_col*2-2):
            row = ('=OFFSET($A$'+ str(w_col*4 + 7) +',$'+ _no2lt(col+2) 
                +str(j+1) +',MATCH($'+ refcell + ',$'+ str(w_col*4 + 7) +
                ':$'+ str(w_col*4 + 7) + ',0)+' + str(i) +')')
            worksheet.write( j , col + 4 + i, row) 
    worksheet.set_column(_no2lt(col+2) + ':' + _no2lt(col + w_col*2 +2), None, None, {'hidden': True})
    return(col + w_col*2 +2)     


def xlsExcel(Path, Para_dict, ParserIn, FEMIn):
    w_col = len(ParserIn['W'])
    Direction = ParserIn.get('Direction')
    _chkfile(Path, ParserIn)

    workbook = xlsxwriter.Workbook('xlsx/' + Path + '/'+ Path + Direction + '.xlsx', {'strings_to_numbers': True})
    worksheet = workbook.add_worksheet(Path)
    
    format_dict = {'M2' : _color(workbook, '01', False)
                  ,'Vy' : _color(workbook, '02', False)
                 ,'head' : _color(workbook, '03', True)}
          
    Para1 = ParserIn['Parameter'][0]
    Para2 = ParserIn['Parameter'][1]
    
    col1 = 1
    col2 = _fill_data(ParserIn, Para_dict,  worksheet, format_dict, Para1, 0)
    chart1 = _eqchart(ParserIn, Path, Para2, workbook, col1, col2)
    _Graph(chart1, Path, FEMIn, ParserIn)
        
    col1 = col2
    col2 = _fill_data(ParserIn, Para_dict,  worksheet, format_dict, Para2, col2-1)
    chart2 = _eqchart(ParserIn, Path, Para2, workbook, col1, col2)
    _Graph(chart2, Path, FEMIn, ParserIn)
    
    Tamplate_dict = {'A1' : 'w',
                     'A'+ str(w_col+2) : 'Hand Calculation',
                     'A'+ str(w_col+3) : 'Moment/Shear',
                     'A'+ str(w_col+5) : 'Eq.width (m)',
                   'A'+ str(w_col*2+7) : 'Eq.width (m)',
        _no2lt(col2+1)+ str(w_col*4+6) : 'Plot',
        _no2lt(col2+1)+ str(w_col*4+7) : Para1 +'_X=' + str(ParserIn['Section_' + Para1][0]),
        _no2lt(col2+1)+ str(w_col*4+8) : Para2 +'_X=' + str(ParserIn['Section_' + Para2][0])}
    
    for key in Tamplate_dict.keys():
        worksheet.write(key, Tamplate_dict[key], format_dict['head'])

    for row in [1 , w_col + 5] + list(range((w_col*2 + 5), (w_col*4 + 5))):
        worksheet.set_row(row, None, None, {'hidden': True})
    
    refcell = { '1' : _no2lt(col2+1)+ str(w_col*4+7) ,
                '2' : _no2lt(col2+1)+ str(w_col*4+8) }

    col1 = col2
    col2 = _fill_formula(ParserIn, Para_dict, worksheet, col2, refcell['1'])
    col1 = col2
    col2 = _fill_formula(ParserIn, Para_dict, worksheet, col2, refcell['2'])

    worksheet.insert_chart( _no2lt(col2 ) + '2', chart1)
    worksheet.insert_chart( _no2lt(col2 +9) + '2', chart2)

    
    workbook.close()

