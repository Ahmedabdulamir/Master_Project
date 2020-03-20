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


def _fill_data(ParserIn, FEMIn, Para_dict, worksheet, format_dict, Para, TableForm, col):
    q1q2 = 1
    col1 = col + 1
    w_col = len(ParserIn['W'])
    sec_dict = Para_dict.get(Para)
    for Sec in sec_dict.keys():
        w_dict = sec_dict.get(Sec)
        worksheet.set_column(_no2lt(col1+1) + ':' + _no2lt(col1+w_col*2-1), None, None, {'hidden': True})       #Hiding col
        worksheet.write(TableForm['Para_avg'], col1, Para + '_avg',format_dict['head'])     #Writting header for avg
        worksheet.write(TableForm['Hand_calc'], col1, Para + '_h',format_dict['head'])     #Writting header for h
        if Sec in FEMIn['ysupp']:
            worksheet.write(TableForm['q1q2'], col1, 'q' + str(q1q2) ,format_dict['head'])     #Writting header for q1q2
            q1q2 + 1
        row = 1
        for W in w_dict.keys():
            data_dict = w_dict.get(W)
            worksheet.write_column(TableForm['Para_avg']+1, 0, ('W=' + str(x) for x in ParserIn['W']),format_dict['head'])    #header for w col for avg
            worksheet.write_column(TableForm['Eq_header']+1, 0, ('W=' + str(x) for x in ParserIn['W']), format_dict['head'])   #header for w col for beq   
            worksheet.write_column(TableForm['Hand_calc']+1, 0, ('W=' + str(x) for x in ParserIn['W']), format_dict['head'])
            worksheet.write_column(TableForm['q1q2']+1, 0, ('W=' + str(x) for x in ParserIn['W']), format_dict['head'])

            worksheet.write(TableForm['Main Header'], col, 'X_w=' + str(W),format_dict['head'])          #header for x col
            worksheet.write_column(TableForm['Main Header']+1, col, data_dict['X'],format_dict['head'])        #X values col
            
            worksheet.write(TableForm['Main Header'], col+1, Para +'_X=' + str(Sec),format_dict['head'])     #header Para col
            worksheet.write_column(TableForm['Main Header']+1, col+1, data_dict[Para],format_dict[Para])       #Para values col
                        
            worksheet.write(row+TableForm['Para_avg'], col1, data_dict[Para + '_avg'],format_dict[Para])              #average Para value 

            worksheet.write(row+TableForm['Hand_calc'], col1, data_dict[Para + '_h'],format_dict[Para])         #Writting hand calc
            if Sec in FEMIn['ysupp']:
                worksheet.write(row+TableForm['q1q2'], col1, data_dict['q' + str(Sec)],format_dict[Para])
                
            if data_dict[Para + '_avg'] != 0:
                worksheet.write(row+TableForm['Eq_header'], col1, data_dict[Para + '_h']/data_dict[Para + '_avg'],format_dict[Para])       # Beq values
                worksheet.write(row+TableForm['-ve_header'], col1, data_dict[Para + '_h']/2/data_dict[Para + '_avg'],format_dict[Para])       # 1/2 Beq values
                worksheet.write(row+TableForm['+ve_header'], col1, data_dict[Para + '_h']/(-2)/data_dict[Para + '_avg'],format_dict[Para])        # -1/2 Beq values
            row+=1
            col+=2
        worksheet.write(TableForm['Eq_header'], col1, Sec, format_dict['head'])    #Section header
        worksheet.write(TableForm['y_head'], col1, Sec, format_dict['head'])
        col1 = col+1
        worksheet.write('A'+ str(TableForm['-ve_header']+4), '=MAX(B'+ str(TableForm['-ve_header']+3) +':'+_no2lt(col) + str(TableForm['+ve_header']+1) + ')')  #Max value of Beq for plot
        worksheet.write('A'+ str(TableForm['-ve_header']+5), '=-A'+ str(TableForm['-ve_header']+4))     #- Max value of Beq for plot
        
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
    
    chart.set_size({'x_scale': chart_dict['size'] , 'y_scale': chart_dict['aspect_ratio']*chart_dict['size']})
    chart.set_title({'name': chart_dict['Tiltle']})
    return(chart)


def _Graph(chart, Path, FEMIn, ParserIn, TableForm):
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
                           'values': '=' + Path + '!A' + str(TableForm['-ve_header']+4) + ':' + 'A'+ str(TableForm['-ve_header']+5),
                             'name': 'Support 1',
                           'marker': {
                             'type': 'square',
                             'size': 2,
                           'border': {'color': 'red'},
                             'fill':   {'color': 'red'}} ,
                             'line': {'color': 'red','width': 2}})


    chart.add_series({'categories': '{' + str(FEMIn['ysupp'][1])  + ',' + str(FEMIn['ysupp'][1]) +'}',
                           'values': '=' + Path + '!A' + str(TableForm['-ve_header']+4) + ':' + 'A'+ str(TableForm['-ve_header']+5),
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

def _eqchart(ParserIn, Path, Para, workbook, col1, col2, TableForm):
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
               'major_unit': 1 ,
             'aspect_ratio': ParserIn['aspect_ratio']}

    chart = _chartformat(workbook.add_chart({'type': 'scatter'}), chart_dict, ParserIn)
    
    scol1 = _no2lt(col1)
    scol2 = _no2lt(col2-1)
    for row in range(TableForm['-ve_header']+3,TableForm['+ve_header']+2):
        chart.add_series({'categories': '=(' + Path + '!$' + scol1 + '$'+ str(TableForm['Eq_header']+1) +':$'+  scol2 +'$'+ str(TableForm['Eq_header']+1) +',' 
                                        + '' + Path + '!$' + scol1 + '$'+ str(TableForm['Eq_header']+1) +':$'+  scol2 +'$'+ str(TableForm['Eq_header']+1) +')',
                               'values': '=(' + Path + '!$' + scol1 + '$' + str(row)  + ':$'+  scol2 +'$'+ str(row) + ','
                                        + '' + Path + '!$' + scol1 + '$' + str(row+w_col)  + ':$'+  scol2 +'$'+ str(row+w_col) + ')'  ,
                                 'name': '=' + Path + '!$A$'+ str(row - w_col),
                                 'marker': {'type': 'automatic'},})
    return(chart)

def _secchart(ParserIn, Path, workbook, col1, col2, dlen, refcell):
    #chart5 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
    chart_dict = {'Tiltle' : '=' + Path + '!$' + refcell['Title'],
                    'size' : 1.2,
                   'style' : 14,
                  'x_unit' : 'm',
                 'x_label' : 'Width (m)',
                  'y_unit' : 'm',
                 'y_label' : 'Length (m)',
               'minor_unit': ParserIn['Mesh'],
               'major_unit': 1,
             'aspect_ratio': 1.7 }

    major_gridlines = {'visible': True,
                       'line': {'width': 1.0, 'dash_type': 'solid'}}
    minor_gridlines = {'visible': True,
                       'line': {'width': 1.0, 'dash_type': 'solid'}}
    
    refcell['x']['major_gridlines'] = major_gridlines
    refcell['x']['minor_gridlines'] = minor_gridlines
    refcell['y']['major_gridlines'] = major_gridlines
    refcell['y']['minor_gridlines'] = minor_gridlines                    

    chart = _chartformat(workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'}), chart_dict, ParserIn)
    chart.set_x_axis(refcell['x'])
    chart.set_y_axis(refcell['y'])
                    
    i=0
    for col in range(col1+3 ,col2+3, 2):
        chart.add_series({'categories': '=(' + Path + '!$' + _no2lt(col) + '$1:$'+  _no2lt(col) +'$'+ str(dlen) + ')',
                              'values': '=(' + Path + '!$' + _no2lt(col+1) + '$1:$'+  _no2lt(col+1) +'$'+ str(dlen) + ')',
                                'name': 'W =' + str(ParserIn['W'][i]),
                              'marker': {
                                'type': 'automatic',
                                'size': 2,} ,
                                'line': {'width': 1.5}})
        i+=1
    return(chart)

def _multisecchart(ParserIn, Para_dict, Path, workbook, col1, col2, Para, refcell, TableForm):
    #chart5 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
    len_itr = len(Para_dict['My'][ParserIn['Section_My'][0]][max(ParserIn['W'])]['X']) 
    w_col = len(ParserIn['W'])
    
    if Para == 'My':
        col2 = col1
        col1 = 0

    chart_dict = {'Tiltle' : refcell['Title'],
                    'size' : 1.2,
                   'style' : 14,
                  'x_unit' : 'm',
                 'x_label' : 'Width (m)',
                  'y_unit' : 'm',
                 'y_label' : 'Length (m)',
               'minor_unit': ParserIn['Mesh'],
               'major_unit': 1,
             'aspect_ratio': 1.7 }

    major_gridlines = {'visible': True,
                       'line': {'width': 1.0, 'dash_type': 'solid'}}
    minor_gridlines = {'visible': True,
                       'line': {'width': 1.0, 'dash_type': 'solid'}}
    
    refcell['x']['major_gridlines'] = major_gridlines
    refcell['x']['minor_gridlines'] = minor_gridlines
    refcell['y']['major_gridlines'] = major_gridlines
    refcell['y']['minor_gridlines'] = minor_gridlines                    

    chart = _chartformat(workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'}), chart_dict, ParserIn)
    chart.set_x_axis(refcell['x'])
    chart.set_y_axis(refcell['y'])
                                                            
    i=0
    for col in range(col1 ,col2, w_col*2):
        chart.add_series({'categories': '=(' + Path + '!$' + _no2lt(col) + '$'+ str(TableForm['Main Header']+2) +':$'+  _no2lt(col) +'$'+ str(TableForm['Main Header']+1 + len_itr) + ')',
                              'values': '=(' + Path + '!$' + _no2lt(col+1) + '$'+ str(TableForm['Main Header']+2) +':$'+  _no2lt(col+1) +'$'+ str(TableForm['Main Header']+1 + len_itr) + ')',
                                'name': '=' + Path + '!$' + _no2lt(col+1) + '$'+ str(TableForm['Main Header']+1),
                              'marker': {
                                'type': 'automatic',
                                'size': 2,} ,
                                'line': {'width': 1.5}})
        i+=1
    return(chart)

def _fill_formula(ParserIn, Para_dict, worksheet, col, refcell, overwrite, TableForm):    
    w_col = len(ParserIn['W'])    
    len_itr = len(Para_dict['My'][ParserIn['Section_My'][0]][max(ParserIn['W'])]['X'])      #fixed to My
    
    worksheet.write_column( TableForm['Main Header']+1 , col + 2, list(range(1,len_itr+1)))

    for j in range (TableForm['Main Header']+1, len_itr + TableForm['Main Header']+1):
        for i in range(-2, w_col*2-2):
            row = ('OFFSET($A$'+ str(TableForm['Main Header']+1) +',$'+ _no2lt(col+2) 
                +str(j+1) +',MATCH($'+ refcell + ',$'+ str(TableForm['Main Header']+1) +
                ':$'+ str(TableForm['Main Header']+1) + ',0)+' + str(i) +')')                
            worksheet.write( j , col + 5 + i, '=IF(' + row + '=0,' + _no2lt(col + 5 + i) + str(j) + ',' + row + ')')
        k = 0
        if overwrite == True:
            for i in range(-1, w_col*2-3,2):
                row = ('=OFFSET($A$'+ str(TableForm['Para_avg'] + 3) +',' + str(k) + ',MATCH($'+ refcell + ',$'
                                    + str(TableForm['Main Header']+1) + ':$'+ str(TableForm['Main Header']+1) + ',0)-1)')
                worksheet.write( j , col + 7 + i, row)
                k+=1
    worksheet.set_column(_no2lt(col+2) + ':' + _no2lt(col + w_col*2 +3), None, None, {'hidden': True})
    return(col + w_col*2, w_col*4 + 8 + len_itr)     


def xlsExcel(Path, Para_dict, ParserIn, FEMIn, TableForm):
    w_col = len(ParserIn['W'])
    Direction = ParserIn.get('Direction')
    #_chkfile(Path, ParserIn)

    workbook = xlsxwriter.Workbook('xlsx/' + Path + '/'+ Path + Direction + '.xlsx', {'strings_to_numbers': True})
    worksheet = workbook.add_worksheet(Path)
    
    format_dict = {'M2' : _color(workbook, '01', False)
                  ,'Vy' : _color(workbook, '02', False)
                ,'head' : _color(workbook, '03', True)
                  ,'My' : _color(workbook, '01', False)
                 ,'Mxy' : _color(workbook, '02', False)
                                                        }
          
    Para1 = ParserIn['Parameter'][0]
    Para2 = ParserIn['Parameter'][1]
    
    col1 = 1
    col2 = _fill_data(ParserIn, FEMIn, Para_dict,  worksheet, format_dict, Para1, TableForm, 0)
    chart1 = _eqchart(ParserIn, Path, Para1, workbook, col1, col2, TableForm)
    _Graph(chart1, Path, FEMIn, ParserIn, TableForm)
        
    col1 = col2
    col2 = _fill_data(ParserIn, FEMIn, Para_dict,  worksheet, format_dict, Para2, TableForm, col2-1)
    chart2 = _eqchart(ParserIn, Path, Para2, workbook, col1, col2, TableForm)
    _Graph(chart2, Path, FEMIn, ParserIn, TableForm)
    

    refcell = { '1' : {'Title' : _no2lt(col2+1)+ '$' + str(w_col*4+42),
                            'x':  {'name' : 'Width (m)',
                                   'min': min(Para_dict['My'][ParserIn['Section_My'][0]][ParserIn['W'][w_col-1]]['X'])*0.75,
                                   'max': max(Para_dict['My'][ParserIn['Section_My'][0]][ParserIn['W'][w_col-1]]['X'])*1.25},
                            'y':  {'name' : Para1 + ' (kN.m/m)',
                                   'minor_unit': '0.01', #min(Para_dict['M2'][ParserIn['Section_M2'][0]][max(ParserIn['W'])]['M2']),
                                   'major_unit': '0.05'}}, #max(Para_dict['M2'][ParserIn['Section_M2'][0]][max(ParserIn['W'])]['M2'])}}, 
                '2' : {'Title' : _no2lt(col2+1)+ '$' + str(w_col*4+43),
                            'x':  {'name' : 'Width (m)',
                                   'min': min(Para_dict['Vy'][ParserIn['Section_Vy'][0]][ParserIn['W'][w_col-1]]['X'])*0.75,
                                   'max': max(Para_dict['Vy'][ParserIn['Section_Vy'][0]][ParserIn['W'][w_col-1]]['X'])*1.25},
                            'y':  {'name' : Para2 + ' (kN/m)',
                                   'minor_unit': '0.01', #min(Para_dict['M2'][ParserIn['Section_M2'][0]][max(ParserIn['W'])]['M2']),
                                   'major_unit': '0.05'}},
                '3' : {'Title' : Para1 + ' all sections',
                            'x':  {'name' : 'Width (m)',
                                   'min': max(Para_dict['My'][ParserIn['Section_My'][0]][max(ParserIn['W'])]['X'])/3,
                                   'max': max(Para_dict['My'][ParserIn['Section_My'][0]][max(ParserIn['W'])]['X'])*2/3},
                            'y':  {'name' : Para1 + ' (kN.m/m)',
                                   'minor_unit': '0.01', #min(Para_dict['M2'][ParserIn['Section_M2'][0]][max(ParserIn['W'])]['M2']),
                                   'major_unit': '0.05'}}, #max(Para_dict['M2'][ParserIn['Section_M2'][0]][max(ParserIn['W'])]['M2'])}}, 
                '4' : {'Title' : Para2 + ' all sections',
                            'x':  {'name' : 'Width (m)',
                                   'min': max(Para_dict['Vy'][ParserIn['Section_Vy'][0]][max(ParserIn['W'])]['X'])/3,
                                   'max': max(Para_dict['Vy'][ParserIn['Section_Vy'][0]][max(ParserIn['W'])]['X'])*2/3},
                            'y':  {'name' : Para2 + ' (kN/m)',
                                   'minor_unit': '0.01', #min(Para_dict['M2'][ParserIn['Section_M2'][0]][max(ParserIn['W'])]['M2']),
                                   'major_unit': '0.05'}}}

    Tamplate_dict = {'A'+ str(TableForm['y_head']+1): 'Y=',
                     'A'+ str(TableForm['Para_avg']+1): 'w',
                     'A'+ str(TableForm['Hand_calc']) : 'Hand Calculation',
                     'A'+ str(TableForm['Hand_calc']+1) : 'My/Vy',
                     'A'+ str(TableForm['q1q2']+1) : 'q_Dummy',
                     'A'+ str(TableForm['Eq_header']+1) : 'Eq.width (m)',
                   'A'+ str(TableForm['-ve_header']+3) : 'Eq.width (m)',
       _no2lt(col2+1)+ str(w_col*4+41) : 'Plot 1',
                 refcell['1']['Title'] : Para1 +'_X=' + str(ParserIn['Section_' + Para1][0]),
                 refcell['2']['Title'] : Para2 +'_X=' + str(ParserIn['Section_' + Para2][0])}
    
    for key in Tamplate_dict.keys():
        worksheet.write(key, Tamplate_dict[key], format_dict['head'])


    hidden_rows = [TableForm['Para_avg']+1 , TableForm['Eq_header']+1, TableForm['q1q2']+1, TableForm['Hand_calc']+1]

    for row in hidden_rows + list(range((TableForm['-ve_header']+1), (w_col+ TableForm['+ve_header']+1))):
        worksheet.set_row(row, None, None, {'hidden': True})
    
    chart5 = _multisecchart(ParserIn, Para_dict, Path, workbook, col1-1, col2, Para1, refcell['3'], TableForm)

    chart6 = _multisecchart(ParserIn, Para_dict, Path, workbook, col1-1, col2-1, Para2, refcell['4'], TableForm)
    
    col1 = col2
    (col2, dlen) = _fill_formula(ParserIn, Para_dict, worksheet, col2, refcell['1']['Title'], True, TableForm)
    chart3 = _secchart(ParserIn, Path, workbook, col1, col2, dlen, refcell['1'])

    col1 = col2+1
    (col2, dlen) = _fill_formula(ParserIn, Para_dict, worksheet, col2+1, refcell['2']['Title'], True, TableForm)
    chart4 = _secchart(ParserIn, Path, workbook, col1, col2, dlen, refcell['2'])

    
    worksheet.insert_chart( _no2lt(col2 ) + '2', chart1)
    worksheet.insert_chart( _no2lt(col2 +13) + '2', chart2)
    worksheet.insert_chart( _no2lt(col2 ) + '66', chart3)
    worksheet.insert_chart( _no2lt(col2 +13) + '66', chart4)
    worksheet.insert_chart( _no2lt(col2 ) + '96', chart5)
    worksheet.insert_chart( _no2lt(col2 +13) + '96', chart6)

    
    workbook.close()

