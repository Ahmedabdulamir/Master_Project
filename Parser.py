import pandas as pd
import os
import x_section_w
import y_intergration_w
import xlsxwriter
import openpyxl as xl
import shutil
import handcal 

def _Sort(df, by):
    df.sort_values(by=[by], inplace = True)
    df.reset_index(drop=True, inplace=True)   
    return df 
   
def _Coordata(Path, ParserIn):
    #Reading the FEM design 19 coordinates batchfile intothe node-coordinates dataframe (coor_df)    
    Index= ParserIn.get(ParserIn.get('Direction'))
    meshsize = ParserIn.get('Mesh')

    coor_df=pd.read_csv('./temp/'+ Path +'/'+ Path +'_coor.txt',sep='	', header=None, engine='python')
    nod_dict = {}

    for Sec in (ParserIn.get('Section')):    
        node_df = coor_df.loc[abs(coor_df[Index[0]]-Sec) <= (meshsize)]         #change the column number 1:X 2:Y      (meshsize/2)0.5= torellence
        nod_dict[Sec] = node_df
    
    return (nod_dict)


def _Secdata(node_df, load_df, Index1, Index2):            
    extract_load_df = pd.DataFrame()   #where the nodes corrdinates are stored   
    filtered_nodes_df = pd.DataFrame()
    #Extracting the related nodes data
    node_df_list = [x for x in node_df[0].tolist() if x != 'nan']
    extract_load_df = extract_load_df.append(load_df[load_df[2].isin(node_df_list)] , ignore_index = True)
    
    #Filtering the node-coordinates dataframe according to the extracted node-load dataframe
    filtered_nodes_df = filtered_nodes_df.append(node_df[node_df[0].isin(extract_load_df[2].tolist())], ignore_index = True)
    filtered_nodes_df = _Sort(filtered_nodes_df, 0)
    extract_load_df = _Sort(extract_load_df, 2)

    result_df = pd.concat([extract_load_df[2], extract_load_df[Index2],filtered_nodes_df[0], 
                            filtered_nodes_df[Index1[1]], filtered_nodes_df[Index1[0]]], axis=1)
    result_df.columns = range(result_df.shape[1])
    result_df = _Sort(result_df, 3)
    print(result_df)
    return(result_df)


def Analyze(Path, ParserIn, FEMIn):
    #initilizing the dataframes
    Index = ParserIn.get(ParserIn.get('Direction')) 
    w = ParserIn.get('W')
    nod_dic =_Coordata(Path, ParserIn)
    Para_dict={}
    for Para in ParserIn.get('Parameter'):
        load_df=pd.read_csv('./temp/'+ Path +'/'+ Path + '_' + Para + '.txt',sep='	', header=None, engine='python')
        data_dict={}
        for Sec in (ParserIn.get('Section')): 
            para_df1 = _Secdata(nod_dic.get(Sec), load_df, Index, ParserIn.get(Para))
            w_dict= {}
            for W in w:
                sec_dict={}
                (No, X, Y, P) = x_section_w.section_interp(Sec, para_df1, W)
                sec_dict[Para] = list(P)
                sec_dict['X'] = list(X)
                sec_dict['Y'] = list(Y)
                sec_dict[Para + '_avg'] = y_intergration_w.averrage (Y, P, W)
                sec_dict[Para + '_h'] = handcal.hand_calculation(FEMIn, Para, Sec)
                w_dict[W] = sec_dict
            data_dict[Sec] = w_dict
        Para_dict[Para] = data_dict
    return(Para_dict)            


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


def xlsExcel(Path, Para_dict, ParserIn):

    Direction = ParserIn.get('Direction')
    _chkfile(Path, ParserIn)
    workbook = xlsxwriter.Workbook('xlsx/' + Path + '/'+ Path + Direction + '.xlsx', {'strings_to_numbers': True})
    worksheet = workbook.add_worksheet()
    cell_format01 = workbook.add_format()
    cell_format02 = workbook.add_format()
    #chart5 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
    #chart2 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
    #chart2.set_style(15)
    
    col = 0
    col1 = 1
    cell_format01.set_pattern(1)
    cell_format01.set_bg_color('green')
    cell_format02.set_pattern(1)
    cell_format02.set_bg_color('yellow')

    format_dict = {'M2' : cell_format01,'Vy' : cell_format02}
    for Para in ParserIn.get('Parameter'):
        sec_dict = Para_dict.get(Para)
        for Sec in sec_dict.keys():
            w_dict = sec_dict.get(Sec)
            worksheet.set_column(_no2lt(col1+1) + ':' + _no2lt(col1+9), None, None, {'hidden': True})
            worksheet.write(0, col1, Para + '_avg')
            worksheet.write(7, col1, w_dict[ParserIn['W'][0]][Para + '_h'])
            row = 1
            for W in w_dict.keys():
                data_dict = w_dict.get(W)
                worksheet.write_column(1, 0, ParserIn['W'])
                worksheet.write(15, col, 'Y_w=' + str(W))
                worksheet.write_column(16, col, data_dict['Y'])
                worksheet.write(15, col+1, Para +'_X=' + str(Sec))
                worksheet.write_column(16, col+1, data_dict[Para],format_dict[Para])
                worksheet.write(row, col1, data_dict[Para + '_avg'])
                worksheet.write_column(9, 0, ParserIn['W'])
                worksheet.write(row+8, col1, data_dict[Para + '_h']/data_dict[Para + '_avg'] )
                row+=1
                col+=2
            worksheet.write(9, col1, Sec)
            col1 = col+1
    # for j in range (2 , length, 8):
    #     col0 = xlsxwriter.utility.xl_col_to_name(j-1) '=' + _no2lt(col1) + str(row+2) + '/' + _no2lt(col1) + str(8)
    #     col1 = xlsxwriter.utility.xl_col_to_name(j)
    #     col2 = xlsxwriter.utility.xl_col_to_name(j+1)
    #     col3 = xlsxwriter.utility.xl_col_to_name(j+4)
    #     chart2.add_series({'categories': '=Sheet1!$'+ col2 +'$2:$'+ col2 +'$'+ str(len(data.columns)),
    #                            'values': '=Sheet1!$'+ col1 +'$2:$'+ col1 +'$'+ str(len(data.columns)),
    #                              'name': '=Sheet1!$'+ col0 +'$1'})
    #     chart2.add_series({'categories': '=Sheet1!$'+ col2 +'$2:$'+ col2 +'$'+ str(len(data.columns)),
    #                            'values': '=Sheet1!$'+ col3 +'$2:$'+ col3 +'$'+ str(len(data.columns)),
    #                              'name': '=Sheet1!$'+ col0 +'$1'})
       

    # worksheet.insert_chart('H1', chart2)

    workbook.close()
    # Add the worksheet data to be plotted.
    
    # Create a new chart object.
   

    # Add a series to the chart.
   

    # Insert the chart into the worksheet.
    



   