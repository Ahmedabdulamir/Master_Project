import pandas as pd
import os
import x_section_w
import x_section
import y_intergration_w
import xlsxwriter
import openpyxl as xl
import shutil

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
    return(result_df)


def Analyze(Path, ParserIn):
    #initilizing the dataframes
    Index1 = ParserIn.get(ParserIn.get('Direction')) 
    Parameter = ParserIn.get('Parameter')
    w = ParserIn.get('W')
    nod_dic =_Coordata(Path, ParserIn)
    
    Para_dict={}
    for Para in Parameter:
        load_df=pd.read_csv('./temp/'+ Path +'/'+ Path + '_' + Para + '.txt',sep='	', header=None, engine='python')
        data_dict={}
        for Sec in (ParserIn.get('Section')): 
            para_df1 = _Secdata(nod_dic.get(Sec), load_df, Index1, ParserIn.get(Para))
            w_dict= {}
            for W in w:
                sec_dict={}
                (No, X, Y, P) = x_section_w.section_interp(Sec, para_df1, W)
                P_avg = y_intergration_w.averrage (Y, P, W)
                sec_dict[Para] = list(P)
                sec_dict['X'] = list(X)
                sec_dict['Y'] = list(Y)
                sec_dict[Para + '_avg'] = float(P_avg)
                #sec_dict['P_h'] = _handcalc(Para, Sec)
                w_dict[W] = sec_dict
            data_dict[Sec] = w_dict
        Para_dict[Para] = data_dict
    return(Para_dict)            


def _xlsExcel(Section, Path, Direction, Moment, w, remove):
    
        
    data =pd.read_csv('temp/sections.csv',  sep=',', header=None, engine='python')
    if remove == True:
        os.remove('temp/sections.csv') 

    data.fillna('', inplace = True)
    #print(data)
    length = len(data.index)

    if not os.path.exists('xlsx/' + Path):
        os.mkdir('xlsx/' + Path)

    if os.path.isfile('xlsx/' + Path + '/w_' + str(w) + Direction + Moment + '.xlsx'):
        overwirte = input("File exists, overwirte?(Y,N) (Default = overwirte)")
        if overwirte == 'N':
            print('Please rename and close the file and retry again')
        else:
            while True:
                try:
                    os.remove('xlsx/' + Path + '/w_' + str(w) + Direction + Moment + '.xlsx')
                except IOError:
                    input('Please close Excelfile to continue')
                    continue
                break
    

    workbook = xlsxwriter.Workbook('xlsx/' + Path + '/w_' + str(w) + Direction + Moment + '.xlsx', {'strings_to_numbers': True})
    worksheet = workbook.add_worksheet()
    #chart5 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
    chart2 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
    chart2.set_style(15)
    
    for i in range (0 , length):
        col = xlsxwriter.utility.xl_col_to_name(i)
        pd.extdata = data.loc[(i),:]
        worksheet.write_column(col + '1', pd.extdata)
        

    for j in range (2 , length, 8):
        col0 = xlsxwriter.utility.xl_col_to_name(j-1)
        col1 = xlsxwriter.utility.xl_col_to_name(j)
        col2 = xlsxwriter.utility.xl_col_to_name(j+1)
        col3 = xlsxwriter.utility.xl_col_to_name(j+4)
        chart2.add_series({'categories': '=Sheet1!$'+ col2 +'$2:$'+ col2 +'$'+ str(len(data.columns)),
                               'values': '=Sheet1!$'+ col1 +'$2:$'+ col1 +'$'+ str(len(data.columns)),
                                 'name': '=Sheet1!$'+ col0 +'$1'})
        chart2.add_series({'categories': '=Sheet1!$'+ col2 +'$2:$'+ col2 +'$'+ str(len(data.columns)),
                               'values': '=Sheet1!$'+ col3 +'$2:$'+ col3 +'$'+ str(len(data.columns)),
                                 'name': '=Sheet1!$'+ col0 +'$1'})
       

    # Add the worksheet data to be plotted.
    
    # Create a new chart object.
   

    # Add a series to the chart.
   

    # Insert the chart into the worksheet.
    worksheet.insert_chart('H1', chart2)

    workbook.close()
    



    # def _Math(Section, Path, Direction, Moment, w, remove):   
    
    # # Calling the interpolation function
    # M_avg = [Moment + '_avg'] + [None]*len(w)
    # for i in range (0, len(w)):
    #     W = w[i]
    # # Selecting the columns to export in form of Lists as in the variables blow    
    #     M =  result_df[1].tolist()       #Moment/ Shear values
    #     Y =  result_df[3].tolist()       #Y coordinates
    #     No = result_df[0].tolist()      #Nodes
    #     X =  result_df[4].tolist()       #X coordinates
    #     (No, X, Y, M) = x_section_w.section_interp(Section, No, X, Y, M, W)
    #     _opExcel([Moment + '_X=' + str(X[1])] + list(M), 'M', path, Direction, Moment)
    #     _opExcel(['Y_w=' + str(W)]+ list(Y), 'Y', path, Direction, Moment)
    #     M_avg[i+1] = y_intergration_w.averrage (Y, M, W)
    # _opExcel(M_avg, 'M_avg', path, Direction, Moment)
    # return

    # def _handcalc(Moment, SecSection):    
    # calc = pd.read_csv('xlsx/Alpha0.5_Y0.4_X10.0/' + Moment + '.csv',  sep=',', header=None, engine='python')
    
    # tolerence = 0.01
    # if Moment == 'Vy':
    #     tolerence = 0.1
    # calc = calc.loc[abs(calc[1]-Section) <= (tolerence)] 
    # #print(calc)

    # (No_h, X_h, Y_h, M_h) = x_section.section_interp(Section, calc[3].tolist(), calc[1].tolist(), calc[2].tolist(), calc[0].tolist())
    # return(M_h)
