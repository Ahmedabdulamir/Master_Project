import pandas as pd
import os
import x_section_w
import y_intergration_w
import handcal 


def _Sort(df, by):
    df.sort_values(by=[by], inplace = True)
    df.reset_index(drop=True, inplace=True)   
    return df 
   
def _Coordata(coor_df, ParserIn, Sec):
    #Reading the FEM design 19 coordinates batchfile intothe node-coordinates dataframe (coor_df)    
    Index= ParserIn.get(ParserIn.get('Direction'))
    meshsize = ParserIn.get('Mesh')

    node_df = coor_df.loc[abs(coor_df[Index[0]]-Sec) <= (meshsize)]         #change the column number 1:X 2:Y      (meshsize/2)0.5= torellence
    #print(node_df)   
    return (node_df)


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
    #print(result_df)
    return(result_df)


def Analyze(Path, ParserIn, FEMIn):
    #initilizing the dataframes
    Index = ParserIn.get(ParserIn.get('Direction')) 
    w = ParserIn.get('W')
    #nod_dic =_Coordata(Path, ParserIn)
    Para_dict={}
    coor_df=pd.read_csv('./temp/'+ Path +'/'+ Path +'_coor.txt',sep='	', header=None, engine='python')
    for Para in ParserIn.get('Parameter'):
        load_df=pd.read_csv('./temp/'+ Path +'/'+ Path + '_' + Para + '.txt',sep='	', header=None, engine='python')
        data_dict={}
       
        for Sec in (ParserIn.get('Section_'+ Para)): 
            para_df1 = _Secdata(_Coordata(coor_df, ParserIn, Sec), load_df, Index, ParserIn.get(Para))
            w_dict= {}
            
            for W in w:
                sec_dict={}
                (No, X, Y, P) = x_section_w.section_interp(Sec, para_df1, W)
                sec_dict[Para] = list(P)
                sec_dict['X'] = list(X)
                sec_dict['Y'] = list(Y)
                sec_dict[Para + '_avg'] = y_intergration_w.averrage (X, P, W)
                sec_dict[Para + '_h'] = handcal.hand_calculation(FEMIn, Para, Sec)
                w_dict[W] = sec_dict
            data_dict[Sec] = w_dict
        Para_dict[Para] = data_dict
    return(Para_dict)            





   