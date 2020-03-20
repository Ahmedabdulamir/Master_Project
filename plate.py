import fdAPI_wrapper_mod as fd
import os


def Deck(Path, ParserIn, FEMIn, close):
  # init model
  fd.initiateModel("S")
  direction = fd.coord(0,0,-1)
           

  # add material
  material = fd.addMaterial(fd.material(FEMIn['materialName'], "0", "0", FEMIn['poisson']))
  
  # add structural parts
  x = FEMIn['x']
  y = FEMIn['y']
  for i in range(0, len(y)-1):
      p0 = fd.coord(0,y[i],0)
      p1 = fd.coord(x,y[i+1],0)
      fd.addPlate(material, FEMIn['t'][i], FEMIn['t'][i+1], " ", p0, p1, "top", str(ParserIn['Mesh']))


  # add load cases
  dl = fd.addLoadCase("DL", False)
  ll = fd.addLoadCase("LL", False)


  # create load
  p0 = fd.coord(FEMIn['xload'] - FEMIn['dia'], FEMIn['yload'] - FEMIn['dia'] ,0)
  p1 = fd.coord(FEMIn['xload'] + FEMIn['dia'], FEMIn['yload'] + FEMIn['dia'] ,0)
  fd.addSurfaceLoad(FEMIn['loadIntensity'], direction,ll, p0, p1 )


  # add load combinations
  fd.addLoadComb("LC1", "U", [dl, ll], [1, 1])

  # create points
  p0 = fd.coord(0,FEMIn['ysupp'][0],0)
  p1 = fd.coord(FEMIn['xsupp'],FEMIn['ysupp'][0],0)
  p2 = fd.coord(0,FEMIn['ysupp'][1],0)
  p3 = fd.coord(FEMIn['xsupp'],FEMIn['ysupp'][1],0)
 

  # create supports
  fd.addLineSupport(p0, p1, "hinged")
  fd.addLineSupport(p2, p3, "zpinned")


  #Verifying and creating working directory
  if not os.path.exists('temp/'+ str(Path)):
    os.mkdir('temp/'+ str(Path))


  # create struxml
  filePath = 'temp/'+ str(Path) +'/'+ str(Path) +'_model.struxml'
  fd.finish(filePath)
  

  #create bsc
  batchfile = ['BSC/' + str(x) + '.bsc' for x in FEMIn['Ext_list']]
  exportfile = ['temp/'+ str(Path) + '/' + str(Path) + str(x) + '.txt' for x in FEMIn['Ext_list']]

  fd.runFD('LIN', False, close, 'no', filePath, batchfile, exportfile)   
  
  #open fd
  #fd.openFD(filePath)

  return



