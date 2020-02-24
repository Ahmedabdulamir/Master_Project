import fdAPI_wrapper_mod as fd
import os


def Deck(x, y, t, xsupp, ysupp, xload, yload, dia, loadIntensity, materialName, mesh, poisson, Ext_list, close, filename, Section):
  # init model
  fd.initiateModel("S")
  direction = fd.coord(0,0,-1)
           

  # add material
  material = fd.addMaterial(fd.material(materialName, "0", "0", poisson))
  
  # add structural parts
  for i in range(0, len(x)-1):
      p0 = fd.coord(x[i],0,0)
      p1 = fd.coord(x[i+1],y,0)
      fd.addPlate(material, t[i], t[i+1], " ", p0, p1, "top", mesh)


  # add load cases
  dl = fd.addLoadCase("DL", False)
  ll = fd.addLoadCase("LL", False)


  # create load
  #fd.addLineLoad(loadIntensity, direction, ll, p0, p1)
  #i=0
  #for i in range(0, len(xload)-1):
    #  p0 = fd.coord(xload[i]-dia,yload[i]-dia,0)
    #  p1 = fd.coord(xload[i]+dia,yload[i]+dia,0)
      #fd.addSurfaceLoad(loadIntensity, direction,ll, p0, p1 )
  p0 = fd.coord(xload-dia,yload-dia,0)
  p1 = fd.coord(xload+dia,yload+dia,0)
  fd.addSurfaceLoad(loadIntensity, direction,ll, p0, p1 )


  # add load combinations
  fd.addLoadComb("LC1", "U", [dl, ll], [1.35, 1.5])

  # create points
  p0 = fd.coord(xsupp[0],0,0)
  p1 = fd.coord(xsupp[0],ysupp,0)
  p2 = fd.coord(xsupp[1],0,0)
  p3 = fd.coord(xsupp[1],ysupp,0)


  # create supports
  fd.addLineSupport(p0, p1, "hinged")
  fd.addLineSupport(p2, p3, "zpinned")


  #Verifying and creating working directory
  if not os.path.exists('temp/'+ str(filename)):
    os.mkdir('temp/'+ str(filename))


  # create struxml
  filePath = 'temp/'+ str(filename) +'/'+ str(filename) +'_model.struxml'
  fd.finish(filePath)
  

  #create bsc
  batchfile = ['temp/' + str(x) + '.bsc' for x in Ext_list]
  exportfile = ['temp/'+ str(filename) + '/' + str(filename) + str(x) + '.txt' for x in Ext_list]

  fd.runFD('LIN', False, close, 'no', filePath, batchfile, exportfile)   
  
  #open fd
  #fd.openFD(filePath)

  return



