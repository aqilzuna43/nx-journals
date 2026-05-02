import sys
import NXOpen

def main():
    session = NXOpen.Session.GetSession()
    lw = session.ListingWindow
    lw.Open()
    lw.WriteFullline("Python version: " + sys.version)
    lw.WriteFullline("Executable: " + str(sys.executable))
    lw.WriteFullline("sys.path:")
    for p in sys.path:
        lw.WriteFullline("  " + str(p))

if __name__ == "__main__":
    main()


### Output 
#Python version: 3.10.15 (main, Sep 30 2024, 11:52:23) [MSC v.1929 64 bit (AMD64)]
#Executable: C:\Program Files\Siemens\NX2312\nxbin\ugraf.exe
#sys.path:
  #C:\Program Files\Siemens\NX2312\NXBIN\python\from_git
  #C:\Program Files\Siemens\NX2312\nxbin\python
  #C:\Program Files\Siemens\NX2312\nxbin\python\Python310.zip
  #C:\Program Files\Siemens\NX2312\design_tools\checkmate\python
  #C:\Program Files\Siemens\NX2312\automated_testing\python
  #C:/Program Files/Siemens/NX2312/NXBIN/python/from_git