# coding: utf-8
from __future__ import unicode_literals
import uno
import os
import unohelper
from com.sun.star.datatransfer import XTransferable, DataFlavor
from com.sun.star.beans import PropertyValue

class Transferable(unohelper.Base, XTransferable):
    """Keep clipboard data and provide them."""
  
    def __init__(self, data,df):
        self.flavors = [df]
        self.data = [data] 
  
    def getTransferData(self, flavor):
        if not flavor: return
        mtype = flavor.MimeType
        for i,f in enumerate(self.flavors):
            if mtype == f.MimeType:
                return self.data[i]
  
    def getTransferDataFlavors(self):
        return tuple(self.flavors)
  
    def isDataFlavorSupported(self, flavor):
        if not flavor: return False
        mtype = flavor.MimeType
        for f in self.flavors:
            if mtype == f.MimeType:
                return True
        return False

def copy_oneFormat(format):
    oDoc = XSCRIPTCONTEXT.getDocument()
    df = None
    selectedContent = oDoc.CurrentController.getTransferable()
    dataFlavors = selectedContent.getTransferDataFlavors()
    for dataFlavor in dataFlavors:
        #print(dataFlavor.HumanPresentableName)
        if format in dataFlavor.HumanPresentableName.upper():
            df=dataFlavor
    if df is None:
        return None
    oData = selectedContent.getTransferData(df)
    transferable = Transferable(oData,df)
    #copie dans le presse-papier pour vérification
    ctx = XSCRIPTCONTEXT.getComponentContext()
    oClip = ctx.getServiceManager().createInstanceWithContext(
        "com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)
    oClip.setContents(transferable, None)
    return transferable


def copy_rtf():
    copy_oneFormat("RICH TEXT")

def copy_png():
    copy_oneFormat("PNG")

def copy_html():
    copy_oneFormat("HTML")

# copier le contenu d'un commentaire dans une cellule :
def copy_comment(cellule,destination):
    CTX = uno.getComponentContext()
    SM = CTX.getServiceManager()
    doc = XSCRIPTCONTEXT.getDocument()
    oSheet = doc.getSheets().getByName( "Feuille1" )
    doc.getCurrentController().select(oSheet.getCellRangeByName( cellule ))
    frame = doc.getCurrentController().getFrame()
    dispatch = SM.createInstance('com.sun.star.frame.DispatchHelper')
    #Edition commentaire + Selection avec OpenOffice
    #dispatch.executeDispatch(frame, ".uno:EditAnnotation", "", 0, ())
    #Edition commentaire + Selection avec LibreOffice
    dispatch.executeDispatch(frame, ".uno:InsertAnnotation", "", 0, ())
    dispatch.executeDispatch(frame, ".uno:SelectAll", "", 0, ())
    #transferable = copy_oneFormat("RICH TEXT")
    transferable = doc.CurrentController.getTransferable()
    dispatch.executeDispatch(frame,".uno:DrawEditNote", "", 0, ())  
    dispatch.executeDispatch(frame, ".uno:HideNote", "", 0, ())
    doc.getCurrentController().select(oSheet.getCellRangeByName(destination ))
    if transferable is not None:
        doc.getCurrentController().insertTransferable(transferable)
    
# copier le contenu d'un commentaire dans une cellule :
def modify_comment(cellule):
    CTX = uno.getComponentContext()
    SM = CTX.getServiceManager()
    doc = XSCRIPTCONTEXT.getDocument()
    oSheet = doc.getSheets().getByName( "Feuille1" )
    doc.getCurrentController().select(oSheet.getCellRangeByName( cellule ))
    frame = doc.getCurrentController().getFrame()
    dispatch = SM.createInstance('com.sun.star.frame.DispatchHelper')
    dispatch.executeDispatch(frame, ".uno:Copy", "", 0, ())
    #transferable = doc.CurrentController.getTransferable()
    #transferable = copy_oneFormat("RICH TEXT")
    #dispatch.executeDispatch(frame, ".uno:EditAnnotation", "", 0, ())
    dispatch.executeDispatch(frame, ".uno:InsertAnnotation", "", 0, ())
    dispatch.executeDispatch(frame, ".uno:SelectAll", "", 0, ())
    #if transferable is not None:
        #doc.getCurrentController().insertTransferable(transferable)
    p=PropertyValue()
    p.Name = 'Format'
    p.Value = 10
    args = (p,)
    dispatch.executeDispatch(frame,".uno:PasteSpecial", "", 0, args)
    dispatch.executeDispatch(frame,".uno:DrawEditNote", "", 0, ())  
    dispatch.executeDispatch(frame, ".uno:HideNote", "", 0, ())

# Functions that can be called from Tools -> Macros -> Run Macro.
g_exportedScripts = copy_rtf, copy_html, copy_png,