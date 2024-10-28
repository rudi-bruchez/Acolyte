import sys
import os
import uno
import unohelper
from com.sun.star.document import XEventListener
from com.sun.star.task import XJobExecutor
import uuid

class LO:
    def __init__(self, doc):
        self.doc_id_name = "Acolyte Document Id"

        self.doc = doc
        # self.context = uno.getComponentContext()
        # self.desktop = self.context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.context)
        # self.frame = self.desktop.getCurrentFrame()
        # self.dispatcher = self.frame.getDispatcher()
        # self.model = self.frame.getController().getModel()
        # self.current_controller = self.frame.getController()
        # self.current_controller_name = self.current_controller.getName()
        # self.current_controller_frame = self.current_controller.getFrame()
        # self.current_controller_frame_name = self.current_controller_frame.getName

    def addDocumentGuid(self) -> str:
        guid = str(uuid.uuid4())

        udp = self.doc.DocumentProperties.UserDefinedProperties

        if not udp.getPropertySetInfo().hasPropertyByName(self.doc_id_name):
            udp.addProperty(self.doc_id_name,128,"")

        udp.setPropertyValue(self.doc_id_name, guid)

        return guid

    def getDocumentGuid(self) -> str:
        udp = self.doc.DocumentProperties.UserDefinedProperties
        guid = udp.getPropertyValue(self.doc_id_name)

        return guid