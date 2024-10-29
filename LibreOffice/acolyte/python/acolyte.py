# -*- coding: utf-8 -*-
import os, sys
import uno, unohelper
from com.sun.star.document import XEventListener
from com.sun.star.task import XJobExecutor
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_core.output_parsers import StrOutputParser
from langgraph.checkpoint.memory import MemorySaver
from langgraph.graph import END, START, StateGraph, MessagesState
from langgraph.prebuilt import ToolNode
from pathlib import Path

from LibreOffice import LO
from Agent import Agent

sys.stderr = sys.stdout

# class Acolyte(unohelper.Base, XJobExecutor, XEventListener):
class Acolyte(unohelper.Base, XJobExecutor):
    def __init__(self, context):
        """
        Initialize the Acolyte class.

        Args:
            context: The component context.
        """
        self.context = context
        self.desktop = self.createUnoService("com.sun.star.frame.Desktop")
        self.dispatchhelper = self.createUnoService("com.sun.star.frame.DispatchHelper")
        self.doc = self.desktop.getCurrentComponent()

        # Langchain and Langgraph
        api_key = 'XXX'
        os.environ["OPENAI_API_KEY"] = api_key
        # self.client = ChatOpenAI(model="gpt-4o-mini")
        self.client = Agent(self.doc)

    def trigger(self, command):
        """
        Called when the service is called from a menu or keyboard shortcut.

        Args:
            command (str): The command that was called.
        """
        if command == "getConfiguration":
            self.getConfiguration()
        elif command == "insertPrompt":
            self.insertPrompt()
        elif command == "rewrite":
            self.rewrite()
        elif command == "preparePrompt":
            self.preparePrompt()
        elif command == "debug":
            self.debug()
        elif command == "createDocumentFromPrompt":
            self.createDocumentFromPrompt()

    @property
    def cursor(self):
        ctr = self.doc.CurrentController
        return ctr.ViewCursor

    def formatMarkdown(self):
        rd = self.doc.createReplaceDescriptor()
        rd.SearchRegularExpression = True
        rd.SearchString = '^# '  # search for a line starting with '# '
        rd.setReplaceAttributes('Style', 'Heading 1')
        rd.ReplaceString = "$1"
        self.doc.replaceAll(rd)

    def show_message_box(self, title, message):
        """
        Show a message box with the specified title and message.

        Args:
            title (str): The title of the message box.
            message (str): The content of the message box.
        """
        frame = self.desktop.ActiveFrame
        window = frame.ContainerWindow
        window.Toolkit.createMessageBox(
            window,
            uno.Enum("com.sun.star.awt.MessageBoxType", "INFOBOX"),
            uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_OK"),
            title,
            message,
        ).execute()

    def preparePrompt(self):
        prompt = """Type de document à générer : {livre, support de cours}
        Titre du document : {Titre du document}
        """

        ctr = self.doc.CurrentController
        cur = ctr.ViewCursor
        cur.setString(prompt)

    def insertPrompt(self):
        cur = self.cursor
        texte = cur.getString()
        print(texte)

        # cur.String = f"{texte}\n"
        cur.gotoEnd(False)

        response = self.client.GetSimpleResponse(texte)

        # cur.String = reponse.choices[0].message.content
        model = self.desktop.getCurrentComponent()
        text = model.Text
        cursor = text.createTextCursor()
        text.insertString(cursor, response, 0)
        # cur.String = reponse.choices[0].message.content
        # cur.gotoEnd(True)

    def rewrite(self):
        cur = self.cursor
        texte = cur.getString()

        cur.String = f"{texte}\n\n"
        cur.gotoEnd(False)

        response = self.client.GetSimpleResponse(
            texte,
            "modife le texte suivant en le rendant plus professionnel et concis. Propose trois version différentes"
            )

        cur.String = response
        cur.gotoEnd(True)

    def createDocumentFromPrompt(self):
        cur = self.cursor
        texte = cur.getString()

        response = self.client.GetSimpleResponse(texte)

        # Create a new Writer document
        new_doc = self.desktop.loadComponentFromURL("private:factory/swriter", 
                                                    "_blank", 0, ())

        # Get the text object of the document
        text = new_doc.getText()
        cursor = text.createTextCursor()

        # Insert some sample text (optional)
        text.insertString(cursor, response, 0)

        return new_doc

    def debug(self):
        cur = self.cursor
        # texte = cur.getString()

        # Get the configuration provider
        # config_provider = self.context.ServiceManager.createInstanceWithContext(
        #     "com.sun.star.configuration.ConfigurationProvider", self.context
        # )

        # # Access the PathSettings configuration
        # path_settings = config_provider.createInstanceWithArguments(
        #     "com.sun.star.configuration.ConfigurationAccess",
        #     (uno.createUnoStruct("com.sun.star.beans.NamedValue", "nodepath", "/org.openoffice.Office.Paths"),)
        # )

        # # Get the User-specific configuration path
        # self.user_config_path = path_settings.getByName("User")

        # dir_path = os.path.dirname(os.path.realpath(__file__))

        # this_doc_url = self.doc.URL
        # this_doc_sys_path = uno.fileUrlToSystemPath(this_doc_url)
        # this_doc_parent_path = Path(this_doc_sys_path).parent

        lo = LO(self.doc)
        guid = lo.addDocumentGuid()


        # Print the path
        # self.show_message_box("Configuration Path", self.user_config_path)
        # print(f"LibreOffice configuration directory: {self.user_config_path}")
        # cur.String = f"""
        #     - LibreOffice configuration directory: {self.user_config_path}
        #     - Current directory: {dir_path}
        # """
        # cur.String = f"""
        #     - Current directory: {dir_path}
        #     - Current document path: {this_doc_parent_path}
        #     - Document Acolyte Id: {guid}
        # """

        cur.String = f"""
            - Document Acolyte Id: {guid}
        """


    # boilerplate code below this point
    def createUnoService(self, name):
        """
        Create an instance of the specified UNO service.

        Args:
            name (str): The name of the service to create.

        Returns:
            object: The created UNO service.
        """
        return self.context.ServiceManager.createInstanceWithContext(name, self.context)

    def disposing(self, _):
        """
        Empty disposing method for the XEventListener interface.

        Args:
            _ (object): The source object of the event.
        """
        pass

    def notifyEvent(self, _):
        """
        Empty notifyEvent method for the XEventListener interface.

        Args:
            _ (object): The event object.
        """
        pass


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    Acolyte,
    "com.pachadata.libreoffice.Acolyte",
    ("com.sun.star.task.JobExecutor",),
)