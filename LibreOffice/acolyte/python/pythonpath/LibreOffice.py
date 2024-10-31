# -*- coding: utf-8 -*-
import sys, os, re, string
import uno, unohelper
from string import ascii_lowercase, whitespace
from com.sun.star.document import XEventListener
from com.sun.star.task import XJobExecutor
from com.sun.star.awt.FontSlant import ITALIC
from com.sun.star.beans import PropertyValue  

class LO:
    def __init__(self, doc):
        self.doc_id_name = "Acolyte Document Id"

        self.doc = doc
        self.context = uno.getComponentContext()
        # self.desktop = self.context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.context)
        # self.frame = self.desktop.getCurrentFrame()
        # self.dispatcher = self.frame.getDispatcher()
        # self.model = self.frame.getController().getModel()
        # self.current_controller = self.frame.getController()
        # self.current_controller_name = self.current_controller.getName()
        # self.current_controller_frame = self.current_controller.getFrame()
        # self.current_controller_frame_name = self.current_controller_frame.getName

    @property
    def UserFolder(self) -> str:
        # obtenir le répertoire utilisateur
        user_folder = self.context.getServiceManager().createInstanceWithContext('com.sun.star.util.PathSubstitution', 
            self.context).getSubstituteVariableValue("user")
        return uno.fileUrlToSystemPath(user_folder)

    @property
    def AcolyteUserFolder(self) -> str:
        # obtenir le répertoire utilisateur Acolyte
        uf = os.path.join(self.UserFolder, 'Acolyte')
        if not os.path.exists(uf):
            os.mkdir(uf)
        return uf

    def getXPreviousParagraphs(self, num: int) -> str:
        vc = self.doc.CurrentController.ViewCursor.getStart()
        tc = self.doc.Text.createTextCursorByRange(vc)

        if not tc.isEndOfParagraph:
            tc.gotoEndOfParagraph(False)

        for i in range(num):
            tc.gotoPreviousParagraph(True)

        return tc.String
        # return vc.getString()

    def getChapterFromParagraph(self) -> str:
        vc = self.doc.CurrentController.ViewCursor.getStart()
        tc = self.doc.Text.createTextCursorByRange(vc)

        if not tc.isEndOfParagraph:
            tc.gotoEndOfParagraph(False)

        paragraphs = []

        while 1 == 1:
            tc.gotoPreviousParagraph(False)
            tc.gotoEndOfParagraph(True)
            p = tc.String
            ol = tc.getPropertyValue("OutlineLevel")
            if ol > 0:
                p = '#'*ol + ' ' + p
            paragraphs.insert(0, p)
            # if tc.ParaStyleName == "Heading 1":
            #     break
            # if tc.getPropertyValue("OutlineLevel") == 0:
            #     paragraphs.append(tc.String)
            if tc.getPropertyValue("OutlineLevel") == 1:
                break

        return '\n\n'.join(paragraphs)

    # def getTOC(self) -> str:
    #     cur = self.doc.Text.createTextCursor()
    #     cur.gotoStart(False)

    def getTOC(self) -> str:
        toc = self.doc.createInstance('com.sun.star.text.ContentIndex')
        toc.setPropertyValue("Level", 10)  # Include all heading levels
        toc.CreateFromOutline = True
        # toc.update()
        return toc

    # def GetHeadings(self):
    #     # https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1style_1_1ParagraphStyleCategory.html

    # def changeLanguage(self):
    #     dim dispatcher as object
    #     document   = ThisComponent.CurrentController.Frame
    #     dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    #     dim args1(0) as new com.sun.star.beans.PropertyValue
    #     args1(0).Name = "Language"
    #     args1(0).Value = "Default_English (UK)"

    #     dispatcher.executeDispatch(document, ".uno:LanguageStatus", "", 0, args1())

    def ReadParagraphs(self):
        parenum = self.doc.Text.createEnumeration()

        while parenum.hasMoreElements():
            par = parenum.nextElement()
            # check par.ParaStyleName here for Heading or body text or any other paragraph styling
            textenum = par.createEnumeration()
            while textenum.hasMoreElements() :
                text = textenum.nextElement()
                # check text.CharPosture and text.CharWeight and other properties here
                print(text.getString())

    def GetChapters(self):
        cur = self.doc.getCurrentController().getTextCursor()

        search = self.doc.createSearchDescriptor
        search.setPropertyValue("SearchStyles", True)
        search.setSearchString("Heading 1")
        FoundRanges = self.doc.findAll(search)
        PickedRange = FoundRanges.getByIndex(2)
        cur.gotoRange(PickedRange, False)  # False means the cursor does not expand when it moves
        # Print "Heading is on page " + ViewCurs.Page

    # def getWord(self) -> str:
    #     Cursor = ThisComponent.Text.createTextCursor()
    #     Cursor.gotoStart(False)
    #     Cursor.gotoEndOfWord(True)
    #     msgbox Cursor.getString()

    def SearchRegex(self, regex):
        search = self.doc.createSearchDescriptor()
        search.SearchString = regex
        search.SearchRegularExpression = True
        search.SearchAll = True
        search.SearchCaseSensitive = True
        search.SearchWords = False

        return self.doc.findAll(search)

    def ImportDocument(self):
        pathname = os.path.dirname(self.doc.getURL())
        importCursor = self.doc.Text.createTextCursor()

        selsFound = self.SearchRegex(self.doc, "\{includetext:.*\}")

        while selsFound.getCount() > 0:
            for selIndex in range(0, selsFound.getCount()):
                selFound = selsFound.getByIndex(selIndex)
                importCursor.gotoRange(selFound, False)
                filename = re.sub(r'{includetext:(.*)}', r'\1', selFound.getString())
                filename = filename.replace("${path}", pathname)

                prop = PropertyValue()
                prop.Name = "DocumentTitle"
                prop.Value = ""
                importCursor.insertDocumentFromURL(filename, (prop,))

            selsFound = self.SearchRegex(self.doc, "\{includetext:.*\}")

    def SelectNextSentence(self, event=None):
        controller = self.doc.CurrentController
        T = self.doc.Text
        cur = T.createTextCursor()
        cur.gotoRange(controller.ViewCursor.End, False)
        if cur.isEndOfSentence():
            cur.gotoNextSentence(False)
        else:
            cur.gotoStartOfSentence(False)
        cur.gotoEndOfSentence(True)
        controller.select(cur)

    def SelectPreviousSentence(self, event=None):
        controller = self.doc.CurrentController
        T = self.doc.Text
        cur = T.createTextCursor()
        cur.gotoRange(controller.ViewCursor.End, False)
        if cur.isEndOfSentence():
            cur.gotoPreviousSentence(False)
        else:
            cur.gotoStartOfSentence(False)
        cur.gotoEndOfSentence(True)
        controller.select(cur)

    def InsertTextIntoCell(self, table, cellName, text, color ):
        tableText = table.getCellByName( cellName )
        cursor = tableText.createTextCursor()
        cursor.setPropertyValue( "CharColor", color )
        tableText.setString( text )

# localContext = uno.getComponentContext()
# resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
# context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
# desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
# model = desktop.getCurrentComponent()
# text = model.Text
# cursor = text.createTextCursor()
# texto = ""
# while (texto != "quit"):
#     texto = input("Enter text (type 'enter' for carry return or 'quit' to exit): ")
#     if (texto == "table"):
#         table = model.createInstance( "com.sun.star.text.TextTable" )
#         rows = input("How many rows: ")
#         columns = input("How many columns: ")
#         table.initialize(rows, columns)

#     if (texto == "enter"):
#         text.insertControlCharacter( cursor, PARAGRAPH_BREAK, 0 )
#     else:
#         if (texto != "quit"):
#             text.insertString(cursor, str(texto), 0)

    def SupprimerFauxRetours(self, event=None):
        '''Supprime les marques de paragraphes situées en milieu de phrase.
        Agit sur le document entier'''
        undo = self.doc.UndoManager
        undo.enterUndoContext('Supprimer faux retours')
        self.doc.lockControllers()
        try:
            self._supprimer_faux_retours(self.doc)
            win = self.doc.CurrentController.ComponentWindow
            win.Toolkit.createMessageBox(win, 0, 1, "", "Ok !").execute()
        finally:
            self.doc.unlockControllers()
            undo.leaveUndoContext()


    def SupprimerFauxRetoursSelection(self, event=None):
        '''Supprime les marques de paragraphes situées en milieu de phrase.
        Agit sur la sélection uniquement.'''
        undo = self.doc.UndoManager
        undo.enterUndoContext('Supprimer faux retours (sélection)')
        self.doc.lockControllers()
        try:
            for selection in self.doc.CurrentSelection:
                T = selection.Text
                tcursor = T.createTextCursorByRange(selection)
                tcursor.gotoEndOfSentence(False)
                while tcursor.gotoNextSentence(True):
                    if T.compareRegionEnds(tcursor, selection) < 0:
                        break
                    if tcursor.String.startswith('\r\n'):
                        s = ' '
                        while s in whitespace:
                            tcursor.goRight(1, True)
                            s = tcursor.String[-1]
                        if s in ascii_lowercase:
                            tcursor.String = ' ' + s
                    tcursor.gotoEndOfSentence(False)
                for p in selection.createEnumeration():
                    try:
                        if p.supportsService('com.sun.star.text.TextTable'):
                            cellnames = p.CellNames
                            for cellname in cellnames:
                                cell = p.getCellByName(cellname)
                                self.__SupprimerFauxRetours(cell)
                    except AttributeError:
                            pass
                for textcontent in selection.createContentEnumeration('com.sun.star.text.TextContent'):
                    if textcontent.supportsService('com.sun.star.text.TextFrame'):
                        self.__SupprimerFauxRetours(textcontent)
                win = doc.CurrentController.ComponentWindow
            win.Toolkit.createMessageBox(win, 0, 1, "", "Ok !").execute()
        finally:
            self.doc.unlockControllers()
            undo.leaveUndoContext()

    def __SupprimerFauxRetours(self, container):
        T = container.Text
        tcursor = T.createTextCursor()
        tcursor.gotoEndOfSentence(False)
        while tcursor.gotoNextSentence(True):
            if tcursor.String.startswith('\r\n'):
                s = ' '
                while s in whitespace:
                    tcursor.goRight(1, True)
                    s = tcursor.String[-1]
                if s in ascii_lowercase:
                    tcursor.String = ' ' + s
            tcursor.gotoEndOfSentence(False)
        try:
            for frame in container.TextFrames:
                self.__SupprimerFauxRetours(frame)
        except AttributeError:
            pass
        try:
            for table in container.TextTables:
                cellnames = table.CellNames
                for cellname in cellnames:
                    cell = table.getCellByName(cellname)
                    self.__SupprimerFauxRetours(cell)
        except AttributeError:
            pass

    def ListStyles(self) -> str:
        cur = self.doc.Text.createTextCursor()
        cur.gotoStart(False)
        styles = ''
        #loop for cursor
        while cur.gotoNextParagraph(False):
            styles = styles + cur.getPropertyValue("ParaStyleName") + '\n'
        return styles

    def EnsureStyles(self):
        name = "Acolyte Question"
        new_style = self.doc.createInstance('com.sun.star.style.ParagraphStyle')

        paragraph_styles = self.doc.getStyleFamilies()['ParagraphStyles']
        standard_props = paragraph_styles.getByName("Standard")

        if not paragraph_styles.hasByName(name):
            paragraph_styles.insertByName(name, new_style)

            properties = dict(
                CharFontName = standard_props.CharFontName,
                ParaBackColor = 998877 # rgb(200,200,0)
            )
            keys = tuple(properties.keys())
            values = tuple(properties.values())
            new_style.setPropertyValues(keys, values)

    def FindQuestions(self):
        parenum = self.doc.Text.createEnumeration()

        while parenum.hasMoreElements():
            par = parenum.nextElement()
            # check par.ParaStyleName here for Heading or body text or any other paragraph styling
            if par.ParaStyleName == "Acolyte Question":
                textenum = par.createEnumeration()
                while textenum.hasMoreElements() :
                    text = textenum.nextElement()
                    # check text.CharPosture and text.CharWeight and other properties here
                    print(text.getString())

    def SearchByFontStyle(self, styleName):
        search = self.doc.createSearchDescriptor()
        search.SearchString = "(.*)"
        search.SearchAll = True
        search.SearchWords = True
        search.SearchRegularExpression = True
        search.SearchCaseSensitive = False
        prop = PropertyValue('CharPosture', 0, ITALIC, 0)
        search.SearchAttributes = (prop,)
        search.ReplaceString = "abc"
        found = self.doc.replaceAll(search)

    def LoopQuestions(self):
        cur = self.doc.Text.createTextCursor()
        cur.gotoStart(False)

        while cur.gotoNextParagraph(False):
            if cur.getPropertyValue("ParaStyleName") == "Acolyte Question":
                question = cur.getString() 
                if cur.gotoNextParagraph(False):
                    while cur.getPropertyValue("ParaStyleName") == "Acolyte Question":
                        question += cur.getString() 
                        cur.gotoNextParagraph(False)
                # HERE : ask the question
                response = ""
                cur.setPropertyValue("ParaStyleName", "Standard")
                self.doc.Text.insertString(cur, response, False)

    #Define a function to get all the heading in text order
    def getAllHeadingsInTextOrder(self):
        heading_styles = ['Heading 1', 'Heading 2']
        headings_in_order = [] #list of heading_text, heading style

        cur = self.doc.Text.createTextCursor()
        cur.gotoStart(False)

        while True:
            current_heading_name = cur.getPropertyValue('ParaStyleName')
            if current_heading_name in heading_styles:
                #get the text used in the heading
                cur.gotoEndOfParagraph(True) #select the text
                heading_text = cur.getString()
                headings_in_order.append([heading_text, current_heading_name])
            if not cur.gotoNextParagraph(False): #have reached the end
                break
            return headings_in_order

    #Create a function to give the level of a heading
    #This level is used for creating hierarchical tables of contents
    def getOutlineLevel(self, heading_style_name):
        '''This will return 0 for default level, 1 for Heading 1 etc.'''
        heading_style = self.doc.getStyleFamilies().getByName('ParagraphStyles').getByName(heading_style_name)
        outline_level = heading_style.getPropertyValue('OutlineLevel')
        return outline_level

        # #Actually print out the headings, but with some pretty indenting according to level.
        # headings = getAllHeadingsInTextOrder()
        # print('_'*4+'Headings for '+document.getTitle()+'_'*4)
        # print('')
        # for heading_text, heading_style_name in headings:
        # level = getOutlineLevel(heading_style_name)
        # print((level-1)*'\t'+heading_text+ ' '+'('+heading_style_name+')')
