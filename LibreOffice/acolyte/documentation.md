# Short documentation for Python LibreOffice extensions

## Sources

+ [LibreOffice Python API](https://api.libreoffice.org/docs/idl/ref/index.html)
+ [LibreOffice Programming](https://flywire.github.io/lo-p/)
+ [python-ooo-dev-tools](https://python-ooo-dev-tools.readthedocs.io/))
+ [ooo-dev Part 2: Writer](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/index.html)
+ [Live LibreOffice Python UNO Examples](https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer)
+ [Live LibreOffice Python](https://github.com/Amourspirit/live-libreoffice-python)
+ [Professional UNO](https://wiki.documentfoundation.org/Documentation/DevGuide/Professional_UNO)

## Extracts

Each Office application (e.g. Writer, Draw, Impress, Calc, Base, Math) is supported by multiple modules (similar to Java packages). For example, most of Writer’s API is in Office’s “text” module, while Impress’ functionality comes from the “presentation” and “drawing” modules. These modules are located in com.sun.star package, which is documented at [API com.sun.star module reference](https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star.html).

Office is started as an OS process, and a Python program communicates with it via a socket or named pipe.

The invocation of Office and the setup of a named pipe link can be achieved with a single call to the soffice binary ( soffice.exe, soffice.bin ). A call starts the Office executable with several command line arguments, the most important being --accept which specifies the use of pipes or sockets for the inter-process link.

A call to XUnoUrlResolver.resolve() creates a remote component context, which acts as proxy for the ‘real’ component context over in the Office process (see Fig. 2). The context is a container/environment for components and UNO objects which I’ll explain below. When a Python program refers to components and UNO objects in the remote component context, the inter-process bridge maps those references across the process boundaries to the corresponding components and objects on the Office side.

## cursor

To do this in Writer in Python, here is an adaptation from the HelloWorld.py script included in the standard LO distribution:

def HelloWorldPython( ):
    #get the doc from the scripting context which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    #get the cursor
    ctrllr = model.CurrentController
    curview = ctrllr.ViewCursor
    curview.String = "Hello World (in Python at cursor)"
    return None

With the head start in the standard HelloWorld macro, it was easy to fill in the rest knowing how to do this in LO Basic is:

cntrllr = ThisComponent.CurrentController
vcur = cntrllr.ViewCursor
vcur.String = "Hello World "

Other useful methods of ViewCursor, which presumably are the same in Python are:

SbxSTRING getString ( void ) ; SbxVOID setString ( SbxSTRING ) ; 
SbxVOID collapseToStart ( void ) ; SbxVOID collapseToEnd ( void ) ; 
SbxBOOL isCollapsed ( void ) ; SbxBOOL goLeft ( SbxINTEGER, SbxBOOL ) ; 
SbxBOOL goRight ( SbxINTEGER, SbxBOOL ) ; SbxVOID gotoStart ( SbxBOOL ) ; 
SbxVOID gotoEnd ( SbxBOOL ) ; SbxVOID gotoRange ( SbxOBJECT, SbxBOOL ) ; 
SbxBOOL isVisible ( void ) ; SbxVOID setVisible ( SbxBOOL ) ; SbxOBJECT getPosition ( void );
SbxBOOL jumpToFirstPage ( void ) ; SbxBOOL jumpToLastPage ( void ) ; 
SbxBOOL jumpToPage ( SbxINTEGER ) ; SbxINTEGER getPage ( void ) ; 
SbxBOOL jumpToNextPage ( void ) ; SbxBOOL jumpToPreviousPage ( void ) ; 
SbxBOOL jumpToEndOfPage ( void ) ; SbxBOOL jumpToStartOfPage ( void ) ; 
SbxBOOL screenDown ( void ) ; SbxBOOL screenUp ( void ) ;
SbxBOOL goDown ( SbxINTEGER, SbxBOOL ) ; SbxBOOL goUp ( SbxINTEGER, SbxBOOL ) ;
SbxBOOL goLeft ( SbxINTEGER, SbxBOOL ) ; SbxBOOL goRight ( SbxINTEGER, SbxBOOL ) ;
SbxBOOL isAtStartOfLine ( void ) ; SbxBOOL isAtEndOfLine ( void ) ;
SbxVOID gotoEndOfLine ( SbxBOOL ) ; SbxVOID gotoStartOfLine ( SbxBOOL ) ; 

Other useful properties of ViewCursor: many formatting options and,

SbxOBJECT Start; SbxOBJECT End; SbxSTRING String; 
SbxOBJECT Position; SbxBOOL Visible; SbxINTEGER Page;

## Cursors

In the LibreOffice Writer object model, ViewCursor and TextCursor are both used to navigate and manipulate text, but they serve different purposes and have distinct characteristics:

### ViewCursor

+ Purpose: The ViewCursor represents the position of the visible cursor in the document as seen by the user. It is tied to the user interface and reflects the user’s current position in the document.
+ Usage: It is used for tasks that require interaction with the visible part of the document, such as moving the cursor to a specific location that the user can see or selecting text that the user has highlighted.
+ Scope: The ViewCursor is specific to the view controller of the document. It is part of the view and not the underlying text content.
+ Example: Moving the visible cursor to the end of the document:

```
Dim oDoc As Object
Dim oViewCursor As Object
oDoc = ThisComponent
oViewCursor = oDoc.CurrentController.ViewCursor
oViewCursor.gotoEnd(False)
```

### TextCursor

+ Purpose: The TextCursor is used for navigating and manipulating the text content within the document. It operates on the underlying text, independent of the user interface.
+ Usage: It is used for tasks like inserting, deleting, or formatting text within the document. It can navigate through the text content and perform operations that affect the document’s text structure.
+ Scope: The TextCursor operates on the Text object of the document. It is not tied to the visible cursor or the user interface.
+ Example: Inserting text at the end of the document:

```
Dim oDoc As Object
Dim oText As Object
Dim oTextCursor As Object
oDoc = ThisComponent
oText = oDoc.Text
oTextCursor = oText.createTextCursor()
oTextCursor.gotoEnd(False)
oText.insertString(oTextCursor, "This is the end of the document.", False)
```

### Key Differences

#### Interface vs. Content:

ViewCursor is linked to the user interface and represents what the user sees.
TextCursor is linked to the document content and can operate on text that might not be visible to the user.

#### Visibility:

Moving the ViewCursor affects the visible cursor position.
Moving the TextCursor does not affect the visible cursor position unless explicitly synchronized with the ViewCursor.

#### Functionality:

ViewCursor is used for tasks involving user interaction and visual feedback.
TextCursor is used for tasks involving text manipulation and document content changes.

#### Practical Example to Illustrate the Difference

Here’s an example where both cursors are used:

The ViewCursor moves to the start of the document.
The TextCursor inserts text at that position without moving the ViewCursor.

```
Sub ExampleUsingBothCursors()
    Dim oDoc As Object
    Dim oViewCursor As Object
    Dim oText As Object
    Dim oTextCursor As Object

    oDoc = ThisComponent
    oViewCursor = oDoc.CurrentController.ViewCursor
    oText = oDoc.Text

    ' Move the ViewCursor to the start of the document
    oViewCursor.gotoStart(False)

    ' Create a TextCursor at the ViewCursor position
    oTextCursor = oText.createTextCursorByRange(oViewCursor.getStart())

    ' Insert text at the TextCursor position
    oText.insertString(oTextCursor, "Inserted text at the start.", False)

    ' The ViewCursor remains at the start
End Sub
```

In this example, the ViewCursor is moved to the start of the document, and the TextCursor is used to insert text at that position. The ViewCursor does not move from the start position, demonstrating how the two cursors can operate independently.

## Text cursors

As far as data is concerned, view cursors have minimum interaction in comparison to Text cursors. The latter can only move within specific text ranges, but is not restricted to viewable content like the former. Multiple text cursors are allowed for one document. XTextCursor interface is implemented from a set of interfaces including[1]:

Interface	Description
com.sun.star.text.XTextCursor	The primary text cursor that defines simple movement methods
com.sun.star.text.XWordCursor	Provides word-related movement and testing methods
com.sun.star.text.XSentenceCursor	Provides sentence-related movement and testing methods
com.sun.star.text.XParagraphCursor	Provides paragraph-related movement and testing methods
com.sun.star.text.XTextViewCursor	Derived from XTextCursor, this describes a cursor in a text document’s view
Getting a text cursor requires the acquisition of ViewCursor

oCursor = ViewCursor.getText().createTextCursorByRange(ViewCursor)

However, page styles provide specific text ranges for Header and Footer locations, a feature used in Page Numbering Addon:

#Header text
 Num_Position = NumberedPage.HeaderText

#Footer text
Num_Position = NumberedPage.FooterText
After getting the required text range, a text cursor for data interaction is acquired: NumCursor = Num_Position.Text.createTextCursor()

Inserting content
Num_Position object implements the XText interface, allowing for text content insertion/deletion using specific methods. As far as Page Numbering Addon is concerned, these methods are used as following:

Basic
‘ NumberingDecorationComboBox is a dialog element that provides a decoration string value.
‘ PageNumber is a specific field that represents page numbering LO struct. 

    Select Case NumberingDecorationComboBox.Text
        Case "#"
            Num_Position.insertTextContent(NumCursor, PageNumber, False)
        Case "-#-"
            Num_Position.insertString(NumCursor, "-", False)
            Num_Position.insertTextContent(NumCursor, PageNumber, False)
            Num_Position.insertString(NumCursor, "-", False)
        Case "[#]"
            Num_Position.insertString(NumCursor, "[", False)
            Num_Position.insertTextContent(NumCursor, PageNumber, False)
            Num_Position.insertString(NumCursor, "]", False)
        Case "(#)"
            Num_Position.insertString(NumCursor, "(", False)
            Num_Position.insertTextContent(NumCursor, PageNumber, False)
            Num_Position.insertString(NumCursor, ")", False)
        Case Else
            Print "Custom decoration unimplemented feature"
    End Select
Python
# NumberingDecorationComboBoxText object is a dialog control that provides decoration option
NumberingDecorationComboBoxText = oDialog1Model.getByName(
        "NumberingDecoration").Text

    if NumberingDecorationComboBoxText == "#":
        Num_Position.insertTextContent(NumCursor, PageNumber, False)
    elif NumberingDecorationComboBoxText == "-#-":
        Num_Position.insertString(NumCursor, "-", False)
        Num_Position.insertTextContent(NumCursor, PageNumber, False)
        Num_Position.insertString(NumCursor, "-", False)
    elif NumberingDecorationComboBoxText == "[#]":
        Num_Position.insertString(NumCursor, "[", False)
        Num_Position.insertTextContent(NumCursor, PageNumber, False)
        Num_Position.insertString(NumCursor, "]", False)
    elif NumberingDecorationComboBoxText == "(#)":
        Num_Position.insertString(NumCursor, "(", False)
        Num_Position.insertTextContent(NumCursor, PageNumber, False)
        Num_Position.insertString(NumCursor, ")", False)
    else:
        raise Exception("Custom decoration unimplemented feature")
More on Cursors on Andrew Pitonyak ’s OpenOffice.org Macros Explained book

[1] Andrew Pitonyak ’s OpenOffice.org Macros Explained book

## scripts

Les fichiers de macros sont détéctés par libreOffice lorsqu'ils sont dans le dossier adéquat qu'il faut créé:

~/.config/libreoffice/4/user/Scripts/python

## intégrer

Il est aussi possible d'integrer une macro à un fichier LibreOffice: http://christopher5106.github.io/office/2015/12/06/openoffice-libreoffice-automate-your-office-tasks-with-python-macros.html#pack-your-script-inside-the-document--the-opendocument-format

Pour cela on créé un script Python include_macro.py qui va ajouter myscript.py dans le document document.ods

#include_macro.py
import zipfile
import shutil
import os
import sys

print("Delete and create directory with_macro")
shutil.rmtree("with_macro",True)
os.mkdir("with_macro")

filename = "with_macro/"+sys.argv[1]
print("Open file " + sys.argv[1])
macroName = sys.argv[2]
shutil.copyfile(sys.argv[1],filename)

doc = zipfile.ZipFile(filename,'a')
doc.write(macroName, "Scripts/python/"+macroName)
manifest = []
for line in doc.open('META-INF/manifest.xml'):
  if '</manifest:manifest>' in line.decode('utf-8'):
    for path in ['Scripts/','Scripts/python/','Scripts/python/'+macroName]:
      manifest.append(' <manifest:file-entry manifest:media-type="application/binary" manifest:full-path="%s"/>' % path)
  manifest.append(line.decode('utf-8'))

doc.writestr('META-INF/manifest.xml', ''.join(manifest))
doc.close()
print("File created: "+filename)
Éxécuter l'integration:

python include_macro.py document.ods myscript.py

## Outils

Extensions Python
APSO (Alternative Script Organizer for Python )
Cette extension permet de gérer dans LibreOffice les macros Python (celles dans le dossier adéquat). Il faut avoir installé libreoffice-script-provider-python . Elle permet aussi d'éxécuter du code Python à la volé avec la console Python disponible. Elle permet aussi aussi de débuger une macro avec une librairie dédiée. Elle permet aussi aussi aussi d'intégrer une macro au document ouvert pour remplacer le script include_macro.py.

https://gitlab.com/jmzambon/apso

Include Python Path Extension for LibreOffice
Avec Python, il est possible de créer des environnement virtuel. Cela permet d'avoir des versions différentes de Python en paralléle et surtout des packages et des librairies spécifique à un environnement.

Cette extension permet d'ajouter un chemin à un environnement Python spécifique pour utiliser les packages de cet environnement par LibreOffice.

https://github.com/Amourspirit/libreoffice-python-path-ext/tree/main#readme

OOO Development Tools (OooDev)
Cette extension permet à des scripts Python d'être exécuté en dehors d'un document LibreOffice et d'interagir avec celui-ci. Il existe aussi l'extension [oooenv](https://pypi.org/project/oooenv/) qui est nécessaire pour créer un envirronement virtuel afin d'utiliser OooDevTools. Plus d'info ici.

MRI
MRI n'est pas une extension spécifique à Python mais elle permet d'inspecter les éléments d'un document LibreOffice. La version 1.1.4 ne s'installe pas avec LibreOffice 7.6.4 mais la version 1.3.4 fonctionne et est disponible sur Github

Librairies Python pour LibreOffice
Afin d'interagir avec un document LibreOffice, il est nécessaire d'utiliser l'API.

Des bibliothèques pour intéragir avec l'API existe:

[OooDevTool](https://python-ooo-dev-tools.readthedocs.io/en/latest/), une librairie compléte
[ScriptForge](https://gitlab.com/LibreOfficiant/scriptforge), un framework plus qu'une librairie
[types-unopys](https://github.com/Amourspirit/python-types-unopy), pour ajouter le typage
D'autres librairies peuvent être ajoutées à un code Python tel que Numpy ou Numexpr

## macros

There are two places where LibreOffice picks up the macro files (i.e. .py files). One is for the system location, which is common for all users. And the other one is user-specific location, and files are accessible to that user only.

For Linux	Path
User-specific path	~/.config/libreoffice/4/user/Scripts/python
All users	/usr/lib64/libreoffice/share/Scripts/python
For Windows	Path
User-specific path	C:\Users\<user>\AppData\Roaming\LibreOffice\4\user\Scripts\python
All users	%APPDATA%\LibreOffice\4\user\Scripts\python
For macOS	Path
User-specific path	/Users/<YourUsername>/Library/Application\ Support/LibreOffice/4/user/Scripts/python
All users	—

## writing a macro doc

In LibreOffice, everything you do, e.g. type, colour, insert, is “watched” by a controller. The controller then dispatches the changes to the document frame, i.e. the main window area of the Calc. So the document variable refers to the main area of Calc.

The createUnoService creates an instance of the DispatchHelper service. This service will help us to dispatch the tasks from the macro to the frame. Almost all LibreOffice macro tasks can be executed using the dispatcher.

Now we will declare an array of properties. Properties are always in a name/value pair. Thus the name contains the property name, and the value contains the value of that property.

dim args1(0) as new com.sun.star.beans.PropertyValue