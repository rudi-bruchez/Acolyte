# Short documentation for Python LibreOffice extensions

## Sources

+ [LibreOffice Python API](https://api.libreoffice.org/docs/idl/ref/index.html)
+ [python-ooo-dev-tools](https://python-ooo-dev-tools.readthedocs.io/))

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