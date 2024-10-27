# Acolyte for LibreOffice

```bash
pip install openai
pip install langchain
```

## Add the needed Python modules to LibreOffice

+ go to `cd "C:\Program Files\LibreOffice\program"`
+ copy get-pip.py there
+ run `./python.exe get-pip.py` to install pip
+ install the needed modules with 
  + `./python.exe -m pip install openai`
  + `./python.exe -m pip install langchain`

## Debug

```powershell
C:\Program Files\LibreOffice\program\soffice.exe --accept="socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" --writer
```

<!-- -accept=socket,host=localhost,port=2002;urp; is deprecated.  Use --accept=socket,host=localhost,port=2002;urp; instead. -->

Next, run the installation of python that comes with LibreOffice, which has uno installed by default.

```powershell
"C:\Program Files\LibreOffice\program\python.exe"
>> import uno
```

Now there is a socket for communicating with the opened LibreOffice on port 2002. To communicate with LibreOffice, you need to start your python code with

```python
import uno
localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
```

## print to output

https://mikekaganski.wordpress.com/2018/11/21/proper-console-mode-for-libreoffice-on-windows/

C:\Program Files\LibreOffice\program ❯ ./soffice.com