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
  + `./python.exe -m pip install langchain langchain-openai langgraph`

## docs to read

+ [Configuration Extension](https://eellak.github.io/gsoc2018-librecust/docs/Configuration-extension/)
+ [OO dev tools](https://github.com/Amourspirit/python_ooo_dev_tools)
+ [MRI](https://github.com/hanya/MRI/releases)
+ [Python Design Guide](https://wiki.documentfoundation.org/Macros/Python_Design_Guide#Developing)
+ [**AGENTS !**](https://lilianweng.github.io/posts/2023-06-23-agent/)

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

## ideas

Here are some suggestions for AI tools for Writer:

I’m currently finishing up an ~500 page book, and I’m still getting ideas I want to include, but finding the appropriate place to add them can be challenging. Ordinary search requires exact references, which I often can’t come up with. An AI search, that would allow me to reference a concept, would be a boon. I could say, “find where I talk about people who hate Italian food.” It would find the passage where I used the phrase “…their main objection to the boot shaped land of their ancestors, is the food…”, which would be difficult to find with standard search if I couldn’t remember that particular word play.

A more scholastic Spelling/Grammar/Punctuation/Sentence Structure/etc. check. I currently copy/paste segments into Google Docs for it’s more advanced corrective capability. And, now Google Docs has a ChatGPT extension–though the performance of the AI extension has been uneven, and even false–but still helpful. And even so, Google Docs, sans extension, does a great job with basic spelling/grammar/punctuation–and always far better than Writer in it’s current form.

A Spell Check that understands colloquial idioms, and can associate them to the written context. Also, a Spell Check that has a more real world reference, so it recognizes names/places/events/etc not only from history, but also from current culture, that often are not in the LibreOffice dictionary.

Google Docs is fairly remarkable at this, as well.

A tool that can read what I’ve written, so far, and assess it’s structure, and assist in the daunting task of organizing the content into a cohesive work. It would, also, be able to locate areas of duplication, places where the thought isn’t complete or could be fleshed out, places where the writing could be tightened up, etc. It would be able to summarize, and offer suggestions for reorganizing it for better flow, making it more readable, and/or more engaging, and/or more persuasive, and/or more convincing, etc. An AI tool that could work on a large document, like my 500+ page book would be awesome! ChatGPT tends to be limited in this respect. A possible way to cut down on AI processing cost might be gained by giving the AI the task of following a particular document throughout it’s lifetime, thus it could be more incremental in it’s scans and subsequent assessments, only scanning changes, rather than scanning the whole document each time a change is detected.