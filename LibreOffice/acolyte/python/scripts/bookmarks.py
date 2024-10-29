import re

def set_bookmarks():
    doc = XSCRIPTCONTEXT.getDocument()
    print(f"{doc.Title = }")
    # search patterns 
    searchBookmarks = doc.createReplaceDescriptor() ##!!
    searchBookmarks.SearchString = "\{bookmark:\s*(.*?)\}"
    searchBookmarks.SearchRegularExpression = True
    searchBookmarks.SearchAll = True
    searchBookmarks.SearchCaseSensitive = True
    searchBookmarks.SearchWords = False
    searchBookmarks.setReplaceString("$1")
    
    bookmarks = doc.findAll(searchBookmarks)
    for pattern in bookmarks:
        name = re.sub(r'{bookmark:\s*?(.*)}', r'\1', pattern.String)
        bookmark = doc.createInstance("com.sun.star.text.Bookmark")
        bookmark.setName(name)
        doc.Text.insertTextContent(pattern, bookmark, True) 
    #and last:
    doc.replaceAll(searchBookmarks)