def prepareDocument(): 
  model = getModel() 
  search = model.createSearchDescriptor() 
  # dev.printObjectProperties(search) # explore the object 
 
  search.setPropertyValue('SearchRegularExpression', True); 
  
  # remove all paragraph padding 
  search.setSearchString('^\s*(.+?)\s*$') # start and trailing whitespace
  search.setReplaceString('$1') 
  replaced = model.replaceAll(search) 
  
  # remove all empty paragraphs 
  search.setSearchString('^$') # empty paragraphs 
  search.setReplaceString('') 
  replaced = model.replaceAll(search) 