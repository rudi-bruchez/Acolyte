# ghostwriter

AI writing assistant tools

## plans

+ the tool will be able to connect to AI API and do some RAG if needed
+ it will generate text into destination file, either in markdown, asciidoc, or word format
+ it will manage different writing project types: article, book, blog, linkedin post, video training script, training material.
+ The format of each type will be different, and the tool will be able to generate the format based on the type. The format will be defined in JSON config files.
+ the tool should be able to run like an agent, improving its prompt, and select the best text from several models (?)
+ the tool will iterate : generate an outline in JSON, then apply each step of the outline to generate the full content
+ A JSON prompt config file will indicate :
  + the type of project
  + the type of AI to use
  + the format of the output
  + the purpose of the text
  + the audience
  + the tone
  + the length
  + the language, in a prompt format: "écris ce projet en français", this piece will be injected in each prompt
  + the tool will add system instructions, to retrieve the data in JSON format