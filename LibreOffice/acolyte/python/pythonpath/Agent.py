# -*- coding: utf-8 -*-

import os, sys, uuid, json
# import sqlite3
import uno, unohelper
import yaml

from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_core.output_parsers import StrOutputParser
# from langgraph.checkpoint.sqlite import SqliteSaver
from langgraph.graph import END, START, StateGraph, MessagesState
from langgraph.prebuilt import ToolNode
from langchain.chains.summarize import load_summarize_chain
from langchain.chains import MapReduceDocumentsChain, ReduceDocumentsChain
from langchain_text_splitters import CharacterTextSplitter
from langchain.chains.combine_documents.stuff import StuffDocumentsChain
from langchain.chains import AnalyzeDocumentChain
from langchain.chains.llm import LLMChain
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_core.messages import HumanMessage
from pathlib import Path
from LibreOffice import LO
from dotenv import load_dotenv
from enum import Enum

sys.stderr = sys.stdout
DOC_ID_NAME = "Acolyte Document Id"
ContextType = Enum('ContextType', ['Full', 'Chapter', 'Paragraphs', 'Digest', 'TOC', 'Nope'])

class LangchainAgent:
    def __init__(self, doc):

        self.doc = doc
        self.doc_id = self.DocumentGuid

        # Initialize sqlite to persist state between graph runs
        self.__memory_write_config = {"configurable": {"thread_id": f"{self.doc_id}", "checkpoint_ns": ""}}
        self.__memory_read_config =  {"configurable": {"thread_id": f"{self.doc_id}"}}
        # self.checkpointer = SqliteSaver.from_conn_string(self.DatabaseName)
        # self.checkpointer = None

        # # Define a new graph
        # self.workflow = StateGraph(MessagesState)
        # self.workflow.add_node("agent", self.__call_model)

        # # Set the entrypoint as `agent`
        # # This means that this node is the first one called
        # self.workflow.add_edge(START, "agent")

        # # Finally, we compile it!
        # # This compiles it into a LangChain Runnable,
        # # meaning you can use it as you would any other runnable.
        # # Note that we're (optionally) passing the memory when compiling the graph
        # self.langchain_app = self.workflow.compile(checkpointer=self.checkpointer)

        self.__lo = LO(self.doc)
        self.__setPromptVariables()
        self.__loadApiKeys()
        self.client = ChatOpenAI(model="gpt-4o-mini")

    def __loadApiKeys(self):
        # Load environment variables from a .env file located in a specific directory
        # env_path = Path(self.AcolyteFolder, '.env')
        folder = self.__lo.AcolyteUserFolder
        # folder = r"C:\Users\rudi\AppData\Roaming\LibreOffice\4\user"
        # env_path = os.path.join(self.AcolyteFolder, '.env')
        env_path = os.path.join(folder, '.env')
        load_dotenv(dotenv_path=env_path)

        api_key = os.getenv('OPENAI_API_KEY')
        if not api_key:
            raise ValueError("API key not found in the environment file")

        os.environ["OPENAI_API_KEY"] = api_key

    def __call_model(self, state: MessagesState):
        messages = state['messages']
        response = self.client.invoke(messages)
        # We return a list, because this will get added to the existing list
        return {"messages": [response]}

    def __loadPrompt(self, prompt_file: str):
        file = os.path.join(self.__lo.AcolyteUserFolder, "prompt_techbook.yaml")
        with open(file, 'r', encoding="utf-8") as stream:
            prompt = yaml.safe_load(stream)
            return prompt

    def __setPromptVariables(self):
        self.PromptVariables = {}

        prop = self.doc.DocumentProperties

        # self.PromptVariables['language'] = prop.CharLocale
        self.PromptVariables['title'] = prop.Title

        self.PromptVariables['subject'] = prop.Subject

        # @property
        # def Characters(self) -> int:
        #     return self.doc.CharacterCount

        self.PromptVariables['author'] = prop.Author
        self.PromptVariables['contributor'] = prop.Contributor 
        self.PromptVariables['coverage'] = prop.Coverage
        self.PromptVariables['description'] = prop.Description

        # @property
        # def DocumentStatistics[PageCount](self) -> str:

        self.PromptVariables['generator'] = prop.Generator
        
        # @property
        # def Identifier(self) -> str:
        self.PromptVariables['keywords'] = prop.Keywords
        self.PromptVariables['language'] = Language.Language
        self.PromptVariables['publisher'] = prop.Publisher
        self.PromptVariables['source'] = prop.Source
        self.PromptVariables['type'] = prop.Type
        
        # @property
        # def doc.Location (url)(self) -> str:
        # @property
        # def doc.Title = nom du fichier(self) -> str:

    def CreatePrompt(self, prompt_file: str, context_type: ContextType = ContextType.Full) -> str:
        prompt_info = self.__loadPrompt(prompt_file)

        template = ChatPromptTemplate([
            ("system", prompt_info.get("system_prompt")),
            ("user", prompt_info.get("user_prompt")),
            # ("system", "Voici le contenu du chapitre en cours du livre {title}"),
            # MessagesPlaceholder("paragraphs"),
        ])

        if context_type == ContextType.Full:
            template.append([
            ("system", "Voici le contenu du livre {title} en cours"),
            MessagesPlaceholder("document"),
            ])
        elif context_type == ContextType.Chapter:
            template.append([
            ("system", "Voici le contenu du chapitre en cours du livre {title}"),
            MessagesPlaceholder("chapter"),
            ])
        elif context_type == ContextType.Paragraphs:
            template.append([
            ("system", "Voici le contenu des quelques paragraphes précédents du livre {title}"),
            MessagesPlaceholder("paragraphs"),
            ])
        elif context_type == ContextType.Digest:
            template.append([
            ("system", "Voici le résumé du livre {title}"),
            MessagesPlaceholder("summary"),
            ])
        elif context_type == ContextType.TOC:
            template.append([
            ("system", "Voici la table des matières du livre {title}"),
            MessagesPlaceholder("toc"),
            ])
        # elif context_type == ContextType.Nope:
        #     template.append([
        #     ("system", "Aucune information contextuelle fournie pour le livre {title}"),
        #     ])

        formatted_prompt = template.format(self.PromptVariables)        
        # formatted_prompt = template.format(subject="SQL Server", title="Réussir SQL Server")        
        # prompt_template.invoke({"subject": "SQL Server", "title": "Réussir SQL Server"})        
        return formatted_prompt

    @property
    def DocumentGuid(self) -> str:
        udp = self.doc.DocumentProperties.UserDefinedProperties

        if not udp.getPropertySetInfo().hasPropertyByName(DOC_ID_NAME):
            udp.addProperty(DOC_ID_NAME ,128,"")
            udp.setPropertyValue(DOC_ID_NAME, str(uuid.uuid4()))

        return udp.getPropertyValue(DOC_ID_NAME)

    @property
    def AcolyteFolder(self):
        return os.path.join(Session.UserProfile, "Acolyte/" )

    @property
    def DatabaseName(self):
        return os.path.join(self.AcolyteFolder, "Acolyte.sqlite" )

    def GetHistory(self):
        return self.checkpointer.get(self.__memory_read_config)

    @property
    def Checkpoints(self) -> list:
        return list(self.checkpointer.list(self.__memory_read_config))

    def StoreCheckpoint(self, checkpoint):
        self.checkpointer.put(self.__memory_write_config, checkpoint , {}, {})

    def GetSimpleResponse(self, user_prompt: str, system_prompt: str = None) -> str:        
        try:
            messages = [
                # SystemMessage(content="Translate the following from English into Italian"),
                HumanMessage(content = user_prompt),
            ]
            parser = StrOutputParser()
            chain = self.client | parser
            return chain.invoke(messages)
        except Exception as e:
            return "Erreur " + str(e)

    def SummarizeChapters(self, chapters: list):
        text_splitter = CharacterTextSplitter.from_tiktoken_encoder(
            chunk_size=1000, chunk_overlap=0
        )

        summarize_document_chain = AnalyzeDocumentChain(
            combine_docs_chain=chapters, text_splitter=text_splitter
        )
        summary = summarize_document_chain.invoke(chapters)
        self.__SaveSummary(json.dumps(summary))

    # def __SaveSummary(self, summary: str):
    #     create = """CREATE TABLE IF NOT EXISTS document_summary (
    #                 doc_id varchar(50) NOT NULL,
    #                 summary TEXT NOT NULL,
    #                 when TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    #                 CONSTRAINT pk_summary PRIMARY KEY (doc_id)
    #             );
    #             """

    #     upsert = f"""INSERT INTO document_summary (doc_id, summary)
    #                 VALUES('{self.doc_id}','{summary}')
    #                 ON CONFLICT(doc_id) DO UPDATE SET
    #                     summary = excluded.summary,
    #                     when = CURRENT_TIMESTAMP;
    #             """

    #     with sqlite3.connect(self.DatabaseName) as cn:
    #         cur = cn.cursor()
    #         cur.execute(create)
    #         cur.execute(upsert)
    #         cur.commit()
    #         cn.close()

    # def LoadSummary(self) -> str:
    #     select = f"""SELECT summary FROM document_summary
    #                  WHERE doc_id = '{self.doc_id}';
    #             """
    #     try:
    #         with sqlite3.connect(self.DatabaseName) as cn:
    #             cur = cn.cursor()
    #             cur.execute(select)
    #             row = cur.fetchone()
    #             cn.close()
    #             return row["summary"]
    #     except sqlite3.OperationalError as e:
    #         print(e)
    #         return ""