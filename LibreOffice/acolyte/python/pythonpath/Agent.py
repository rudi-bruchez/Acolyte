# -*- coding: utf-8 -*-

import os, sys, uuid, json
# import sqlite3
import uno, unohelper

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
from langchain_core.prompts import PromptTemplate
from pathlib import Path
from Session import Session
from dotenv import load_dotenv

sys.stderr = sys.stdout
DOC_ID_NAME = "Acolyte Document Id"

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

        self.__loadApiKeys()
        self.client = ChatOpenAI(model="gpt-4o-mini")

    def __loadApiKeys(self):
        # Load environment variables from a .env file located in a specific directory
        # env_path = Path(self.AcolyteFolder, '.env')
        folder = r"C:\Users\rudi\AppData\Roaming\LibreOffice\4\user"
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