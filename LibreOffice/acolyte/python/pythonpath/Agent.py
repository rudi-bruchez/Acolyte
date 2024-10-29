# -*- coding: utf-8 -*-

import os, sys, uuid, sqlite3, json
import uno, unohelper

from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_core.output_parsers import StrOutputParser
from langgraph.checkpoint.sqlite import SqliteSaver
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

sys.stderr = sys.stdout
DOC_ID_NAME = "Acolyte Document Id"

class Agent:
    def __init__(self, doc):

        self.doc = doc
        self.doc_id = self.DocumentGuid

        # Initialize sqlite to persist state between graph runs
        self.__memory_write_config = {"configurable": {"thread_id": f"{self.doc_id}", "checkpoint_ns": ""}}
        self.__memory_read_config =  {"configurable": {"thread_id": f"{self.doc_id}"}}
        self.checkpointer = SqliteSaver.from_conn_string(self.DatabaseName)

        # Define a new graph
        self.workflow = StateGraph(MessagesState)
        self.workflow.add_node("agent", self.__call_model)

        # Set the entrypoint as `agent`
        # This means that this node is the first one called
        self.workflow.add_edge(START, "agent")

        # Finally, we compile it!
        # This compiles it into a LangChain Runnable,
        # meaning you can use it as you would any other runnable.
        # Note that we're (optionally) passing the memory when compiling the graph
        self.langchain_app = self.workflow.compile(checkpointer=self.checkpointer)

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
        os.path.join(Session.UserProfile, "Acolyte/" )

    @property    
    def DatabaseName(self):
        os.path.join(self.AcolyteFolder, "Acolyte.sqlite" )

    def GetHistory(self):
        self.checkpointer.get(self.__memory_read_config)

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

    def SummarizeChaptersMapReduce(self, chapters: list):
        llm = ChatOpenAI(temperature=0)

        # Map
        map_template = """The following is a set of documents
        {chapters}
        Based on this list of docs, please identify the main themes 
        Helpful Answer:"""
        map_prompt = PromptTemplate.from_template(map_template)
        map_chain = LLMChain(llm=llm, prompt=map_prompt)

        # Reduce
        reduce_template = """The following is set of summaries:
        {chapters}
        Take these and distill it into a final, consolidated summary of the main themes. 
        Helpful Answer:"""
        reduce_prompt = PromptTemplate.from_template(reduce_template)

        # Run chain
        reduce_chain = LLMChain(llm=llm, prompt=reduce_prompt)

        # Takes a list of documents, combines them into a single string, and passes this to an LLMChain
        combine_documents_chain = StuffDocumentsChain(
            llm_chain=reduce_chain, document_variable_name="chapters"
        )

        # Combines and iteratively reduces the mapped documents
        reduce_documents_chain = ReduceDocumentsChain(
            # This is final chain that is called.
            combine_documents_chain=combine_documents_chain,
            # If documents exceed context for `StuffDocumentsChain`
            collapse_documents_chain=combine_documents_chain,
            # The maximum number of tokens to group documents into.
            token_max=4000,
        )

        # Combining documents by mapping a chain over them, then combining results
        map_reduce_chain = MapReduceDocumentsChain(
            # Map chain
            llm_chain=map_chain,
            # Reduce chain
            reduce_documents_chain=reduce_documents_chain,
            # The variable name in the llm_chain to put the documents in
            document_variable_name="chapters",
            # Return the results of the map steps in the output
            return_intermediate_steps=False,
        )

        text_splitter = CharacterTextSplitter.from_tiktoken_encoder(
            chunk_size=1000, chunk_overlap=0
        )
        split_docs = text_splitter.split_documents(chapters)
        map_reduce_chain.run(split_docs)

    def SummarizeChaptersRefine(self, chapters: list):
        llm = ChatOpenAI(temperature=0)

        prompt_template = """Write a concise summary of the following:
        {text}
        CONCISE SUMMARY:"""
        prompt = PromptTemplate.from_template(prompt_template)

        refine_template = (
            "Your job is to produce a final summary\n"
            "We have provided an existing summary up to a certain point: {existing_answer}\n"
            "We have the opportunity to refine the existing summary"
            "(only if needed) with some more context below.\n"
            "------------\n"
            "{text}\n"
            "------------\n"
            "Given the new context, refine the original summary in Italian"
            "If the context isn't useful, return the original summary."
        )
        refine_prompt = PromptTemplate.from_template(refine_template)
        chain = load_summarize_chain(
            llm=llm,
            chain_type="refine",
            question_prompt=prompt,
            refine_prompt=refine_prompt,
            return_intermediate_steps=True,
            input_key="input_documents",
            output_key="output_text",
        )

        text_splitter = CharacterTextSplitter.from_tiktoken_encoder(
            chunk_size=1000, chunk_overlap=0
        )
        split_docs = text_splitter.split_documents(chapters)
        
        result = chain({"input_documents": split_docs}, return_only_outputs=True)

    def SummarizeChapters(self, chapters: list):
        text_splitter = CharacterTextSplitter.from_tiktoken_encoder(
            chunk_size=1000, chunk_overlap=0
        )

        summarize_document_chain = AnalyzeDocumentChain(
            combine_docs_chain=chapters, text_splitter=text_splitter
        )
        summary = summarize_document_chain.invoke(chapters)
        self.__SaveSummary(json.dumps(summary))

    def __SaveSummary(self, summary: str):
        create = """CREATE TABLE IF NOT EXISTS document_summary (
                    doc_id varchar(50) NOT NULL,
                    summary TEXT NOT NULL,
                    when TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    CONSTRAINT pk_summary PRIMARY KEY (doc_id)
                );
                """

        upsert = f"""INSERT INTO document_summary (doc_id, summary)
                    VALUES('{self.doc_id}','{summary}')
                    ON CONFLICT(doc_id) DO UPDATE SET
                        summary = excluded.summary,
                        when = CURRENT_TIMESTAMP;
                """

        with sqlite3.connect(self.DatabaseName) as cn:
            cur = cn.cursor()
            cur.execute(create)
            cur.execute(upsert)
            cur.commit()
            cn.close()

    def LoadSummary(self) -> str:
        select = f"""SELECT summary FROM document_summary 
                     WHERE doc_id = '{self.doc_id}';
                """
        try:
            with sqlite3.connect(self.DatabaseName) as cn:
                cur = cn.cursor()
                cur.execute(select)
                row = cur.fetchone()
                cn.close()
                return row["summary"]
        except sqlite3.OperationalError as e:
            print(e)
            return ""