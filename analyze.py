# imports

# The basics
import os
import sys
import logging
import argparse

# Utilities
from pathlib import Path
from datetime import datetime

from langchain_ollama import OllamaLLM
from langchain_openai import ChatOpenAI
from typing_extensions import TypedDict, List
from dotenv import load_dotenv

# LangChain foundation
from langchain_core.output_parsers import BaseOutputParser
from langchain_core.prompts import PromptTemplate, ChatPromptTemplate, SystemMessagePromptTemplate, \
    HumanMessagePromptTemplate

# LangChain utilities
from langchain.chains.combine_documents import create_stuff_documents_chain
from langchain_community.document_loaders import DirectoryLoader, TextLoader
from langchain.retrievers.multi_query import MultiQueryRetriever
from langchain_text_splitters import RecursiveCharacterTextSplitter

# Vector database
from langchain_chroma import Chroma

# Ollama
from langchain_community.embeddings import OllamaEmbeddings

# OpenAI
from langchain_community.embeddings import OpenAIEmbeddings

# Oracle and OCI
import oci
from langchain_community.chat_models import ChatOCIGenAI
from langchain_community.embeddings import OCIGenAIEmbeddings

# UI
import gradio as gr

# ----------------------------------------------------------------------------------------------------------------------
# Constants
llm_embeddings_model = {
    "ollama": "nomic-embed-text",
    "openai": "text-embedding-ada-002",
    "oci": "cohere.embed-english-v3.0"
}

llm_chat_model = {
    "ollama": "phi3:3.8b",
    "openai": "gpt-4o-mini",
    "oci": "cohere.command-r-08-2024"
}

db_name = "vector_db"

oci_genai_service_endpoint = "https://inference.generativeai.us-chicago-1.oci.oraclecloud.com"

# ----------------------------------------------------------------------------------------------------------------------
# Parse parameters

def parse_args():
    parser = argparse.ArgumentParser(description="Script to read in exported OneNote files in Markdown and make them available to query via an LLM")

    parser.add_argument("llm", type=str, help="The LLM to use (openai, ollama, oci)")
    parser.add_argument("--docdir", type=str, help="The directory to search for markdown documents")
    parser.add_argument("--skipembedding", action="store_true", help="Skip embedding the documents")
    parser.add_argument("--configfile", type=str, help="Location of LLM's configuration file")
    parser.add_argument("--compartment", type=str, help="The OCI compartment to use")
    parser.add_argument("-v", "--verbose", action="count", default=0, help="Logging level")

    args = parser.parse_args()

    if args.verbose == 4:
        logging.basicConfig(level=logging.DEBUG)
    elif args.verbose == 3:
        logging.basicConfig(level=logging.INFO)
    elif args.verbose == 2:
        logging.basicConfig(level=logging.WARNING)
    elif args.verbose == 1:
        logging.basicConfig(level=logging.ERROR)
    else:
        logging.basicConfig(level=logging.CRITICAL)

    return {
        "llm": args.llm,
        "docdir": args.docdir,
        "skipembedding": args.skipembedding,
        "configfile": args.configfile,
        "compartment": args.compartment
    }

# ----------------------------------------------------------------------------------------------------------------------
def validate_llm(llm):
    if not(llm in llm_embeddings_model):
        logging.error("Invalid LLM name. Expected one of the of the following: {}".format(list(llm_embeddings_model.keys())))
        sys.exit(1)

# ----------------------------------------------------------------------------------------------------------------------
def validate_configfile(configfile, llm):
    if llm == "oci":
        # If we're running the LLM on OCI, then validate the configuration file.
        if configfile:
            if (not((configfile) and os.access(configfile, os.R_OK))):
                logging.error("Invalid config file. Either it does not exist or it cannot be read. File: \"{}\"".format(configfile))
                sys.exit(1)
            else:
                logging.info("Reading LLM from \"{}\"".format(configfile))
                os.environ["OCI_CONFIG_FILE"] = configfile
                logging.info("OCI_CONFIG_FILE environment variable set to \"{}\"".format(os.environ["OCI_CONFIG_FILE"]))
        else:
            logging.error( "ASSERT: OCI was specified but no configuration file was specified. A configuration file is required.")
            sys.exit(1)

# ----------------------------------------------------------------------------------------------------------------------
def setup_llm(llm, configfile, compartment_name):
    oci_compartment = None

    if llm == "ollama":
        pass

    elif llm == "openai":
        load_dotenv()

        if ( not (os.environ.get("OPENAI_API_KEY"))):
            logging.error("OpenAI_API_KEY not set in .env file.")
            sys.exit(1)

        logging.info("OpenAI key: \"{}...{}\".".format(os.environ.get("OPENAI_API_KEY")[:6], os.environ.get("OPENAI_API_KEY")[-6:]))

    elif llm == "oci":
        """
        Because the LangChain libraries call oci.config.from_file with no ability to specify a file, the configuration
        file must be specified in an environment variable, OCI_CONFIG_FILE. 
        """
        if (not (os.environ.get("OCI_CONFIG_FILE"))):
            logging.error("OCI_CONFIG_FILE environment variable not set. It must be set when using OCI.")
            sys.exit(1)
        oci_config_file = Path(configfile)

        try:
            full_path_oci_config_file = oci_config_file.resolve()
            logging.info("Configuration file located at: \"{}\".".format(full_path_oci_config_file))

            if (not((full_path_oci_config_file) and os.access(full_path_oci_config_file, os.R_OK))):
                logging.error("OCI configuration file either does not exist or is not readable")
                sys.exit(1)
        except Exception as e:
            logging.error("OCI configuration file either does not exist or is not readable")
            sys.exit(1)

        try:
            config = oci.config.from_file()
        except Exception as e:
            logging.error("OCI configuration file could not be validated.")
            sys.exit(1)

        try:
            signer = oci.signer.Signer(
                tenancy=config["tenancy"],
                user=config["user"],
                fingerprint=config["fingerprint"],
                private_key_file_location=config.get("key_file"),
                # pass_phrase is optional and can be None
                pass_phrase=oci.config.get_config_value_or_default(config, "pass_phrase"),
                # private_key_content is optional and can be None
                private_key_content=config.get("key_content")
            )
        except Exception as e:
            logging.error("A request could not be authenticated to OCI.")
            sys.exit(1)

        try:
            identity_client = oci.identity.IdentityClient(config, signer=signer )
        except Exception as e:
            logging.error("A request to OCI IAM could not be authenticated.")
            sys.exit(1)

        try:
            tenancy_data = identity_client.get_tenancy(config["tenancy"]).data
        except Exception as e:
            logging.error("Tenancy information could not be retrieved.")
            sys.exit(1)

        if (tenancy_data is None):
            logging.error("No tenancy information was retrieved from OCI.")
            sys.exit(1)

        logging.info("Successfully connected to OCI.")
        tenancy_compartment = tenancy_data.id
        logging.info("Tenancy compartment: \"{}\".".format(tenancy_compartment))

        if compartment_name:
            if compartment_name != "tenancy":
                list_compartments_response = identity_client.list_compartments(
                    compartment_id = tenancy_compartment
                )

                if list_compartments_response.data is None:
                    logging.error("No compartments retrieved from OCI.")
                    sys.exit(1)

                # This is a list of oci.identity.models.compartment.Compartment
                compartment_list = list_compartments_response.data

                for compartment in compartment_list:
                    if compartment.name.lower() == compartment_name.lower():
                        oci_compartment = compartment.id

                if oci_compartment is None:
                    logging.error("No compartment was found to match \"{}\"".format(compartment_name))
                    sys.exit(1)

                logging.info("Compartment: \"{}\" = \"{}\".".format(compartment_name, oci_compartment))

            else:
                oci_compartment = tenancy_compartment
        else:
            oci_compartment = tenancy_compartment

    else:
        logging.error("ASSERT: An invalid LLM model was passed.[{}]".format(llm))
        sys.exit(1)

    return {
        "oci_compartment": oci_compartment
    }

# ----------------------------------------------------------------------------------------------------------------------
def read_documents(docdir, skipembedding):
    text_loader_kwargs = {'encoding': 'utf-8'}

    if skipembedding:
        pass
    else:
        if docdir:
            full_path_docdir = Path(docdir)
            if (not((full_path_docdir) and os.access(full_path_docdir, os.R_OK))):
                logging.error("Invalid document directory. Either it does not exist or it cannot be read. Directory: \"{}\"".format(full_path_docdir))
                sys.exit(1)
            else:
                logging.info("Reading Markdown documents from \"{}\"".format(full_path_docdir))
        else:
            logging.error("ASSERT: No document directory parameter passed")
            sys.exit(1)

        try:
            loader = DirectoryLoader(full_path_docdir.name, glob="**/*.md", loader_cls=TextLoader, loader_kwargs=text_loader_kwargs)
            documents = loader.load()
        except Exception as e:
            logging.error("Unable to load documents from \"{}\"".format(full_path_docdir))
            sys.exit(1)

        return documents

# ----------------------------------------------------------------------------------------------------------------------
def enrich_documents(documents, skipembedding):
    """
    In the directory, the following hierarchy is expected:
    + Top directory
      + (intermediate directories, optional)
        + Section directory
          + Page directory
            - Paragraph #1 file
            - Paragraph #2 file
            - Paragraph #n file

    The format of the paragraph file names should be "YYYY-MM-DD". If it is, then this will be added as metadata in the
    form of a date (YYYY-MM-DD) and a paragraph name (Monday, January 1, 2000). If the file name is not in a date
    format, then the date metadata will be set to blank and the paragraph name will just be set to the name of the
    paragraph.

    In addition, the page and section values will be added as metadata.
    """

    """
    TODO:
    What if the 'docdir' parameter points to the Page directory? This means that when walking the directory hierarchy
    from the bottom, the parent of the Page directory is assumed to be the Section directory, which is not guaranteed.
    """
    if skipembedding:
        pass
    else:
        for doc in documents:
            fileName = Path(doc.metadata["source"]).name
            baseFileName = fileName.removesuffix(".md")
            parentDir = Path(doc.metadata["source"]).parent
            pageDir = (Path(parentDir).name).removesuffix(" page")
            sectionDir = (Path(parentDir).parent.name).removesuffix(" section")

            doc.metadata["page"] = pageDir
            doc.metadata["section"] = sectionDir

            try:
                paragraphDate = datetime.strptime(baseFileName, "%Y-%m-%d").strftime("%Y-%m-%d")
            except ValueError:
                paragraphDate = ""

            if (paragraphDate):
                doc.metadata["date"] = paragraphDate
                doc.metadata["paragraph"] = datetime.strptime(baseFileName, "%Y-%m-%d").strftime("%A, %B %d, %Y")
            else:
                doc.metadata["date"] = ""
                doc.metadata["paragraph"] = baseFileName

        return documents

# ----------------------------------------------------------------------------------------------------------------------
def store_vector_db(llm, documents, skipembedding, oci_compartment):
    if skipembedding:
        pass
    else:
        text_splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=200)
        chunks = text_splitter.split_documents(documents)

        logging.info("Total number of chunks: {}".format(len(chunks)))
        logging.debug("Document types found for \"date\": {}".format(len(set(doc.metadata['date'] for doc in documents))))
        logging.debug("Document types found for \"paragraph\": {}".format(len(set(doc.metadata['paragraph'] for doc in documents))))

    if llm == "ollama":
        embeddings = OllamaEmbeddings(model=llm_embeddings_model[llm])

    elif llm == "openai":
        embeddings = OpenAIEmbeddings(model=llm_embeddings_model[llm])

    elif llm == "oci":
        embeddings = OCIGenAIEmbeddings(
            model_id = llm_embeddings_model[llm],
            service_endpoint=oci_genai_service_endpoint,
            compartment_id=oci_compartment
        )
    else:
        logging.error("ASSERT: no llm_model_choice found to choose for embeddings. llm_model_choice: {}".format(llm))
        sys.exit(1)

    # -------------
    # Put the chunks of data into a Vector Store that associates a Vector Embedding with each chunk
    # Chroma is a popular open source Vector Database based on SQLLite

    if skipembedding:
        vector_db = Chroma(persist_directory=db_name, embedding_function=embeddings)
    else:
        # Delete if already exists
        if os.path.exists(db_name):
            Chroma(persist_directory=db_name, embedding_function=embeddings).delete_collection()

        # Create vectorstore
        vector_db = Chroma.from_documents(documents=chunks, embedding=embeddings, persist_directory=db_name)
        vector_db.persist_directory()
        logging.debug("Writing vectordb to \"{}\"".format(db_name))

    retriever = vector_db.as_retriever()

    vector_db_count = vector_db._collection.count()
    sample_embedding = vector_db._collection.get(limit=1, include=["embeddings"])["embeddings"][0]
    dimensions = len(sample_embedding)

    logging.info("Vectordb has {} documents.".format(vector_db_count))
    logging.debug("Sample embedding dimension: {}".format(dimensions))

    return retriever

# ----------------------------------------------------------------------------------------------------------------------
def start_chat(llm, retriever, compartment):
    if llm_model_choice == "ollama":
        llm = OllamaLLM(
            model=llm_chat_model["ollama"],
            temperature=0.8,  # Adjust creativity level
            top_p=0.9         # Adjust nucleus sampling
        )
    elif llm_model_choice == "openai":
        llm = ChatOpenAI(
            model_name=llm_chat_model["openai"],
            temperature=0.8
        )
    elif llm_model_choice == "oci":
        llm = ChatOCIGenAI(
            model_id=llm_chat_model["oci"],
            service_endpoint=oci_genai_service_endpoint,
            compartment_id=oci_compartment,
            model_kwargs={"temperature": 0.8, "max_tokens": 1000},
        )

    # -------------
    # Setup the multiple-question parser for whatever prompt the user gives in order to retrieve as many relevant
    # documents as possible

    class LineListOutputParser(BaseOutputParser[List[str]]):
        #Output parser for a list of lines.

        def parse(self, text: str) -> List[str]:
            lines = text.strip().split("\n")
            return list(filter(None, lines))  # Remove empty lines

    output_parser = LineListOutputParser()

    vector_doc_prompt = PromptTemplate(
        input_variables=["question"],
        template= """You are an AI language model assistant. Your task is to generate five 
        different versions of the given user question to retrieve relevant documents from a vector 
        database. By generating multiple perspectives on the user question, your goal is to help
        the user overcome some of the limitations of the distance-based similarity search. 
        Provide these alternative questions separated by newlines.
        Original question: {question}""",
    )
    llm_chain = vector_doc_prompt | llm | output_parser

    retriever_from_llm = MultiQueryRetriever(
        retriever=retriever, llm_chain=llm_chain
    )

    def chat_with_rag(message, chat_history):ow 
        # Using the vector database, retrieve relevant document chunks
        retrieved_docs = retriever_from_llm.invoke(message)

        system_message_prompt = """
        You are a helpful assistant for question-answering tasks.
        Use the following notes of retrieved context to answer the question. 
        Include links from the retrieved context that point back to the retrieved content.
        If the relevant note has a date, include the date in the answer.
        If you don't know the answer, say that you don't know.
        {context} 
        """

        user_message_prompt = HumanMessagePromptTemplate.from_template(message)
        chain = create_stuff_documents_chain(llm, ChatPromptTemplate.from_messages([system_message_prompt, user_message_prompt]))
        response = chain.invoke({"context":retrieved_docs})
        return response

    view = gr.ChatInterface(fn=chat_with_rag, type="messages", title="OneNote query").launch(inbrowser=True,share=False)

# ======================================================================================================================
# Main

if __name__ == "__main__":

    all_args = parse_args()
    validate_llm(all_args["llm"])
    llm_model_choice = all_args["llm"]

    configfile = all_args["configfile"]
    validate_configfile(configfile, llm_model_choice)

    oci_compartment_name = all_args["compartment"]
    llm_setup_results = setup_llm(llm_model_choice, configfile, oci_compartment_name)

    docdir = all_args["docdir"]
    skipembedding = all_args["skipembedding"]
    documents = read_documents(docdir, skipembedding)
    documents = enrich_documents(documents, skipembedding)

    oci_compartment = llm_setup_results["oci_compartment"]
    retriever = store_vector_db(llm_model_choice, documents, skipembedding, oci_compartment)

    start_chat(llm_model_choice, retriever, oci_compartment)