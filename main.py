from openai import RateLimitError
import pandas as pd
from langchain_experimental.agents.agent_toolkits import create_pandas_dataframe_agent
from langchain.agents.agent_types import AgentType
from langchain_openai import ChatOpenAI
from langchain.memory import ConversationBufferMemory
from dotenv import load_dotenv
import os

from ExcelParser import ExcelParser

# Simulating session state
session_state = {
    "messages": [],
    "conversation_memory": ConversationBufferMemory(memory_key="chat_history")
}

def display_chat_history():
    """Display the chat history from the session state."""
    for message in session_state["messages"]:
        print(f"{message.get('role')}: {message.get('content')}")

def load_data(file_path):
    """Load data from an Excel file."""
    return pd.read_excel(file_path)

def get_main_menu_input() -> str:
    """Get menu options"""
    print("1. load excel file \n2. New AI Chat. \n3. Exit program. ")
    return input("Select an option: ")

def get_chat_menu_input() -> str:
    """Get main menu chat options"""
    print("1. Back to main menu. \n2. Exit program.")
    return input("Select an option, or enter your chat message: ")

def get_excel_filepath_input() -> str:
    file_path = input("Enter the path to your Excel file (ex: ./tests/example_0.xlsx): ")
    return file_path

def main():
    load_dotenv()
    openai_key = os.getenv("OPENAI_API_KEY")

    display_chat_history()


    print("***** Welcome to Keyann's dope AI PE Analyzer ***** ")

    dfs = []

    is_on = True
    while is_on:

        choice = get_main_menu_input()

        if choice == "1":
            file_path = get_excel_filepath_input()
            try:
                parser = ExcelParser(file_path)
                parser.parse_all_tables()
                dfs = parser.to_dataframes()
                print("Excel file loaded and parsed successfully.")
            except Exception as e:
                print(f"Error reading file: {e}")
                continue

        elif choice == "2":
            if not dfs:
                print("Please load an Excel file first.")
                continue
            

            while True:
                prompt = get_chat_menu_input()

                if prompt == "1":
                    break

                if prompt == "2":
                    exit()

                if prompt:
                    # Add user message to session state
                    session_state["messages"].append({"role": "user", "content": prompt})

                    # Initialize language model
                    llm = ChatOpenAI(
                        temperature=0, model="gpt-3.5-turbo", 
                        openai_api_key=openai_key, 
                        max_tokens=None,
                        timeout=None,
                        max_retries=2
                    )

                    # Create agent
                    agent = create_pandas_dataframe_agent(
                        llm,
                        dfs,
                        verbose=True,
                        number_of_head_rows=10,
                    )

                    try:
                        # Get response from agent
                        response = agent.invoke(prompt)
                    except RateLimitError:
                        print("Rate limit exceeded. Please check your OpenAI API quota.")
                        break

                    # Add assistant message to session state
                    session_state["messages"].append(
                        {"role": "assistant", "content": response}
                    )

                    # Display the conversation
                    display_chat_history()

        elif choice == "3":
            print("Exiting program....")
            print("Goodbye!")
            exit()


if __name__ == "__main__":
    main()
