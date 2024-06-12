from pathlib import Path
from openai import OpenAI
import os
from dotenv import load_dotenv
import openai

from ExcelParser import ExcelParser

load_dotenv()
openai_key = os.getenv("OPENAI_API_KEY")
openai.api_key = openai_key
client = OpenAI()


path_to_file = Path("./tests/example_0.xlsx")
parser = ExcelParser(path_to_file)
parser.parse_all_tables()
parser.display_tables()

tables = parser.extracted_tables

try:
    context = "The following are tables extracted from an Excel file:\n"
    for table in tables:
        context += f"Table: {table}\n"
    
    print(context)

    question = "What is the Total Cash and Cash Equivalent of Nov. 2023?"
    
    response = client.completions.create(
        model="gpt-3.5-turbo",  # Make sure this engine is available to you
        prompt=[
            {"role": "system", "content": "You are analyst in private equity."},
            {"role": "user", "content": context + "\nQuestion: " + question}
        ],
    )
    print(response.choices[0].message)
except Exception as e:
    print(f"An error occurred: {e}")
