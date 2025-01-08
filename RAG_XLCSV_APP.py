# app.py
from dotenv import load_dotenv
import streamlit as st
import pandas as pd
import os
from langchain_openai import ChatOpenAI
from langchain_experimental.agents import create_pandas_dataframe_agent
import xlrd
import openpyxl

# Load environment variables
load_dotenv()
OPENAI_API_KEY = os.getenv("API_KEY")

def load_file(uploaded_file):
    """Load Excel or CSV file and return a dictionary of dataframes for each sheet"""
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    if file_extension == 'csv':
        df = pd.read_csv(uploaded_file)
        return {'Sheet1': df}
    elif file_extension in ['xlsx', 'xls']:
        xls = pd.ExcelFile(uploaded_file)
        sheets = {}
        for sheet_name in xls.sheet_names:
            sheets[sheet_name] = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        return sheets
    else:
        raise ValueError("Unsupported file format")

def main():
    st.title("Excel/CSV Question Answering System")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Excel or CSV file", type=['csv', 'xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Load file
        try:
            sheets = load_file(uploaded_file)
            
            # Sheet selection
            sheet_name = st.selectbox("Select Sheet", list(sheets.keys()))
            df = sheets[sheet_name]
            
            # Display dataframe preview
            st.subheader("Data Preview")
            st.dataframe(df.head())
            
            # Initialize LLM
            llm = ChatOpenAI(temperature=0, model="gpt-3.5-turbo-0125", api_key=OPENAI_API_KEY)
            
            # Create pandas agent for structured queries
            agent = create_pandas_dataframe_agent(
                llm, 
                df, 
                verbose=True,
                allow_dangerous_code=True  # Note: Only enable this in a trusted environment
            )
            
            # Query input
            query = st.text_input("Ask a question about your data:")
            
            if query:
                try:
                    with st.spinner("Analyzing your data..."):
                        response = agent.run(query)
                        st.write("Analysis Result:", response)
                
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
        
        except Exception as e:
            st.error(f"Error loading file: {str(e)}")

if __name__ == "__main__":
    main()