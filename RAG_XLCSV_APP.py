# app.py
from dotenv import load_dotenv
import streamlit as st
import pandas as pd
import os
import re  # Added this import
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

def create_agent_with_strict_output(llm, df):
    """Create pandas agent with strict output formatting"""
    prefix = """You are a friendly data analysis expert. Be conversational yet professional.
    ALWAYS follow these rules:
    1. Start with a brief, clear explanation of what you found
    2. Present the results after your explanation
    3. For lists/groups, show each item on a new line with a bullet point (•)
    4. For currency, always use $ with 2 decimal places
    5. Sort numerical results from highest to lowest
    6. Never mention technical terms, code, or variables
    7. Be conversational but precise with numbers
    8. Focus on insights that would be valuable to the user
    
    Examples of good responses:
    Q: What's the revenue by product?
    Looking at the product revenue breakdown, Laptops are the highest revenue generator, followed by Phones and Tablets. Here are the specific figures:
    • Laptops: $5,230.50
    • Phones: $3,420.80
    • Tablets: $2,150.25

    Q: What's the average sales?
    After analyzing the sales data, I found that the average sales figure stands at $3,245.75, which represents the typical transaction value across all products.

    Q: How many units were sold by region?
    The sales distribution across regions shows that the North region leads in unit sales, while the West has the lowest number. Here's the complete breakdown:
    • North: 1,234 units
    • South: 987 units
    • East: 856 units
    • West: 654 units"""

    return create_pandas_dataframe_agent(
        llm,
        df,
        verbose=True,
        allow_dangerous_code=True,
        prefix=prefix
    )

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
            
            # Create pandas agent with strict output formatting
            agent = create_agent_with_strict_output(llm, df)
            
            # Query input
            query = st.text_input("Ask a question about your data:")
            
            if query:
                try:
                    with st.spinner("Analyzing your data..."):
                        # Enforce result-only response
                        enhanced_query = f"""ANALYZE AND SHOW ONLY THE RESULTS FOR: {query}
                        REMEMBER: 
                        - Show ONLY the calculated results
                        - NO explanations
                        - NO variables
                        - NO suggestions
                        - NO technical terms"""
                        
                        response = agent.run(enhanced_query)
                        # Clean up any redundant phrases but keep the explanation
                        response = response.replace("Analysis Result:", "").strip()
                        st.write(response)
                
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
        
        except Exception as e:
            st.error(f"Error loading file: {str(e)}")

if __name__ == "__main__":
    main()
