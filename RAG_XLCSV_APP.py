from dotenv import load_dotenv
import streamlit as st
import pandas as pd
import os
import re
from langchain_openai import ChatOpenAI
from langchain_experimental.agents import create_pandas_dataframe_agent
from langchain.agents.agent import AgentExecutor
from langchain.agents.agent_types import AgentType
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

def create_agent_with_error_handling(llm, df):
    """Create pandas agent with error handling and proper output formatting"""
    prefix = """You are a helpful and friendly data analysis expert. When analyzing data:
    1. Start with a brief explanation of your findings
    2. Present numerical results clearly with proper formatting:
       - Use bullet points (•) for lists
       - Include $ and 2 decimal places for currency
       - Sort numerical results from highest to lowest
    3. Never mention code, variables, or technical terms
    4. Keep responses conversational but precise
    
    Example good responses:
    Q: Show me the sales summary
    Looking at the overall sales performance, here's what I found:
    • Total Revenue: $1,234,567.89
    • Average Sales: $45,678.90
    • Highest Sale: $98,765.43
    • Lowest Sale: $12,345.67
    """

    # Create the agent with error handling
    agent = create_pandas_dataframe_agent(
        llm,
        df,
        agent_type=AgentType.OPENAI_FUNCTIONS,
        verbose=True,
        handle_parsing_errors=True,
        max_iterations=5,
        max_execution_time=60,
        allow_dangerous_code=True,
        prefix=prefix
    )
    
    # Wrap with AgentExecutor for better error handling
    executor = AgentExecutor.from_agent_and_tools(
        agent=agent,
        tools=agent.tools,
        verbose=True,
        handle_parsing_errors=True,
        max_iterations=5,
        early_stopping_method="force",
    )
    
    return executor

def main():
    st.title("Excel/CSV Question Answering System")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Excel or CSV file", type=['csv', 'xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            sheets = load_file(uploaded_file)
            
            # Sheet selection
            sheet_name = st.selectbox("Select Sheet", list(sheets.keys()))
            df = sheets[sheet_name]
            
            # Display dataframe preview
            st.subheader("Data Preview")
            st.dataframe(df.head())
            
            # Initialize LLM and create agent
            llm = ChatOpenAI(
                temperature=0,
                model="gpt-4",
                api_key=OPENAI_API_KEY
            )
            
            # Create agent with error handling
            agent = create_agent_with_error_handling(llm, df)
            
            # Query input
            query = st.text_input("Ask a question about your data:")
            
            if query:
                try:
                    with st.spinner("Analyzing your data..."):
                        # Format the query to encourage proper response formatting
                        formatted_query = f"""Please analyze this and provide a clear, friendly response: {query}
                        Remember to:
                        - Start with a brief explanation
                        - Use bullet points for lists
                        - Format numbers properly
                        - Keep it conversational but precise"""
                        
                        response = agent.run(formatted_query)
                        
                        # Clean up the response but preserve the explanation
                        if response:
                            response = response.replace("Here's what I found:", "").strip()
                            response = response.replace("Analysis Result:", "").strip()
                            st.write(response)
                
                except Exception as e:
                    error_msg = str(e)
                    if "parsing errors" in error_msg.lower():
                        # Handle parsing errors gracefully
                        try:
                            # Try to extract the actual response from the error message
                            match = re.search(r'Could not parse LLM output: `(.*?)`', error_msg)
                            if match:
                                actual_response = match.group(1)
                                st.write(actual_response)
                            else:
                                st.error("Failed to process the response. Please try rephrasing your question.")
                        except:
                            st.error("An error occurred while processing the response. Please try again.")
                    else:
                        st.error(f"An error occurred: {error_msg}")
        
        except Exception as e:
            st.error(f"Error loading file: {str(e)}")

if __name__ == "__main__":
    main()
