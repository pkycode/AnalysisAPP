from dotenv import load_dotenv
import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
from langchain_openai import ChatOpenAI
from langchain_experimental.agents import create_pandas_dataframe_agent
from pymongo import MongoClient
import xlrd
import openpyxl

# Load environment variables
load_dotenv()
OPENAI_API_KEY = os.getenv("API_KEY")

def get_database():
    try:
        mongo_uri = os.getenv("MONGODB_URI")
        if not mongo_uri:
            st.error("Missing MongoDB URI in environment variables")
            return None
        client = MongoClient(mongo_uri)
        client.server_info()
        return client.excel_analyzer_db
    except Exception as e:
        st.error(f"Failed to connect to database: {str(e)}")
        return None

def validate_email(email):
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, email) is not None

def log_email(email):
    try:
        db = get_database()
        if db is not None:
            user_collection = db.users
            timestamp = datetime.now()
            result = user_collection.insert_one({
                "email": email,
                "timestamp": timestamp
            })
            return bool(result.inserted_id)
        return False
    except Exception as e:
        st.error(f"Failed to log email: {str(e)}")
        return False

def log_question(email, question, answer):
    try:
        db = get_database()
        if db is not None:
            question_collection = db.questions
            timestamp = datetime.now()
            result = question_collection.insert_one({
                "email": email,
                "question": question,
                "answer": answer,
                "timestamp": timestamp
            })
            return bool(result.inserted_id)
        return False
    except Exception as e:
        st.error(f"Failed to log question: {str(e)}")
        return False

def load_file(uploaded_file):
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

def create_agent(llm, df):
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

    return create_pandas_dataframe_agent(
        llm,
        df,
        verbose=True,
        handle_parsing_errors=True,
        max_iterations=5,
        allow_dangerous_code=True,
        prefix=prefix
    )

def main():
    st.title("Excel/CSV Question Answering System")
    
    if 'email_verified' not in st.session_state:
        st.session_state.email_verified = False
    if 'user_email' not in st.session_state:
        st.session_state.user_email = None

    if not st.session_state.email_verified:
        st.write("Please enter your email to continue:")
        email = st.text_input("Email Address")
        
        if st.button("Submit Email"):
            if validate_email(email):
                if log_email(email):
                    st.session_state.email_verified = True
                    st.session_state.user_email = email
                    st.success("Email verified successfully!")
                    st.rerun()
                else:
                    st.error("Failed to save email. Please try again.")
            else:
                st.error("Please enter a valid email address")
        return

    uploaded_file = st.file_uploader("Upload Excel or CSV file", type=['csv', 'xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            sheets = load_file(uploaded_file)
            sheet_name = st.selectbox("Select Sheet", list(sheets.keys()))
            df = sheets[sheet_name]
            
            st.subheader("Data Preview")
            st.dataframe(df.head())
            
            llm = ChatOpenAI(
                temperature=0,
                model="gpt-4",
                api_key=OPENAI_API_KEY
            )
            
            agent = create_agent(llm, df)
            query = st.text_input("Ask a question about your data:")
            
            if query:
                try:
                    with st.spinner("Analyzing your data..."):
                        formatted_query = f"""Please analyze this and provide a clear, friendly response: {query}
                        Remember to:
                        - Start with a brief explanation
                        - Use bullet points for lists
                        - Format numbers properly
                        - Keep it conversational but precise"""
                        
                        response = agent.run(formatted_query)
                        
                        if response:
                            response = response.replace("Here's what I found:", "").strip()
                            response = response.replace("Analysis Result:", "").strip()
                            st.write(response)
                            
                            log_question(
                                st.session_state.user_email,
                                query,
                                response
                            )
                
                except Exception as e:
                    error_msg = str(e)
                    if "parsing errors" in error_msg.lower():
                        try:
                            match = re.search(r'Could not parse LLM output: `(.*?)`', error_msg)
                            if match:
                                actual_response = match.group(1)
                                st.write(actual_response)
                            else:
                                st.error("Failed to process the response.")
                        except:
                            st.error("An error occurred. Please try again.")
                    else:
                        st.error(f"An error occurred: {error_msg}")
        
        except Exception as e:
            st.error(f"Error loading file: {str(e)}")

if __name__ == "__main__":
    main()
