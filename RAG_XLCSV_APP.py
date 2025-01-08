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

# MongoDB setup
def get_database():
    try:
        # Get MongoDB connection string from .env file
        mongo_uri = os.getenv("MONGODB_URI")
        
        if not mongo_uri:
            st.error("Missing MongoDB URI in environment variables")
            return None
            
        client = MongoClient(mongo_uri)
        # Test the connection
        client.server_info()
        return client.excel_analyzer_db
    except Exception as e:
        st.error(f"Failed to connect to database: {str(e)}")
        return None

def validate_email(email):
    """Validate email format using regex"""
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, email) is not None

def log_email(email):
    """Log email with timestamp to MongoDB"""
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
    """Log user questions and answers to MongoDB"""
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
    
    # Initialize session state for email verification
    if 'email_verified' not in st.session_state:
        st.session_state.email_verified = False
    if 'user_email' not in st.session_state:
        st.session_state.user_email = None

    # Email verification section
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

    # Main application (only shown after email verification)
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
                        
                        # Log the question and answer
                        log_question(
                            st.session_state.user_email,
                            query,
                            response
                        )
                
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
        
        except Exception as e:
            st.error(f"Error loading file: {str(e)}")

if __name__ == "__main__":
    main()
