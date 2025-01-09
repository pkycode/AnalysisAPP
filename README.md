# Excel & CSV Analyzer with LLM

A streamlined application that allows users to have natural conversations with their Excel and CSV data using GPT-3.5. Upload your spreadsheet and ask questions in plain English to get instant insights.

## Features

- 📊 Support for both Excel (.xlsx, .xls) and CSV files
- 📑 Multi-sheet Excel file support
- 💬 Natural language querying of your data
- 🔍 Instant data analysis and insights
- 📈 Accurate numerical calculations and statistics
- 💫 User-friendly interface

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/excel-csv-analyzer.git
cd excel-csv-analyzer
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the root directory and add your OpenAI API key:
```
API_KEY=your-openai-api-key-here
```

## Usage

1. Start the application:
```bash
streamlit run app.py
```

2. Open your web browser and navigate to the displayed local URL (typically http://localhost:8501)

3. Upload your Excel or CSV file

4. Select the sheet you want to analyze (for Excel files)

5. Start asking questions about your data!

### Example Questions

- "What's the total revenue by product?"
- "Which region has the highest sales?"
- "Show me the trend of sales over time"
- "What's the average customer satisfaction rating?"
- "What are the top 5 performing products?"

## Technologies Used

- [Streamlit](https://streamlit.io/) - For the web interface
- [LangChain](https://python.langchain.com/) - For LLM integration
- [OpenAI GPT-3.5](https://openai.com/) - For natural language processing
- [Pandas](https://pandas.pydata.org/) - For data manipulation
- Python libraries for Excel/CSV handling (xlrd, openpyxl)

## Requirements

See `requirements.txt` for a complete list of dependencies.

## Contributing

Feel free to submit issues, fork the repository, and create pull requests for any improvements.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Security Note

This application uses OpenAI's API for processing queries. Ensure you handle your API keys securely and review any sensitive data before uploading.

## Acknowledgments

- OpenAI for providing the GPT API
- LangChain for the excellent framework
- Streamlit for making web apps simple to build
