# Doc AI

![Grammar Correction](https://img.shields.io/badge/Grammar-Correction-blue)
![Document Processing](https://img.shields.io/badge/Document-Processing-green)
![Streamlit App](https://img.shields.io/badge/Streamlit-App-red)
![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Python Version](https://img.shields.io/badge/python-3.8%2B-brightgreen.svg)

> An intelligent document processing tool that corrects grammar in text documents while preserving the original formatting.

## 📝 Overview

Doc AI optimizes document quality by correcting grammatical errors while maintaining full formatting integrity. Compatible with DOCX, DOC, and TXT formats, it automatically identifies issues and implements superior word selections. The solution generates comprehensive analytical reports of corrections, enabling enterprises to produce impeccable documentation without sacrificing structural or visual elements—enhancing both efficiency and professional presentation.

**Access the application:** [Doc AI Application](https://docai.streamlit.app/)

## ✨ Key Features

- **Grammar Correction with Format Preservation**
  - Accurately corrects grammar errors while maintaining document structure
  - Preserves tables, images, headers, footers, and styling in DOCX files
  - Maintains paragraph formatting, lists, and other structural elements

- **Multiple Compatibility Modes**
  - **Preserve All**: Maintains all document elements including XML structure
  - **Safe Mode**: Balances formatting preservation with compatibility
  - **Ultra Safe Mode**: Creates clean document with perfect compatibility

- **Detailed Error Analysis**
  - Identifies grammar errors (punctuation, capitalization, etc.)
  - Generates visual reports showing error distribution
  - Provides detailed paragraph-by-paragraph correction explanations

- **Custom Correction Instructions**
  - Allows users to specify correction style (academic, technical, casual)
  - Supports preservation of domain-specific terminology
  - Enables custom handling of contractions, formal language, etc.

- **Interactive User Interface**
  - Streamlined Streamlit interface with real-time processing updates
  - Side-by-side comparison of original and corrected text
  - Multiple result views with downloadable reports

## 🔍 How It Works

The application processes documents through a sophisticated pipeline that preserves formatting while correcting grammar:

![Document Processing Pipeline](https://github.com/user-attachments/assets/5da90239-b57b-49ca-9a41-30d3808a05e9)

## 💻 Technical Architecture

The application is built with a modular architecture that separates concerns and promotes maintainability:

![Technical Architecture](https://github.com/user-attachments/assets/43d17297-16c2-462f-a083-1274dd2179b5)

## 🛠️ Installation

```bash
# Clone the repository
git clone https://github.com/company-name/DocAI.git
cd DocAI

# Create and activate virtual environment (optional but recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## 📋 Requirements

The application requires the following dependencies:

```
streamlit==1.29.0
python-docx==0.8.11
pandas==2.1.0
matplotlib==3.7.2
plotly==5.15.0
lxml==4.9.3
requests==2.31.0
zipfile36==0.1.3
pathlib==1.0.1
```

## 🚀 Usage

```bash
# Start the application
streamlit run DocAI.py
```

Once running, access the application in your browser at `http://localhost:8501`.

### Using the Application

1. **Configure the Application:**
   - Set LM Studio API URL
   - Select model name
   - Adjust temperature for grammar correction

2. **Provide Additional Instructions (Optional):**
   - Specify academic style
   - Request preservation of technical terminology
   - Customize grammar correction style

3. **Choose Compatibility Mode:**
   - **Preserve All** - Maintain all formatting, tables, and images
   - **Safe Mode** - Balance formatting preservation with compatibility
   - **Ultra Safe Mode** - Create a completely new document with perfect compatibility

4. **Upload and Process Document:**
   - Upload your DOCX or DOC file
   - Click "Correct Grammar" to begin processing

5. **Review and Download Results:**
   - Download corrected document
   - Analyze error types with visual reports
   - View detailed correction report
   - Download analysis report in Markdown format

## 🔄 Processing Flow

This sequence diagram shows the detailed flow of document processing:

```mermaid
sequenceDiagram
    %% Define actors with emojis for visual distinction
    actor User as 🧑‍💻 User
    participant UI as 🖥️ DocumentCorrectionAppView
    participant Processor as 🔄 DocumentProcessorForm
    participant XMLCorrector as 📝 XMLDocumentCorrector
    participant Creator as 📄 SafeDocxCreatorForm
    participant Corrector as ✏️ GrammarCorrectorForm
    participant Analyzer as 📊 DocumentAnalyzerForm
    participant LM as 🧠 Language Model API
    
    %% Initial user interactions with UI
    User->>UI: Upload document
    Note over User,UI: User selects input file
    User->>UI: Select compatibility mode
    Note over User,UI: Choose processing approach
    User->>UI: Provide additional instructions
    User->>UI: Click "Correct Grammar"
    
    %% Processing setup
    UI->>+Processor: Create processor
    UI->>Processor: ProcessDocument(file, mode, instructions)
    Processor->>Processor: ExtractTextFromDoc(file)
    
    %% The three processing modes
    alt Preserve All Mode (Full formatting)
        Note over Processor,XMLCorrector: Preserves all document formatting
        
        Processor->>+XMLCorrector: CorrectDocument(file)
        XMLCorrector->>XMLCorrector: Extract document to temp dir
        XMLCorrector->>XMLCorrector: Process document.xml
        
        loop For each paragraph batch
            XMLCorrector->>+Corrector: CorrectGrammar(text, instructions)
            Corrector->>+LM: API Request with prompt
            LM-->>-Corrector: Return corrected text
            Corrector-->>XMLCorrector: Return corrected text
            XMLCorrector->>XMLCorrector: Update XML with corrections
        end
        
        XMLCorrector->>XMLCorrector: Process headers/footers
        XMLCorrector->>XMLCorrector: Create corrected DOCX
        XMLCorrector-->>-Processor: Return corrected document
        
    else Safe Mode (Basic formatting)
        Note over Processor,Creator: Maintains basic document formatting
        
        Processor->>+Corrector: CorrectTextInSections(text, instructions)
        
        loop For each text section
            Corrector->>+LM: API Request with prompt
            LM-->>-Corrector: Return corrected text
        end
        
        Corrector-->>-Processor: Return corrected text
        Processor->>+Creator: CreateSafeDocxWithFormatting(file, text)
        Creator-->>-Processor: Return document with basic formatting
        
    else Ultra Safe Mode (Minimal)
        Note over Processor,Creator: Creates clean document with minimal formatting
        
        Processor->>+Corrector: CorrectTextInSections(text, instructions)
        
        loop For each text section
            Corrector->>+LM: API Request with prompt
            LM-->>-Corrector: Return corrected text
        end
        
        Corrector-->>-Processor: Return corrected text
        Processor->>+Creator: CreateSafeDocx(text, text)
        Creator-->>-Processor: Return minimal document
    end
    
    %% Analysis phase
    Note over Processor,Analyzer: Analysis phase
    
    Processor->>+Analyzer: AnalyzeCorrections(originalText, correctedText)
    
    loop For each changed paragraph
        Analyzer->>Analyzer: DetectParagraphErrorTypes(change)
        opt If needed for error reasoning
            Analyzer->>+Corrector: Use corrector for LM reasoning
            Corrector->>+LM: API Request for error classification
            LM-->>-Corrector: Return error classification
            Corrector-->>-Analyzer: Return error reasoning
        end
    end
    
    Analyzer->>Analyzer: Generate error statistics
    Analyzer-->>-Processor: Return analysis results
    
    Processor->>+Analyzer: GenerateSummaryReport(analysis)
    Analyzer-->>-Processor: Return summary report
    
    %% Results presentation
    Processor-->>-UI: Return processing results
    UI->>User: Display results tabs (document, comparison, analysis)
    UI->>User: Enable download of corrected document
    
    %% Final note
    Note over User,LM: Document correction process complete
```

## 🧩 Key Components

1. **DocxValidator** - Ensures document XML structure is valid and fixes issues
2. **XMLDocumentCorrector** - Preserves formatting by correcting text within XML structure
3. **GrammarCorrectorForm** - Handles communication with the language model
4. **DocumentAnalyzerForm** - Analyzes and categorizes grammar corrections
5. **DocumentProcessorForm** - Orchestrates the document processing workflow
6. **DocumentCorrectionAppView** - Manages the Streamlit UI

## 📊 Error Analysis

The application performs comprehensive error analysis on the corrections:

![Error Analysis](https://github.com/user-attachments/assets/7a04b5e4-ed88-41ff-9230-df42c02c2182)

## 🌐 Grammar Correction Process

The application uses language models to correct grammar while preserving meaning and style:

![Grammar Correction Process](https://github.com/user-attachments/assets/b38f7145-2d27-45d5-9e26-021b88f4e89b)

## 📸 Application Screenshots

<details>
<summary>Click to view screenshots</summary>

### Upload Interface
![Upload Interface](https://github.com/srini1812/Testrepo/blob/main/Front%20page.png)

### Download Results
![Download Results](https://github.com/srini1812/Testrepo/blob/main/Corrected%20Document.png)

### Text Comparison
![Text Comparison](https://github.com/srini1812/Testrepo/blob/main/Text%20Comparison.png)

### Detailed Correction Report
![Detailed Correction Report](https://github.com/srini1812/Testrepo/blob/main/Detailed%20Error%20Report.png)

### Error Analysis Dashboard
![Error Analysis Dashboard](https://github.com/srini1812/Testrepo/blob/main/Error%20Analysis.png)
</details>

## 📖 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgements

- Grammar correction is powered by LM Studio and supports various language models
- Document processing utilizes python-docx and lxml for XML manipulation
- Visualization components use Plotly and Matplotlib

## 👤 Contact Information

- Author: [Srinivas K M](https://github.com/srini1812)
- Email: 1812srini@gmail.com
