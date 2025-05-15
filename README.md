# Text Document Analyzer
![Grammar Correction](https://img.shields.io/badge/Grammar-Correction-blue)
![Document Processing](https://img.shields.io/badge/Document-Processing-green)
![Streamlit App](https://img.shields.io/badge/Streamlit-App-red)
![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Python Version](https://img.shields.io/badge/python-3.8%2B-brightgreen.svg)

> An intelligent document processing tool that corrects grammar in text documents while preserving the original formatting.

## 📝 Overview

Text Document Analyzer is a powerful application designed to help users improve the grammatical quality of their documents without losing formatting elements like tables, images, and styles. The tool analyzes DOCX and DOC files, identifies and corrects grammar errors, and generates detailed reports on the types of corrections made.

Access the application: [TextDocumentAnalyzer](https://textdocanalyer.streamlit.app/)

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
  - Identifies 20+ types of grammar errors (punctuation, capitalization, etc.)
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

```mermaid
flowchart LR
    %% Main user flow
    User([User]):::user --> Upload[Upload Document]:::action
    Upload --> Select[Select Mode]:::action
    Select --> Instruct[Add Instructions]:::action
    Instruct --> Process[Process Document]:::action
    
    %% Processing paths
    Process --> Decision{Compatibility Mode}:::decision
    
    Decision -->|Preserve All| PreserveMode[XMLDocumentCorrector]:::preserve
    Decision -->|Safe Mode| SafeMode[Text-based Correction]:::safe
    Decision -->|Ultra Safe| UltraMode[Basic Text Correction]:::ultrasafe
    
    %% Preserve All pathway
    PreserveMode --> PA1[Extract to temp directory]:::preserve
    PA1 --> PA2[Process document.xml]:::preserve
    PA2 --> PA3[Batch paragraph processing]:::preserve
    PA3 --> PA4[Fix headers & footers]:::preserve
    PA4 --> PA5[Create fully formatted DOCX]:::preserve
    
    %% Safe Mode pathway
    SafeMode --> SA1[Extract text content]:::safe
    SA1 --> SA2[Split into sections]:::safe
    SA2 --> SA3[LM grammar correction]:::safe
    SA3 --> SA4[Create basic formatted doc]:::safe
    
    %% Ultra Safe pathway
    UltraMode --> US1[Extract plain text]:::ultrasafe
    US1 --> US2[LM correction]:::ultrasafe
    US2 --> US3[Create clean document]:::ultrasafe
    
    %% Analysis stage
    PA5 --> Analysis[Analyze Corrections]:::analysis
    SA4 --> Analysis
    US3 --> Analysis
    
    Analysis --> AN1[Compare texts]:::analysis
    AN1 --> AN2[Identify errors]:::analysis
    AN2 --> AN3[Generate visualizations]:::analysis
    
    %% Results
    AN3 --> Present[Present Results]:::results
    Present --> Download[Download Corrected Doc]:::results
    Present --> ViewReport[View Error Report]:::results
    
    %% Language Model API
    PA3 -.->|API Requests| LM[(Language Model API)]:::lm
    SA3 -.->|API Requests| LM
    US2 -.->|API Requests| LM
    AN2 -.->|Error Classification| LM
    
    %% Classes & styling
    classDef user fill:#F8F8FF,stroke:#333,stroke-width:2px,color:#333
    classDef action fill:#D7E9F7,stroke:#1D3557,stroke-width:2px,color:#1D3557
    classDef decision fill:#FFE6E6,stroke:#9E2A2B,stroke-width:2px,color:#9E2A2B
    classDef preserve fill:#DAEEC7,stroke:#1B4332,stroke-width:2px,color:#1B4332
    classDef safe fill:#FFF3B0,stroke:#E09F3E,stroke-width:2px,color:#8B4513
    classDef ultrasafe fill:#FFD6D6,stroke:#9E2A2B,stroke-width:2px,color:#9E2A2B
    classDef analysis fill:#E9D8FD,stroke:#553C9A,stroke-width:2px,color:#553C9A
    classDef results fill:#FFC6FF,stroke:#9D4EDD,stroke-width:2px,color:#9D4EDD
    classDef lm fill:#FFE6E6,stroke:#9E2A2B,stroke-width:2px,color:#9E2A2B
```

## 💻 Technical Architecture

The application is built with a modular architecture that separates concerns and promotes maintainability:

```mermaid
graph TB
    %% Node definitions with styling
    subgraph UI["UI Layer"]
        style UI fill:#D7E9F7,stroke:#1D3557,stroke-width:2px
        View["DocumentCorrectionAppView<br/>- Renders interface<br/>- Handles user input<br/>- Displays results"]
        style View fill:#D7E9F7,stroke:#1D3557,stroke-width:2px,color:#1D3557
    end
    
    subgraph Core["Core Processing"]
        style Core fill:#D4F1F9,stroke:#05445E,stroke-width:2px
        Processor["DocumentProcessorForm<br/>- Coordinates document processing<br/>- Determines process workflow<br/>- Integrates components"]
        style Processor fill:#D4F1F9,stroke:#05445E,stroke-width:2px,color:#05445E
    end
    
    subgraph DocHandling["Document Handling"]
        style DocHandling fill:#FFF3B0,stroke:#E09F3E,stroke-width:2px
        XML["XMLDocumentCorrector<br/>- Processes XML structure<br/>- Preserves all formatting"]
        style XML fill:#FFF3B0,stroke:#E09F3E,stroke-width:2px,color:#8B4513
        
        Creator["SafeDocxCreatorForm<br/>- Creates safe document<br/>- Basic formatting preservation"]
        style Creator fill:#E9D8FD,stroke:#553C9A,stroke-width:2px,color:#553C9A
        
        Validator["DocxValidator<br/>- Validates document structure<br/>- Fixes corrupted files"]
        style Validator fill:#FFD6D6,stroke:#9E2A2B,stroke-width:2px,color:#9E2A2B
    end
    
    subgraph TextProc["Text Processing"]
        style TextProc fill:#DAEEC7,stroke:#1B4332,stroke-width:2px
        Corrector["GrammarCorrectorForm<br/>- Handles text correction<br/>- Interfaces with Language Models<br/>- Processes text in sections"]
        style Corrector fill:#DAEEC7,stroke:#1B4332,stroke-width:2px,color:#1B4332
    end
    
    subgraph Analysis["Analysis & Reporting"]
        style Analysis fill:#FFC6FF,stroke:#9D4EDD,stroke-width:2px
        Analyzer["DocumentAnalyzerForm<br/>- Analyzes corrections<br/>- Identifies error types<br/>- Generates reports"]
        style Analyzer fill:#FFC6FF,stroke:#9D4EDD,stroke-width:2px,color:#9D4EDD
    end
    
    subgraph External["External Services"]
        style External fill:#FFE6E6,stroke:#9E2A2B,stroke-width:2px
        LM["Language Model API<br/>- Grammar correction<br/>- Error classification"]
        style LM fill:#FFE6E6,stroke:#9E2A2B,stroke-width:2px,color:#9E2A2B
    end
    
    %% Major data flows
    View --> Processor
    Processor --> View
    
    %% Core processor connections
    Processor --> Corrector
    Processor --> XML
    Processor --> Validator
    Processor --> Creator
    Processor --> Analyzer
    
    %% Mode-specific workflows
    Processor --> XML
    Processor --> Creator
    Processor --> Creator
    
    %% Component interactions
    XML --> Corrector
    Analyzer --> Corrector
    Corrector --> LM
```

## 🛠️ Installation

```bash
# Clone the repository
git clone https://github.com/username/TextDocumentAnalyzer.git
cd TextDocumentAnalyzer

# Create and activate virtual environment (optional but recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## 📋 Requirements

The application requires the following dependencies:

### requirements.txt
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
streamlit run TextDocumentAnalyzer.py
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

```mermaid
graph LR
    A[Original & Corrected Text] --> B[Compare Paragraphs]
    B --> C[Detect Changed Paragraphs]
    
    C --> D[Pattern-Based Error Detection]
    D --> E{Error Types Detected?}
    
    E -->|Yes| F[Categorize Error Types]
    E -->|No| G[LM-Based Error Classification]
    G --> F
    
    F --> H[Generate Error Statistics]
    H --> I[Create Visual Reports]
    
    I --> J[Pie Chart: Error Types]
    I --> K[Bar Chart: Top Errors]
    I --> L[Detailed Error List]
    
    %% Add colors to main flow elements
    style A fill:#f5f5f5,stroke:#d9d9d9,stroke-width:2px
    style B fill:#e5f2ff,stroke:#99ccff,stroke-width:2px
    style C fill:#e5f2ff,stroke:#99ccff,stroke-width:2px
    style D fill:#ffe6cc,stroke:#ffc266,stroke-width:2px
    style E fill:#ffe6cc,stroke:#ffc266,stroke-width:2px
    style F fill:#e6ffe6,stroke:#b3ffb3,stroke-width:2px
    style G fill:#ffe6e6,stroke:#ffb3b3,stroke-width:2px
    style H fill:#e6f2ff,stroke:#b3d9ff,stroke-width:2px
    style I fill:#e6f2ff,stroke:#b3d9ff,stroke-width:2px
    style J fill:#f2e6ff,stroke:#d9b3ff,stroke-width:2px
    style K fill:#f2e6ff,stroke:#d9b3ff,stroke-width:2px
    style L fill:#f2e6ff,stroke:#d9b3ff,stroke-width:2px
```

## 🌐 Grammar Correction Process

The application uses language models to correct grammar while preserving meaning and style:

```mermaid
graph LR
    %% Define node styles
    classDef input fill:#D4F1F9,stroke:#05445E,stroke-width:2px,color:#05445E,font-weight:bold
    classDef process fill:#DAEEC7,stroke:#1B4332,stroke-width:2px,color:#1B4332
    classDef api fill:#FFE6E6,stroke:#9E2A2B,stroke-width:2px,color:#9E2A2B
    classDef decision fill:#FFF3B0,stroke:#E09F3E,stroke-width:2px,color:#8B4513,font-weight:bold
    classDef output fill:#E9D8FD,stroke:#553C9A,stroke-width:2px,color:#553C9A
    classDef analysis fill:#FFC6FF,stroke:#9D4EDD,stroke-width:2px,color:#9D4EDD
    classDef error fill:#FFD6D6,stroke:#9E2A2B,stroke-width:2px,color:#9E2A2B
    
    A[Text Input]:::input --> B[Split into Sections]:::process
    B --> C[Create Correction Prompt]:::process
    
    subgraph "For Each Section"
        C -->|API Request| D[Language Model]:::api
        D -->|Response| E[Extract Corrected Text]:::process
        E -->|Error?| F{Retry?}:::decision
        F -->|Yes: Attempts < 3| C
        F -->|No: Max Retries| G[Use Original Text]:::error
        E -->|Success| H[Add to Results]:::output
    end
    
    H --> I[Combine Sections]:::output
    G --> I
    I --> J[Apply Format Preservation]:::output
    J --> K[Return Corrected Document]:::output
    
    subgraph "Error Analysis"
        K --> L[Compare Original & Corrected]:::analysis
        L --> M[Pattern-Based Error Detection]:::analysis
        M --> N[LM-Based Error Classification]:::analysis
        N --> O[Generate Error Statistics]:::analysis
        O --> P[Create Analysis Report]:::analysis
    end
```

## 📖 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgements

- Grammar correction is powered by LM Studio and supports various language models
- Document processing utilizes python-docx and lxml for XML manipulation
- Visualization components use Plotly and Matplotlib

## Contact Information
- Author: [Srinivas K M](https://github.com/srini1812)
- Email id: 1812srini@gmail.com
