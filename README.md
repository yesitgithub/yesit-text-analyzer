# Text Document Analyzer
![Grammar Correction](https://img.shields.io/badge/Grammar-Correction-blue)
![Document Processing](https://img.shields.io/badge/Document-Processing-green)
![Streamlit App](https://img.shields.io/badge/Streamlit-App-red)

## Executive Summary

Text Document Analyzer is an enterprise-grade Streamlit application designed to enhance document quality by correcting grammatical errors while maintaining formatting integrity. The solution leverages advanced language models to analyze and improve DOCX, DOC, and TXT files while preserving document structure, formatting elements, tables, and other components essential for professional documentation.

Access the application: [TextDocumentAnalyzer](https://textdocanalyer.streamlit.app/)

Please note: The application is hosted on Streamlit Cloud and may require brief initialization time upon first access.

## Key Capabilities

- **Advanced Grammar Correction:** Identifies and resolves diverse grammatical issues while preserving intended meaning
- **Format Retention Technology:** Employs sophisticated XML processing to maintain document formatting, tables, images, and structural elements
- **Configurable Processing Modes:**
  - **Full Preservation:** Maintains comprehensive document formatting, tables, and visual elements
  - **Standard Mode:** Optimizes balance between formatting preservation and compatibility
  - **Enhanced Compatibility Mode:** Generates new documentation with maximum compatibility
- **Comprehensive Error Analytics:** Provides detailed error classification with data visualizations
- **Customizable Correction Parameters:** Implement specific guidelines to direct the grammar correction process
- **Multiple Export Options:** Download corrected documents in DOCX format alongside detailed error analysis reports

## Technical Specifications

### Core Components

1. **DocxValidator:** Ensures document integrity through validation and XML structure remediation
2. **XMLDocumentCorrector:** Performs precision modifications to document XML while maintaining formatting integrity
3. **GrammarCorrectorForm:** Interfaces with language model APIs to execute grammar corrections
4. **DocumentAnalyzerForm:** Conducts comprehensive analysis and categorization of grammatical errors
5. **DocumentProcessorForm:** Manages the end-to-end document processing workflow
6. **DocumentCorrectionAppView:** Provides intuitive Streamlit user interface for the application

### Error Categories Addressed

The solution identifies and corrects over 25 distinct categories of grammatical errors, including:
- Punctuation inaccuracies
- Capitalization inconsistencies
- Verb tense alignment issues
- Subject-verb agreement errors
- Article usage optimization
- Preposition selection errors
- Run-on sentence restructuring
- Sentence fragment completion
- Redundancy elimination
- Conciseness improvement
- Active voice enhancement
- Additional error categories

### System Architecture

The application implements a sophisticated multi-stage processing pipeline:
1. Document parsing with comprehensive validation
2. Text extraction with intelligent segmentation
3. Grammar correction via enterprise-grade language models
4. XML structure preservation with targeted modification
5. Error analysis with multilevel classification
6. Report generation with interactive visualization

## Implementation Guide

### System Requirements

- Python 3.7+
- Streamlit framework
- Language model server (e.g., LM Studio) with API access

### Deployment Instructions

```bash
# Repository acquisition
git clone https://github.com/yourusername/TextDocumentAnalyzer.git
cd TextDocumentAnalyzer

# Dependency installation
pip install -r requirements.txt

# Application launch
streamlit run TextDocumentAnalyzer.py
```

### Configuration Options

The solution provides configuration for:
- Language model API endpoint
- Model selection parameters
- Temperature calibration for correction precision
- Document compatibility mode selection
- Analysis visualization customization

## Operational Workflow

1. **Document Upload:** Users submit documents (.docx, .doc, or .txt)
2. **Configuration:** Establish processing parameters and specific requirements
3. **Processing:** The system processes the document, implementing grammatical improvements while preserving formatting
4. **Analysis:** Error types are systematically identified and visualized
5. **Delivery:** Access to corrected documentation and comprehensive error analysis

## Business Applications

- **Corporate Communications:** Enhance the quality of business correspondence, proposals, and reports
- **Legal Documentation:** Ensure grammatical precision in contracts and legal documents
- **Financial Reporting:** Maintain professionalism in financial statements and analysis
- **Technical Documentation:** Improve clarity in product specifications and technical manuals
- **Educational Institution Support:** Provide writing improvement tools with detailed feedback mechanisms

## Contribution Guidelines

Professional contributions are welcome. Please submit Pull Requests through standard channels.

## Requirements.txt
- streamlit
- python-docx
- pandas
- matplotlib
- plotly
- numpy
- requests
- lxml

## Contact Information
- Author: [Srinivas K M](https://github.com/srini1812)
- Email id: 1812srini@gmail.com
