

# SheetsolverwithVBA
(Image)[VBA.png]
## 📖 Overview

**SheetsolverwithVBA** is a Python script that integrates with Microsoft Word to automatically bold and underline specified answers within a Word document using a dynamically generated VBA macro. This tool is particularly useful for educators and students who need to highlight answers in examination papers or assignments.

## 🚀 Features

- **Dynamic VBA Macro Generation**: Automatically generates and injects a VBA macro into the Word document to process answers.
- **Batch Processing**: Supports handling large lists of answers by splitting them into manageable groups.
- **Easy Integration**: Utilizes the `pywin32` library to interact with Word, allowing for seamless automation.
- **Automatic Cleanup**: Removes the macro after execution to maintain a clean document.

## 🛠️ Requirements

- Python 3.x
- `pywin32` library (install via `pip install pywin32`)

## 📦 Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/0xBugatti/VBARTTs/Ex/sheetsolverwithvba.git
   cd sheetsolverwithvba
   ```

2. Install the required package:

   ```bash
   pip install pywin32
   ```

## 📋 Usage

Run the script from the command line with the following syntax:

```bash
python sheetsolverwithvba.py <doc_path> <answers_path>
```

### Parameters:

- `doc_path`: Path to the input Word document (e.g., `sheet.doc`).
- `answers_path`: Path to the text file containing answers (e.g., `asnr.txt`).

### Example:

```bash
python sheetsolverwithvba.py sheet.doc asnr.txt
```

## 🔧 Example Answer File Format

The answers should be listed one per line in a text file (e.g., `asnr.txt`):

```
Answer 1
Answer 2
Answer 3
```

## 🛡️ Error Handling

The script includes error handling to check for the existence of both the Word document and the answers file. If any file is not found, a descriptive error message will be displayed.

## 🤝 Contributing

Contributions are welcome! If you have suggestions for improvements or new features, please submit a pull request or open an issue.
