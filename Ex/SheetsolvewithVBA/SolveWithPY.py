import argparse
from docx import Document

def load_answers(filename):
    # Open the file with UTF-8 encoding
    with open(filename, 'r', encoding='utf-8') as file:
        answers = [line.strip() for line in file if line.strip()]
    return answers

def bold_underline_answers(doc_path, answers, output_path):
    # Load the Word document
    doc = Document(doc_path)
    
    # Loop through each paragraph in the document
    for paragraph in doc.paragraphs:
        for answer in answers:
            if answer in paragraph.text:
                # Apply bold and underline to each word that matches an answer
                inline = paragraph.runs
                for i in range(len(inline)):
                    if answer in inline[i].text:
                        inline[i].bold = True
                        inline[i].underline = True

    # Save the modified document
    doc.save(output_path)
    print(f"Processed document saved as '{output_path}'")

def main():
    parser = argparse.ArgumentParser(description="Bold and underline answers in a Word document.")
    parser.add_argument("doc_path", help="Path to the input Word document (e.g., sheet.doc)")
    parser.add_argument("answers_path", help="Path to the answers text file (e.g., Answers.txt)")
    parser.add_argument("output_path", help="Path to save the output Word document (e.g., solved_sheet.doc)")

    args = parser.parse_args()

    # Load answers from the text file
    answers = load_answers(args.answers_path)

    # Process the document
    bold_underline_answers(args.doc_path, answers, args.output_path)

if __name__ == "__main__":
    main()
