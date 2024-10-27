import os
import argparse
import win32com.client as win32

def load_answers(filename):
    with open(filename, 'r', encoding='utf-8') as file:
        answers = [line.strip() for line in file if line.strip()]
    return answers

def add_vba_macro(doc_path, answers):
    # Split answers into four lists
    answerList1 = answers[:len(answers)//4]
    answerList2 = answers[len(answers)//4:len(answers)//2]
    answerList3 = answers[len(answers)//2:3*len(answers)//4]
    answerList4 = answers[3*len(answers)//4:]

    # Updated VBA Macro Code
    vba_code = f'''
Sub BoldUnderlineAnswers()
    Dim answerList1 As Variant, answerList2 As Variant, answerList3 As Variant, answerList4 As Variant
    Dim answer As Variant
    Dim rng As Range
    Dim doc As Document

    ' Define answers for each question in four arrays
    answerList1 = Array({', '.join(f'"{answer}"' for answer in answerList1)})
    answerList2 = Array({', '.join(f'"{answer}"' for answer in answerList2)})
    answerList3 = Array({', '.join(f'"{answer}"' for answer in answerList3)})
    answerList4 = Array({', '.join(f'"{answer}"' for answer in answerList4)})

    Set doc = ActiveDocument

    ' Function to process each answer list
    Call ProcessAnswers(doc, answerList1)
    Call ProcessAnswers(doc, answerList2)
    Call ProcessAnswers(doc, answerList3)
    Call ProcessAnswers(doc, answerList4)

    MsgBox "Answers have been bolded and underlined!"
End Sub

Sub ProcessAnswers(doc As Document, answers As Variant)
    Dim answer As Variant
    Dim rng As Range

    ' Loop through each answer
    For Each answer In answers
        Set rng = doc.Content
        With rng.Find
            .Text = answer
            .Replacement.Text = answer
            .Replacement.Font.Bold = True
            .Replacement.Font.Underline = wdUnderlineSingle
            .Execute Replace:=wdReplaceAll
        End With
    Next answer
End Sub
'''

    # Open Word application and document
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(doc_path)

    # Add macro to the document
    vb_module = doc.VBProject.VBComponents.Add(1)  # 1 represents vbext_ct_StdModule
    vb_module.CodeModule.AddFromString(vba_code)

    # Run the macro
    word.Run("BoldUnderlineAnswers")

    # Save the document with changes
    doc.Save()

    # Remove the macro after execution
    doc.VBProject.VBComponents.Remove(vb_module)

    # Save final document without macros
    doc.SaveAs(doc_path, FileFormat=win32.constants.wdFormatDocument)
    doc.Close(False)
    word.Quit()

    print(f"Document processed and saved without macros: {doc_path}")

def main():
    parser = argparse.ArgumentParser(description="Add VBA macro to bold and underline answers in a Word document.")
    parser.add_argument("doc_path", help="Path to the input Word document (e.g., sheet.doc)")
    parser.add_argument("answers_path", help="Path to the answers text file (e.g., asnr.txt)")

    args = parser.parse_args()

    # Verify that both file paths exist
    doc_path = os.path.abspath(args.doc_path)  # Convert to absolute path
    answers_path = os.path.abspath(args.answers_path)  # Convert to absolute path

    if not os.path.exists(doc_path):
        print(f"Error: Document file '{doc_path}' does not exist.")
        return

    if not os.path.exists(answers_path):
        print(f"Error: Answers file '{answers_path}' does not exist.")
        return

    # Load answers from the text file
    answers = load_answers(answers_path)

    # Add VBA macro, run it, and remove it from the document
    add_vba_macro(doc_path, answers)

if __name__ == "__main__":
    main()
