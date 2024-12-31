import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import pandas as pd

pdf_file_path = '190+_Python_MCQs.pdf'
pages = convert_from_path(pdf_file_path)
data = []
for page in range(3, len(pages)):
    question, explanation, options1, options2, options3, options4,correct_answer  = '', '', '', '', '', '', ''
    parsing_question = False
    parsing_explanation = False
    try:
        page_text = pytesseract.image_to_string(pages[page])
        page_data = page_text.split('\n')
        for line in page_data:
            line= line.strip()
            if line.startswith('Question:'):
                question += line
                question = question.split("Question:")[-1]
                parsing_question =True
                parsing_explanation = False
            elif line.startswith('Option 1:'):  # option line
                options1 += line
                options1 = options1.split('Option 1:')[-1]
                parsing_question = False
            elif line.startswith('Option 2:'):  # option line
                options2 += line
                options2 = options2.split('Option 2:')[-1]
                parsing_question = False
            elif line.startswith('Option 3:'):  # option line
                options3 += line
                options3 = options3.split('Option 3:')[-1]
                parsing_question = False
            elif line.startswith('Option 4:'):  # option line
                options4 += line
                options4 = options4.split('Option 4:')[-1]
                parsing_question = False
            elif line.startswith('Correct Response'):  # correct Response line
                correct_answer += line
                correct_answer = correct_answer.split('Correct Response:')[-1]
                parsing_question = False
            elif line.startswith('Explanation:'):  # Explanation line
                explanation += line
                parsing_explanation =True
                parsing_question = False
            else:
                if parsing_question:
                    question += ' '+line
                    question = question.split("Question:")[-1]
                elif parsing_explanation:
                    explanation += ' '+line
                    explanation = explanation.split("Explanation:")[-1]
    except Exception as e:
        print(f'Inside Exception block: {str(e)}')
    data.append(
        {'Question': question, 'Option 1': options1, 'Option 2': options2, 'Option 3': options3, 'Option 4': options4,
         'Correct_answer': correct_answer, 'Explanation': explanation})
try:
    df = pd.DataFrame(data)
    df.to_excel('extracted_data1.xlsx')
except Exception as e:
    print(f'{str(e)}')