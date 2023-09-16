
from docx import Document
import os

def get_output():
    with open('output.txt', 'r', encoding='utf-8') as f:
        lines = f.readlines()
        return [x.lower() for x in lines]
    return []

def target_exist(lines=[], key_str="", id_str="", target_str=""):
    match_str = key_str + id_str + target_str
    match_str = match_str.lower()
    for line in lines:
        if match_str in line:
            return True
    return False

def save_to_docx():
    document = Document(docx=os.path.join(os.getcwd(), 'temple.docx'))    

    # for l, line in enumerate(get_output()):
    #     print(l)
    #     print(line)

    # paragraphs = document.paragraphs
    # for p in paragraphs:
    #     print(p.text)■

    for i, row in enumerate(document.tables[3].rows):
        for j, cell in enumerate(row.cells):
            for p in cell.paragraphs:
                print(i,j)
                print(p.text)

_item_col = 1
_expect_col = 3
_result_col = 4
_conclusion_col = 5

def generate_docx():
    document = Document(docx=os.path.join(os.getcwd(), 'temple.docx'))
    lines = get_output()
    for i, row in enumerate(document.tables[2].rows):
        if(i == 0):
            continue

        pass_parag = row.cells[_conclusion_col].paragraphs[0]
        fail_parag = pass_parag
        for p in row.cells[_conclusion_col].paragraphs:
            if("□合格" in p.text):
                pass_parag = p
            elif("□不合格" in p.text):
                fail_parag = p

        for j, p in enumerate(row.cells[_expect_col].paragraphs):
            result_run = row.cells[_result_col].paragraphs[j].runs[0]
            for run in row.cells[_result_col].paragraphs[j].runs:
                if(run.underline):
                    result_run = run
                    break

            if(target_exist(lines=lines, target_str=p.text)):
                result_run.text = " Ok "
                pass_parag.text="■ 合格"
            else:
                result_run.text = " Error "
                fail_parag.text="■ 不合格"
    document.save("test.docx")

def main():
    generate_docx()


main()