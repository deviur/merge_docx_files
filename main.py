import docx
"""
TODO: write notes here ⤵
Файл doc покетом docx открыть нельзя.
"""
doc = docx.Document()
doc.add_paragraph('Это первый абзац')
doc.save('example.docx')
doc = docx.Document('example.docx')
doc.add_paragraph('Это второй абзац')
doc.save('example2.docx')


def append(from_where: docx.Document, to_where: docx.Document()):
    for from_sec in from_where.sections:
        to_sec = to_where.add_section()
        to_sec = from_sec


if __name__ == '__main__':
    to_doc = docx.Document()

    from_doc = docx.Document('example.docx')
    append(from_doc, to_doc)

    from_doc = docx.Document('example2.docx')
    append(from_doc, to_doc)

    to_doc.save('example3.docx')