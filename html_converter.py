import os

from docx2html import convert


def docx_converter():

    while True:
        docx_path = raw_input('Please enter path to folder containing docx files:')
        path_exists = os.path.exists(docx_path)

        if path_exists:
            for subdir, dirs, files in os.walk(docx_path):
                for file in files:
                    ext = os.path.splitext(file)[-1].lower()

                    if ext == '.docx':
                        file_path = os.path.join(docx_path, file)
                        html = convert(file_path)

                        return html

        else:
            print 'Please enter a valid path.'

docx_converter()
