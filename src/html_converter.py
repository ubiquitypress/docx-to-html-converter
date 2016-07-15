import os

from docx2html import convert
from shutil import copyfile


def handle_image(image_id, relationship_dict):
    image_path = relationship_dict[image_id]
    _, filename = os.path.split(image_path)
    destination_path = os.path.join('/tmp', filename)
    copyfile(image_path, destination_path)

    return 'file://%s' % destination_path

def docx_converter():

    while True:
        docx_path = raw_input('Please enter path to folder containing .docx files:')
        path_exists = os.path.exists(docx_path)

        if path_exists:
            for subdir, dirs, files in os.walk(docx_path):
                for file in files:
                    ext = os.path.splitext(file)[-1].lower()

                    if ext == '.docx':
                        file_path = os.path.join(docx_path, file)
                        html = convert(file_path, image_handler=handle_image)

                        return html

        else:
            print 'Please enter a valid path.'

docx_converter()

