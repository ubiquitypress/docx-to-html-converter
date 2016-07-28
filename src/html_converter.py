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
        docx_path_exists = os.path.exists(docx_path)

        if docx_path_exists:
            write_path = raw_input('Please enter path to write html files to (conversion begins automatically): ')
            write_path_exists = os.path.exists(write_path)

            if write_path_exists:
                    for subdir, dirs, files in os.walk(docx_path):
                        for file in files:
                            ext = os.path.splitext(file)[-1].lower()

                            if ext == '.docx':
                                file_path = os.path.join(docx_path, file)
                                html = convert(file_path, image_handler=handle_image)
                                print 'Converting: ' + file

                                # Give new html file same name as .docx original
                                html_file_path = os.path.join(write_path, os.path.splitext(file)[0])
                                html_file = open(html_file_path, 'w')
                                html_file.write(html)
                                html_file.close()

            else:
                print 'Please enter a valid path.'


        else:
            print 'Please enter a valid path.'


docx_converter()
