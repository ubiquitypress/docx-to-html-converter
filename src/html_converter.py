import os
import subprocess

def docx_converter():

    BASE_DIR = os.path.dirname(os.path.dirname(__file__))

    while True:
        docx_path = raw_input('Please enter path to folder containing .docx files:')
        docx_path = os.path.join(BASE_DIR, docx_path)
        docx_path_exists = os.path.exists(docx_path)

        if docx_path_exists:
            write_path = raw_input('Please enter path to write html files to (conversion begins automatically): ')
            write_path = os.path.join(BASE_DIR, write_path)
            write_path_exists = os.path.exists(write_path)

            if write_path_exists:
                    for subdir, dirs, files in os.walk(docx_path):
                        for file in files:
                            ext = os.path.splitext(file)[-1].lower()

                            if ext == '.docx':
                                file_path = os.path.join(docx_path, file)
                                _id = file[:3]
                                print "pandoc -s {0} -t markdown -o {1}.md --extract-media={1}".format(file_path, _id)
                                

                                subprocess.Popen("pandoc -s {0} -t markdown -o sdo/{1}.md --extract-media={2}/{1}/".format(file_path, _id, write_path), shell=True)

                                print 'Converting: ' + file

            else:
                print 'Please enter a valid path.'


        else:
            print 'Please enter a valid path.'


docx_converter()
