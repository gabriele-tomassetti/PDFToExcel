import os
import subprocess
import sys
import camelot

def is_good_enough(table):
    if table['accuracy'] > 80.0 and table['whitespace'] < 25.0:
        return True
    else:
        return False

def main(argv):
    script_dir = os.path.dirname(os.path.realpath(__file__))
    print(script_dir + '/pdf_to_excel.py: Start')
    
    if len(sys.argv) > 1:
        start_dir = sys.argv[1]
    else:
        # we will process the current directory
        start_dir = '.'
    
    for dir_name, subdirs, file_list in os.walk(start_dir):    
        os.chdir(dir_name)
        for filename in file_list:
            file_ext = os.path.splitext(filename)[1]
            if file_ext == '.pdf':
                full_path = dir_name + '/' + filename
                new_filename = os.path.splitext(filename)[0] + "-OCR.pdf"
    
                # let's try to find some text into the PDF
                print("attempting to OCR: " + full_path)
                # add --deskew only if we are sure that the images might need it
                # because it might create large file together with the option --optmize 0
                cmd = ["ocrmypdf",  "--skip-text --deskew --rotate-pages --clean --optimize 0", '"' + filename + '"', '"' + new_filename + '"']
               
                # you should run OCRmyPDF as a command line tool
                # see https://ocrmypdf.readthedocs.io/en/latest/batch.html#api
                proc = subprocess.run(
                    cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
                result = proc.stdout
    
                if proc.returncode == 0:
                    print("OCR complete")
                elif proc.returncode == 6:
                    print("OCR processing skipped")
                else:
                    print("Error in OCR processing")                            
    
                
                print("Starting searching for tables in ", new_filename)
                # now we can detect tables
                lattice_tables = camelot.read_pdf(new_filename, process_background=True, flavor='lattice', pages='1-end')
                stream_tables = camelot.read_pdf(new_filename, flavor='stream', pages='1-end')

                # the final tables
                tables = []

                # we check whether the number of tables are the same
                if len(lattice_tables) > 0 and len(stream_tables) > 0 and len(lattice_tables) == len(stream_tables):
                    for index in range(len(lattice_tables)):
                        # we check whether the tables are both good enough
                        if is_good_enough(lattice_tables[index].parsing_report) and is_good_enough(stream_tables[index].parsing_report):        
                            # they probably represent the same table
                            if lattice_tables[index].parsing_report['page'] == stream_tables[index].parsing_report['page'] and lattice_tables[index].parsing_report['order'] == stream_tables[index].parsing_report['order']:
                                total_lattice = 1
                                total_stream = 1

                                for num in lattice_tables[index].shape:
                                    total_lattice *= num
                                for num in stream_tables[index].shape:
                                    total_stream *= num

                                # we pick the table with the most cells
                                if(total_lattice >= total_stream):
                                    tables.append(lattice_tables[index])
                                else:
                                    tables.append(stream_tables[index])

                        elif is_good_enough(lattice_tables[index].parsing_report):
                            tables.append(lattice_tables[index])

                        elif is_good_enough(stream_tables[index].parsing_report):
                            tables.append(stream_tables[index])
                elif len(lattice_tables) >= len(stream_tables):
                    tables = lattice_tables        
                else:
                    tables = stream_tables            

                if tables is not None and len(tables) > 0:
                    # let's check whether is TableList object or just a list of tables
                    if isinstance(tables, camelot.core.TableList) is False:
                        tables = camelot.core.TableList(tables)
                    
                    tables.export(os.path.splitext(new_filename)[0] + '.xls', f='excel')
    
    print(script_dir + '/pdf_to_excel.py: End')

if __name__ == '__main__':
    main(sys.argv)