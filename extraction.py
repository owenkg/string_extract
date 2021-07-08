import textract
import os
import csv
import sys

#THE FUNCTION BELOW EXTRACTS SENTENCES FROM A WORD DOCUMENT AND PUTH THEM IN A CSV FILE
#IT DOES THIS BY FIRST CONVERTING THE WORD DOCUMENT TO A TEXT DOCUMENT
#THEN PICKING SENTENCE BY SENTENCE AND THE PLACING IT IN THE CSV FILE

def extraction():
    # this folder contains the document files(.doc, .docx) containing the sentences.
    source_directory = os.path.join(os.getcwd(), "source")
    # this is a temporary folder created to hold the text files(.txt) generated from the documents.
    docs_folder = os.mkdir("working_data")
    working_directory = os.path.join(os.getcwd(), "working_data")
    dest_file_path = ""
    sentences = []
    i = 0


    # this loop reads all the document files in the source folder and generates .txt files for each document file
    for process_file in os.listdir(source_directory):
        file, extension = os.path.splitext(process_file)
        # print(file)
        # print(extension)
        if extension == ".docx":
            # We create a new text file name by concatenating the .txt extension to file
            dest_file_path = file + '.txt'
            # print(dest_file_path)
            # extract text from the file
            content = textract.process(
                os.path.join(source_directory, process_file))

            # this opens the new txt file for writing to using the mode wb - Write
            output_txt = open(os.path.join(
                working_directory, dest_file_path), "wb")

            # write the content and close the newly created file
            output_txt.write(content)
            output_txt.close()
        else:
            # print()
            sys.exit("No Document Files Provided!")

    # read a whole paragraph from the text document and put it in a list (sentences)
    with open(os.path.join(working_directory, dest_file_path), 'r') as source_file:
        sentences = source_file.readlines()


    for sentence in sentences:
        # if a paragraph is an empty line, then skip it
        if sentence == "\n" or sentence == "\n\n" or sentence == "\n\n\n" or sentence == "\n\n\n" or sentence == "\n\n\n\n":
            continue
        else:
            # separate individual sentence from paragraph(long_sentence) by using fullstop(.)
            long_sentence = sentence.split(". ")
            #print(f'line {i}: {sentence}')
            # print(long_sentence)
            for ind_sent in long_sentence:
                # print(ind_sent)
                # create and open CSV file to which sentences are to be written
                with open('sentences.csv', 'a', newline='') as dest_file:
                    writer = csv.writer(dest_file)
                    if ind_sent == "":
                        continue
                    else:
                        # write each individual sentence in its own row
                        writer.writerow([ind_sent])

        i += 1


    print("Sentences copied successfully!")
    os.remove(os.path.join(working_directory, dest_file_path))
    os.rmdir(working_directory)

extraction()