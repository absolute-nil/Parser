from os import chdir, getcwd, listdir, path

from time import strftime

from win32com import client

import os

import PyPDF2



def count_files(filetype):

   
 
    count_files = 0

    for files in listdir(folder):

        if files.endswith(filetype):

            count_files += 1

    return count_files



def check_path(prompt):

    ''' (str) -> str

    Verifies if the provided absolute path does exist.

    '''

    abs_path = input(prompt)

    while path.exists(abs_path) != True:

        print ("\nThe specified path does not exist.\n")

        abs_path = input(prompt)

    return abs_path   


   

print ("\n")


folder = check_path("Provide absolute path for the folder: ")




chdir(folder)




num_docx = count_files(".docx")

num_doc = count_files(".doc")

num_pdf=count_files(".pdf")




if num_docx + num_doc +num_pdf  == 0:

    print ("\nThe specified folder does not contain docx,doc or pdf files.\n")

    print( "There are no files to convert. BYE, BYE!.")

    exit()

else:

    print ("\nNumber of doc,docx and pdf files: ", num_docx + num_doc +num_pdf, "\n")

    print ( "Starting to convert files ...\n")



try:

    word = client.DispatchEx("Word.Application")

    for files in listdir(getcwd()):

        if files.endswith(".docx"):

            new_name = files.replace(".docx", r".txt")

            in_file = path.abspath(folder + "\\" + files)

            new_file = path.abspath(folder + "\\" + new_name)

            doc = word.Documents.Open(in_file)

            print (strftime("%H:%M:%S"), " docx -> txt ", path.relpath(new_file))

           

            doc.SaveAs(new_file, FileFormat = 2)

            doc.Close()

        if files.endswith(".doc"):

            new_name = files.replace(".doc", r".txt")

            in_file = path.abspath(folder + "\\" + files)

            new_file = path.abspath(folder + "\\" + new_name)

            doc = word.Documents.Open(in_file)

            print(  " doc  -> txt ", path.relpath(new_file))

            doc.SaveAs(new_file, FileFormat = 2)

            doc.Close()





       



           

except Exception as e:

    print( e)

finally:

    word.Quit()


list=[]

directory=folder

for root,dirs,files in os.walk(directory):

    for filename in files:

        if filename.endswith('.pdf'):

            t=os.path.join(directory,filename)

            list.append(t)




for item in list:
    path=item

    head,tail=os.path.split(path)

    var="\\"

   

    tail=tail.replace(".pdf",".txt")

    name=head+var+tail

    

   

   

    

    content = ""

    

    pdf = PyPDF2.PdfFileReader(path, "rb")

    

    for i in range(0, pdf.getNumPages()):

        

        content += pdf.getPage(i).extractText() + "\n"
        

    print (strftime("%H:%M:%S"), " pdf  -> txt ")

    with open(name,'a') as out:
        out.write(content )


print ("\n", "Finished converting .doc .docx and .pdf files.")





# Count the number of txt files.


num_txt = count_files(".txt")  


print ("\nNumber of txt files: ", num_txt)

