# Masters-Project
The project for Masters in University of Leeds

Installation
--------------------------------------------------------------------------
1. Open command line terminal
2. type 'sudo pip install -U nltk' to install NLTK in Python
3. type 'pip install cloudconvert' to install CloudConvert third party API
4. type 'pip install python-docx' to install python-docx library
5. tkinter library should be embedded within Python
--------------------------------------------------------------------------

Download
--------------------------------------------------------------------------
1. open the URL: https://nlp.stanford.edu/software/stanford-ner-2018-02-27.zip for downloading the model of stanfordNLP
2. Decompress the 'stanford-ner-2018-02-27.zip'
3. open 'Named Entities Anonymisation for Academic Journal Articles.py'
4. change the value of 'stner' at the beginning.
example: if you download the zip and decompress on your desktop, you need to change the value of 'stner' to 

stner = StanfordNERTagger(model_filename=r'/xxx/xxx/Desktop/stanford-ner-2018-02-27/classifiers/english.conll.4class.distsim.crf.ser.gz',path_to_jar=r'/xxx/xxx/Desktop/stanford-ner-2018-02-27/stanford-ner.jar')

5. save the code file
--------------------------------------------------------------------------

Run
--------------------------------------------------------------------------
1. run the code file
2. a GUI will jump out for users
3. The API key can be changed through the textfile, and users can register their own accounts on https://cloudconvert.com/. After registering an account, you can check your own unique AIP key in your user profile.
4. Click the 'Choose file' to choose the input PDF file. The path of it will be shown at the interface.
5. Click 'RUN' to start processing the program
6. The result will be shown in an independent window.
7. If you want to anonymise other content, you can directly type the text in the textfiled at the top of 'ANONYMISE' button. Then click 'ANONYMISE' to process.
8. Click 'GENERATE' to generate the output PDF file.
--------------------------------------------------------------------------


