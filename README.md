# Word_Doc_Generator
Students in engineering are required to create a docx file for all the practicals they perform in labs. The most annoying and tedious task is to format the data. To save time, I created a Python program to generate a docx file with predefined formatting.

# Instructions and Notes
1. The program is completely GUI based so no need to worry about running it from CLI
2. You can add only one picture as of now. To add more pictures you will need to add them manually by editing the document
3. The picture added will be of size 3.5 inches x 2 inches so most probably you will need to resize it after the doc file is created
4. This program can now also run without need of installing Python.Download the entire repository and then check for the exe file in '/Practical-File-Generator-Windows/dist' folder

# Libraries required to run this program
1. python-docx - pip install python-docx
2. tkinter - pip install tk

# Predefined Styles in the Program
1. Heading - Times New Roman, 16, Bold
2. Sub-Heading - Times New Roman, 14, Bold
3. Content - Times New Roman, 12
4. Header and Footer - Times New Roman, 12, Bold

# Sections Generated in Word File
1. Header containing ID Number and Subject Name and Code 
2. Heading
3. Aim
4. Program Code
5. Output
6. Conclusion

# Documentation Links
1. docx : https://python-docx.readthedocs.io/en/latest/
2. tkinter : https://docs.python.org/2/library/tkinter.html
