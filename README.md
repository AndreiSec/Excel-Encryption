# Excel-Encryption
Cipher-table encryption application made in VBA.
What it does:
My application is a cipher text excel file encryption program. It uses a cipher table found on the second sheet of the file to encrypt every single character in an excel file. It then outputs each sheet from the input excel file into a separate .crzy file named after the worksheet. The program works best with ASCII value data in the input cells. For example, it will not encrypt formulas, but instead take the data value of that cell and encrypt it.

The steps that it takes are as follows:
Encryption:
1.	Program will ask user to give the location of the excel file of which to encrypt each worksheet.
2.	Program will ask user to input the desired directory in which to output the .crzy files to.
3.	Sub will loop through each sheet, convert each line to an array.
4.	Each cell in that array will then be encrypted and written into a .crzy file.
a.	Each character in the original array will be encrypted according to a VLOOKUP encryption algorithm found in the applications second sheet, titled “Encryption Table”.
b.	It will be separate by “|” and each line will represent an excel row.
5.	This new encrypted string will be outputted to a file that matches the name of the sheet in the directory specified.
a.	Write to database the file being encrypted, the sheet name, number of characters encrypted, the translation factor, the date of the encryption, and the output file directory.
b.	Append to word file a chart containing how many cells were processed in each cell.

Decryption:
1.	Program will ask user where they would like the outputted .xlsx file to be, and the name of the file
2.	Program will ask users how many .crzy files they want to decrypt
3.	Program will ask users to give full file locations of the .crzy files
4.	Sub will loop through each line of the .crzy file and decipher each cell / character according to a table titled “Decryption Table” in the second sheet. It will use VLOOKUP to find the corresponding deciphered characters.
5.	Will then paste each deciphered cell according to where it should be in the worksheet. Remember that each line in the .crzy file is a row in the worksheet, and each “|” represents a column separator in the worksheet.
6.	Program will then save outputted .xlsx file where the user dictated
a.	Write to database the file being encrypted, the sheet name, number of characters encrypted, the translation factor, the date of the encryption, and the output file directory.
b.	Append to word file a chart containing how many cells were processed in each cell.
Why I chose it:
I’ve been interested in encryption ever since I was 12 years old, when I used OCLHashcat to encrypt and crack MD5 Checksum passwords. A career in cyber security was always on my mind and I thought of this as not only a great way to develop my skills in both VBA and Encryption, but also to add something to my resume when applying for co-op jobs this winter. I had a lot of fun coding this program and spent 2 and a half full days working on it. I wanted a program that was a challenge to write, yet useful, and this fit the bill.


Andrei Secara's Cipher Encryption & Decryption Application MANUAL

Thank you for deciding to encrypt your file with my cipher encryption program! To get started, you must closely follow the following steps before running the app:
------------------------------------------------------------------------------------------------------------------------------------------
For this application to run correctly, you must enable a reference to the VB script run-time, Activex, and word library. This is so VBA can use the File System Object and access databases & word files.  
This can be done in the following steps:
1. Open a5_seca2560.xlsm
2. Open the VBA Editor
3. Select Tools > References from the drop-down menu
4. A listbox of available references will be displayed
5. Tick the check-box next to 'Microsoft Scripting Runtime'
6. Tick the check-box next to 'Microsoft ActiveX Data Objects Library' Select the latest version installed on your PC
7. Tick the check-box next to 'Microsoft Word xx.x Object Library'. Select the latest version.
8. Click on the OK button
------------------------------------------------------------------------------------------------------------------------------------------
Great! Now your computer is all set up to run the encryption application.
------------------------------------------------------------------------------------------------------------------------------------------
The steps to use the program to ENCRYPT files are as follows:
1. On the 'Cipher Tables' worksheet, choose an appropriate translation factor. This will determine how the files are encrypted. Be aware! A translation factor of 0 will result in no meaningful encryption.
2. Press the lock above "ENCRYPT FILE"
3. Enter the exact directory of the file you wish to encrypt
4. Enter the directory (folder location) where you would like the program to output the resulting .crzy files
5. Done! Your excel file is now encrypted and outputted where you specified.
A sample document titled "sample.xlsx" has been given to you to test out the encryption.
The application will then encrypt your file to the resulting files with the same name as your workbook's worksheets.
------------------------------------------------------------------------------------------------------------------------------------------
The steps to use the program to DECRYPT files are as follows:
1. Make sure the translation factor on the 'Cipher Tables' worksheet is the same as when the files were encrypted. If it is not, the files will not decrypt properly.
2. Press the lock above "DECRYPT FILE"
3. Enter the exact directory (folder location) of where you would like your decrypted .xlsx file.
4. Enter the desired output file name. EG: "output.xlsx"
5. Enter how many .crzy files you would like to decrypt into the outputted workbook.
6. Enter the full file directories of all the .crzy files you wish to decrypt.
7. Done! Your decrypted file will be where you specified.


After this, you can check the Output.accdb for a log, and the Chart.docx file for a visual pie chart representation of how many cells were processed by Sheet name in your input / output workbook.
