# downloadPhotosWithExcel
Basic code to download photos from a web/portal with Excel using a referential link, the photo must be on the web (own sequential style).

Execution Test:

![Imgur](https://i.imgur.com/060vIHr.gif)

* Prerequisites (only for windows 7 - 10):

Office 2012-2019 (32/64 bit)


Instructions:

* Open or create a macro-enabled Excel file.
* Create table of contents as shown below.

![Imgur1](https://i.imgur.com/gvOJ8Ov.png)


* Import the code as a module to the Excel file (.xlsm).

* In the code references add the following marked:

![Imgur2](https://i.imgur.com/YXZphpC.png)

* Select the codes in the procedure column and execute the macro: "DOWNLOAD_PHOTOS_SELECTION"

![Imgur3](https://i.imgur.com/060vIHr.gif)

Important note: When the macro is executed, it creates a predefined folder tree as follows:
  - In the current location of the file create the folder "DOWNLOAD"
  - Next, within this, create another folder with the name of the spreadsheet where the selection for downloading photos was made.
  - Finally, in this folder, individually create folders with photos from the "Process" column of the previously created table.
  - The images are created with the name of the header where they are found in the table and are assigned to the folder that was created with the process number/code.

