# MS-Word-Chords-Transposer
A VBA script that can transpose chords in a MS Word document

This script can transpose chords in a MS Word document. It scans the entire text in the document, detects only 
the chords and transpose them, up or down based on your input.

The chord detection goes as follows:
1. You need to have a unique font for your chords. Only the chords should use this font. 
Your lyrics or other text should use another font.
2. When you run this script - it will scan the entire text and will look for **Uppercase** letters with that font.
Only these letters will be transposed.
3. For example: If you have the chord Cm7, only the uppercase 'C' will be changed.
4. You should put the name of the chord's font in the second line of the script:
`Public Const FONT_NAME = "Courier New" '<---Put the name of your chord font here`

How to install and run the script:
* The script is a standard MS Word VBA code.
* You need to enable developer support & add a new module with this script. Check this [link](https://www.datanumen.com/blogs/how-to-run-vba-code-in-your-word/#:%7E:text=Firstly%2C%20click%20%E2%80%9CVisual%20Basic%E2%80%9D,to%20open%20a%20new%20module)
* After creating a new module, put the code found in the file '**MS-Word-Chords-Transposer-vba-script.txt**' inside the module
* Then - you need to copy this code to the '**Normal.dotm**' file: Check the section '**Make a macro available in all documents**' in this [link](https://support.microsoft.com/en-us/office/create-or-run-a-macro-c6b99036-905c-49a6-818a-dfb98b7c3c9c)
* In order to execute the scrip - you need to run the '**transpose**' macro: Check the section '**Run a macro**' in this [link](https://support.microsoft.com/en-us/office/create-or-run-a-macro-c6b99036-905c-49a6-818a-dfb98b7c3c9c)
* You can create a button that execute this script. Check this [link](https://support.microsoft.com/en-us/office/assign-a-macro-to-a-button-728c83ec-61d0-40bd-b6ba-927f84eb5d2c#:~:text=Click%20File%20%3E%20Options%20%3E%20Quick%20Access,on%20the%20Quick%20Access%20Toolbar.)

