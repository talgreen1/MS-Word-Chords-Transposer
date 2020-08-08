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
* You need to enable developer support & add a new module with this script.
* Then - you need to copy this code to the 


https://www.datanumen.com/blogs/how-to-run-vba-code-in-your-word/#:~:text=Firstly%2C%20click%20%E2%80%9CVisual%20Basic%E2%80%9D,to%20open%20a%20new%20module. 

