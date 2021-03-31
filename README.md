# WordCat - WORK IN PROGRESS
A simple php library for manipulation of docx word processed document; in particular the library is designed to allow content from one document to be inserted into another document, and it also provides some features allowing for searching and replacing content, and inserting new text and images.

# Dependencies
This library requires `DOMDocument` and `SimpleXML`.

# Features
This library was designed to solve some fairly specific issues I was having with (mentioning no names) another popular PHP library when creating new documents from a given template; As such there are a **lot** of features that simply aren't implemented as I have had no need for them. I have added a "limitations" section below to list particularly glaring features which you may want but are missing, but if a feature you want isn't in the bullet list below, take it as read that the library won't do it for you!

* Open a docx document
* Read files from the document
* Write files to the document
* Manipulate XML files within the document
* Save docx, either overwriting the original file or as a new document
* Search for text, either as a binary/plain text search or regular expression
* Replace text, either as a binary/plain text search or regular expression
* Insert paragraphs containing plain text
* Insert image files with a specified size (aspect ratio is preserved while constraining the image within the specified size)

# Limitations
As this library was designed to solver some fairly specific issues, it will not do a lot of things you may require; I may or may not come to add features

* You can't create a new, blank document
* You can't insert rich text, or have any complex settings for images
* Images are inserting using an old format; this still works, but it's not compliant with the latest standards

