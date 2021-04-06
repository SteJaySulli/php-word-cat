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
* Insert the contents of one docx into another

# Limitations
As this library was designed to solver some fairly specific issues, it will not do a lot of things you may require; I may or may not come to add features

* You cannot create a new document from scratch; you need to load an existing docx file to work on
* The library is naive; most operations require some knowledge of the internal XML structure of docx files

# Installation

You can install this library using composer:

```bash
composer install stejaysulli/php-word-cat
```

It should be possible to use the library without composer, but this has not been tested and you will need to provide your own method to autoload the files in the `src` directory.

# Usage

A nice example of all the basic features are available in [test/example.php](./test/example.php) - It is adviseable to check that out for proper usage details. Just so you can get a feel for the kind of code you'll be writing with WordCat, here's a brief example though:

```php
<?php
use WordCat\WordCat;

// Open a docx file for processing:
$wordcat = WordCat::instance("/path/to/document.docx");

// Case sensitive search and replace:
$wordcat->clearSearch()->replaceText("wordcat", "WordCat");

// Replace text using a regular expression:
$wordcat->clearSearch()->replaceText("integer", "integer (" . random_int(1,100) . ")");

// Add a new random number after all occurences of the word "random":
$wordcat->clearSearch()->replaceText("random", function($found) {
    return "$found (" . random_int(0,100) . ")";
});

// Use a regex to replace anything that's formatted like a Y-m-d date with
// today's date:
$wordcat->clearSearch()->replaceText("/[0-9]{4}-[0-9]{2}-[0-9]{2}/", date("Y-m-d"), true);

// $wordcat->clearSearch()->replaceRegex("/[0-9]{4}-[0-9]{2}-[0-9]{2}/", date("Y-m-d"));

// // Use a regex to make a more complex replacement:
// $wordcat->clearSearch()->replaceRegex("/i (.+ )?like (.+) a lot/i", 'I $1love $2 completely');

// // Use a regex to make an even more complex replacement; here we use a
// // callback to change only some matches (at random):
// $wordcat->clearSearch()->replaceRegex("/shuffle ([a-z0-9]+)/i", function($found) {
//     if(random_int(0,99) < 50) {
//         return "shuffled " . str_shuffle(substr($found, 8));
//     } else {
//         return 'unshuffled $1';
//     }
// });

// Inserting new content
///////////////////////////////////////////////////////////////////////////////
echo "Inserting new content\n";

/*
    WordCat provides some basic functions to insert content into the document.
    This is currently limited to the insertion of plain text (with no styling)
    and images which fit within a set size limit.

    When inserting images, you specify the size in points - The image will be
    sized within these dimensions while maintaining the aspect ratio of the
    image.

    Note that images should be inserted into a run within a paragraph; this
    isn't compulsory, but it is recommended to avoid unwanted effects such as
    text flowing around the image in odd ways. To do this in the example
    below, we use insertRun() to insert a new run to the end of a paragraph.
*/

// Insert some paragraphs after each paragraph containing some text:
$wordcat->findText("inserted paragraph next")
    ->forSearch(
        function($insertionPoint) use($wordcat, $imageFile) {
            // Insert a simple paragraph containing one text run:
            $p1 = $wordcat->insertParagraph(
                "This is the first inserted paragraph",
                $insertionPoint
            );
            
            // Insert another paragraph after the one we just inserted:
            $p2 = $wordcat->insertParagraph(
                [
                    "This is the second inserted paragraph.",
                    "Note that we can include several text runs",
                    "by specifying an array."
                ],
                $p1
            );
            
            // Insert an image into the last paragraph inside a new run:
            $wordcat->insertImage(
                $imageFile, // The image file we want to insert
                $wordcat->insertRun($p2), // Create a new insertion point
                600, 600, // The width and height constraints for the image
                "Inserted image" // The title of the image
            );

        }
    );

// Document Concatenation/Insertion
///////////////////////////////////////////////////////////////////////////////
echo "Inserting one document into another\n";

/*
    The primary function this library needed to provide was the ability to
    concatenate two docx documents together. 

    This is a fairly simple implementation but the task is more complex than
    it may seem; docx files have internal IDs to assets (images in particular)
    which need to be translated to avoid conflicts between the source document
    IDs and the ones that already exist within the destination document.

    WordCat will try to take care of this for you, so you can insert some or
    all content from one document into another. The styles are not guaranteed
    to be preserved, but the content itself (including images) should be.

    Note that the source document will be altered during the process (for
    example all of the IDs that conflict with our destination document will be
    changed) - It is assumed that the source document is a temporary document
    that will not be saved, so changes can be made in this before inserting it
    into the destination document. Basically DO NOT SAVE YOUR SOURCE DOCUMENT
    AFTER INSERTING IT INTO ANOTHER DOCUMENT!
*/

// Get an element to insert the document after:
$search = $wordcat->findText("INSERT DOCUMENT HERE")->getSearch();

// Check we have at least one insertion point:
if(count($search) > 0) {
    // Use the first insertion point:
    $wordcat->insertDocument($wordcatSource, $search[0]);
    // Remove the insertion point element to get rid of the "INSERT DOCUMENT
    // HERE" text:
    $wordcat->removeNode($search[0]);
}

// Saving & Closing
///////////////////////////////////////////////////////////////////////////////
echo "Saving and closing files\n";

/*
    There are two save functions; You probably do NOT want save() but instead
    want to use saveAs so you can save the results of your changes to a new
    file.

    Once you are done with a WordCat document, you should always close it.
    This ensures that the temporary files are cleaned up. While WordCat does
    attempt to do this with a destructor, destructors in PHP seem to be
    unreliable, so I recommend you always do this manually!

    Once a WordCat object is closed, you should discard it; any attempts to
    use the object after this point will result in an error.
*/

// Save a new docx file
$wordcat->saveAs($outputFile);

// Close the WordCat instance
$wordcat->close();
$wordcatSource->close();

echo "Example script completed.\n";
```