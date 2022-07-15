# WordCat - Limited manipulation of docx word processed documents
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
composer require stejaysulli/php-word-cat
```

It should be possible to use the library without composer, but this has not been tested and you will need to provide your own method to autoload the files in the `src` directory.

# Usage

A nice example of all the basic features are available in [tests/example.php](./tests/example.php) - It is adviseable to check that out for proper usage details. Just so you can get a feel for the kind of code you'll be writing with WordCat, here's a brief example though:

```php
<?php
use WordCat\WordCat;

$wordcat = WordCat::instance("example1.doc");

// Find XML elements containing the text "apple", "orange" or "pear"
$wordcat->findText("apple")
        ->andFindText("orange")
        ->andFindText("pear")
        ->forSearch(function($element) {
            // Print the text content of each element:
            echo "{$element->textContent}\n";
        });

// Get an array of all the matching XML elements (DOMNode objects):
$results = $wordcat->getSearch();

// Clear the search results:
$wordcat->clearSearch();

// Do the same, but use a regex this time, and make it case insensitive:
$wordcat->findRegex("/(apple|orange|pear)/i")
        ->forSearch(function($element) {
            echo "{$element->textContent}\n";
        });

// Replace all instances of the word "wordcat" with "WordCat":
$wordcat->clearSearch()->replaceText("wordcat", "WordCat");

// Use a regex to replace anything that's formatted like a Y-m-d date with
// today's date:
$wordcat->clearSearch()->replaceRegex("/[0-9]{4}-[0-9]{2}-[0-9]{2}/", date("Y-m-d"));

// Insert some paragraphs after each paragraph containing some text:
$wordcat->findText("inserted paragraph next")
    ->forSearch(
        function($insertionPoint) use($wordcat, $imageFile) {
            // Insert a simple paragraph containing one text run:
            $p1 = $wordcat->insertParagraph(
                "This is the first inserted paragraph",
                $insertionPoint
            );
        }
    );


// Get an element to insert the document after:
$search = $wordcat->findText("INSERT DOCUMENT HERE")->getSearch();

// Open a document to insert
$wordcatSource = WordCat::instance("example2.docx");

// Check we have at least one insertion point:
if(count($search) > 0) {
    // Use the first insertion point:
    $wordcat->insertDocument($wordcatSource, $search[0]);
    // Remove the insertion point element to get rid of the "INSERT DOCUMENT
    // HERE" text:
    $wordcat->removeNode($search[0]);
}

// Save a new docx file
$wordcat->saveAs("test-output.docx");

// Close the WordCat instance
$wordcat->close();
$wordcatSource->close();

```
