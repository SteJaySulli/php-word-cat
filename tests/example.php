<?php
require_once __DIR__ . '/../vendor/autoload.php';
use WordCat\WordCat;

// Set the filenames for the files we'll use in the examples below:
$mainDocumentFile = __DIR__ . "/example1.docx";
$sourceDocumentFile = __DIR__ . "/example2.docx";
$imageFile = __DIR__ . "/tulips.png";
$outputFile = __DIR__ . "/../test-output-example.docx";

// Load a docx file:
///////////////////////////////////////////////////////////////////////////////
echo "Open WordCat documents:\n";

/* 
    You should always instantiate WordCat with an existing docx file to load;
    this can be done using the "instance" helper, or by instantiating a WordCat
    object yourself, as shown below.

    Behind the scenes, WordCat will make a copy of the docx file to a temporary
    file which it will keep open while you are working on the document and 
    until you close the WordCat instance. It's worth keeping this in mind; 
    Although WordCat's destructor is supposed to take care of deleting the
    temporary file if you forget to, testing has shown that this doesn't always
    happen!
*/

$wordcat = WordCat::instance($mainDocumentFile);
$wordcatSource = new WordCat($sourceDocumentFile);

// Find Text
///////////////////////////////////////////////////////////////////////////////
echo "Find Text:\n\n";

/*
    WordCat provides a facility to find xml elements which contain text that
    matches a given string or regular expression pattern. To make this as
    flexible as possible, searches can be chained together using various
    functions.

    There are a few variations on findText; here we'll show "findText",
    "andFindText" and "findRegex".

    In addition we'll show the handy "forSearch" method which allows you to
    iterate over the search results and operate on each element found.
*/

// Find a list of elements which contain the words "apple", "orange" and 
// "pear", and print the full text of the elements, each to a single line
$wordcat->findText("apple")
        ->andFindText("orange")
        ->andFindText("pear")
        ->forSearch(function($element) {
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



// Find and replace text
///////////////////////////////////////////////////////////////////////////////
echo "\nFind and replace text\n";

/*
    This is a very simple and essential tool for using a docx as a template for
    a new file. WordCat makes this easy, but there are a few gotchas which I
    will list in a section below these examples.

    The find/replace functions augment the findText methods (shown below); when
    doing a find/replace without it being related to a search you've previously
    done, it's good idea to clear the search first to ensure the find/replace
    operation is not limited to whatever you last searched - this includes the
    last find/replace operation you did!

    The replacement text can be a string, or a callback to provide more
    complex functionality.
*/

// Replace all instances of the word "wordcat" with "WordCat":
$wordcat->clearSearch()->replaceText("wordcat", "WordCat");

// Add the same random number after all occurences of the word "integer":
$wordcat->clearSearch()->replaceText("integer", "integer (" . random_int(1,100) . ")");

// Add a new random number after all occurences of the word "random":
$wordcat->clearSearch()->replaceText("random", function($found) {
    return "$found (" . random_int(0,100) . ")";
});

// Use a regex to replace anything that's formatted like a Y-m-d date with
// today's date:
$wordcat->clearSearch()->replaceRegex("/[0-9]{4}-[0-9]{2}-[0-9]{2}/", date("Y-m-d"));

// Use a regex to make a more complex replacement:
$wordcat->clearSearch()->replaceRegex("/i (.+ )?like (.+) a lot/i", 'I $1love $2 completely');

// Use a regex to make an even more complex replacement; here we use a
// callback to change only some matches (at random):
$wordcat->clearSearch()->replaceRegex("/shuffle ([a-z0-9]+)/i", function($found) {
    if(random_int(0,99) < 50) {
        // This shuffles the actual text found after "shuffle ":
        return "shuffled " . str_shuffle(substr($found, 8));
    } else {
        // This doesn't shuffle, so we use a regex reference:
        return 'unshuffled $1';
    }
});

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