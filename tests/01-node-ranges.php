<?php
require_once __DIR__ . '/../vendor/autoload.php';
use WordCat\WordCat;

// Set the filenames for the files we'll use in the examples below:
$mainDocumentFile = __DIR__ . "/node-ranges.docx";
$wordcat = WordCat::instance($mainDocumentFile);

$paragraphs = [];
// Get a list of paragraphs by finding the containing w:p for each text match (which will be a w:t)
$wordcat->findText("Paragraph")->forSearch( function($node) use(&$paragraphs, $wordcat) {
    $paragraphs[] = $wordcat->getClosestTagName("w:p", $node);
});

// Test nodeIsBefore
if($wordcat->nodeIsBefore($paragraphs[0], $paragraphs[1])) {
    echo "Paragraph 1 is before Paragraph 2\n";
} else {
    echo "Paragraph 1 is not before Paragraph 2\n";
}
if($wordcat->nodeIsBefore($paragraphs[0], $paragraphs[4])) {
    echo "Paragraph 1 is before Paragraph 5\n";
} else {
    echo "Paragraph 1 is not before Paragraph 5\n";
}
if($wordcat->nodeIsBefore($paragraphs[3], $paragraphs[2])) {
    echo "Paragraph 4 is before Paragraph 3\n";
} else {
    echo "Paragraph 4 is not before Paragraph 3\n";
}
if($wordcat->nodeIsBefore($paragraphs[8], $paragraphs[1])) {
    echo "Paragraph 9 is before Paragraph 2\n";
} else {
    echo "Paragraph 9 is not before Paragraph 2\n";
}

// Test nodeIsAfter
if($wordcat->nodeIsAfter($paragraphs[0], $paragraphs[1])) {
    echo "Paragraph 1 is after Paragraph 2\n";
} else {
    echo "Paragraph 1 is not after Paragraph 2\n";
}
if($wordcat->nodeIsAfter($paragraphs[0], $paragraphs[4])) {
    echo "Paragraph 1 is after Paragraph 5\n";
} else {
    echo "Paragraph 1 is not after Paragraph 5\n";
}
if($wordcat->nodeIsAfter($paragraphs[3], $paragraphs[2])) {
    echo "Paragraph 4 is after Paragraph 3\n";
} else {
    echo "Paragraph 4 is not after Paragraph 3\n";
}
if($wordcat->nodeIsAfter($paragraphs[8], $paragraphs[1])) {
    echo "Paragraph 9 is after Paragraph 2\n";
} else {
    echo "Paragraph 9 is not after Paragraph 2\n";
}
echo "\n";
// Test getNextTagName
for($i=0; $i < 8; $i++) {
    $node = $wordcat->getNextTagName("w:p", $paragraphs[$i]);
    echo $paragraphs[$i]->textContent . " is followed by " . $node->textContent . "\n";
}

echo "\n";

// Test getPreviousTagName
for($i=1; $i < 9; $i++) {
    $node = $wordcat->getPreviousTagName("w:p", $paragraphs[$i]);
    echo $paragraphs[$i]->textContent . " is preceded by " . $node->textContent . "\n";
}
echo "\n";

$wordcat->close();