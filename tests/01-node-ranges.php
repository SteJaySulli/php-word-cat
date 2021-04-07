<?php
require_once __DIR__ . '/../vendor/autoload.php';
use WordCat\WordCat;

/*
    Here we test some features provided to deal with ranges of nodes.

    We can tell if a node is before or after a node, and we can get a node
    with a given tagname after or before a given node.
 */

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
    echo $paragraphs[$i]->textContent . " is next after " . $node->textContent . "\n";
}
echo "\n";

// Now we perform a search; return all nodes from Paragraph 3 through Paragraph 7:
echo "Find nodes between paragraph 3 and 7:\n";
$wordcat->clearSearch()
    ->findTagName("p")
    ->searchFrom($paragraphs[2])
    ->searchTo($paragraphs[6])
    ->forSearch( function( $node ) {
        echo "{$node->textContent}\n";
    });
echo "\n";

// We can also use searchFrom and searchTo to get all nodes which are within a given
// node using a tagName search for "*"
echo "The xml tags within paragraph 5:\n";
$wordcat->clearSearch()
    ->findTagName("*")
    ->searchFrom($paragraphs[4])
    ->searchTo($paragraphs[5])
    ->forSearch( function( $node ) {
        echo "<{$node->nodeName}>{$node->textContent}</{$node->nodeName}>\n";
    });
echo "\n";

// We can get a similar effect by searching only inside a specified node:
echo "The xml tags within paragraph 6:\n";
$wordcat->clearSearch()
    ->searchInside($paragraphs[5])
    ->findTagName("*")
    ->forSearch( function( $node ) {
        echo "<{$node->nodeName}>{$node->textContent}</{$node->nodeName}>\n";
    });
echo "\n";

// We'll extract a single w:t node from paragraph 6 for use in the next tests
$textNode = $wordcat->clearSearch()
    ->searchInside($paragraphs[5])
    ->findTagName("t")
    ->getSearch()[0];

// We can tell if a given node is a descendant of another given node:
if($wordcat->nodeIsDescendant($textNode, $paragraphs[4])) {
    echo "Text node from paragraph 6 is within paragraph 5\n";
} else {
    echo "Text node from paragraph 6 is not within paragraph 5\n";
}
if($wordcat->nodeIsDescendant($textNode, $paragraphs[5])) {
    echo "Text node from paragraph 6 is within paragraph 6\n";
} else {
    echo "Text node from paragraph 6 is not within paragraph 6\n";
}
echo "\n";

// Likewise we can tell if a given node contains another node:
if($wordcat->nodeContainsNode($paragraphs[4], $textNode)) {
    echo "Text node from paragraph 6 is within paragraph 5\n";
} else {
    echo "Text node from paragraph 6 is not within paragraph 5\n";
}
if($wordcat->nodeContainsNode($paragraphs[5], $textNode)) {
    echo "Text node from paragraph 6 is within paragraph 6\n";
} else {
    echo "Text node from paragraph 6 is not within paragraph 6\n";
}
echo "\n";

// As text replace functions operate on a search, we can do replacement only
// within a specified block of nodes:

$wordcat->clearSearch()
    ->searchFrom($paragraphs[2])
    ->searchTo($paragraphs[7])
    ->replaceRegex('/Paragraph ([0-9])/', 'Paragraph Number $1');

// We can do an even more selective search and replace by doing a search manually first:
$wordcat->clearSearch()
    ->findText('Paragraph Number 3')
    ->andFindText('Paragraph Number 5')
    ->andFindText('Paragraph Number 7')
    ->replaceRegex('/Paragraph Number ([0-9])/', 'Paragraph No. $1');

// Lets check the results of the last test by printing all paragraphs' text:
$wordcat->clearSearch()
    ->findText("Paragraph")
    ->forSearch(function( $node ) {
        echo "{$node->textContent}\n";
    });
echo "\n";

// We can find nodes from attributes too:
$wordcat->clearSearch()
    ->findAttribute("w:val")
    ->searchTo($paragraphs[2])
    ->forSearch(function( $node ) use ($wordcat) {
        echo $wordcat->getNodeXML($node);
        if($p = $wordcat->getClosestTagName("w:p", $node)) {
            echo "Within {$p->textContent}\n";
        } else {
            echo "Not within paragraph\n";
        }
    });
echo "\n";

// We can even find nodes with specific attributes:
    $wordcat->clearSearch()
    ->findAttribute("w:val", "Normal")
    ->forSearch(function( $node ) use ($wordcat) {
        echo $wordcat->getNodeXML($node);
        if($p = $wordcat->getClosestTagName("w:p", $node)) {
            echo " Within {$p->textContent}\n";
        } else {
            echo " Not within paragraph\n";
        }
    });
echo "\n";

// We can get top level parent of a given node. Lets get the top level parent of our
// first paragraph: this should be the w:body node:
$body = $wordcat->getTopLevelParent($paragraphs[1]);
echo "Top level node: {$body->nodeName}\n";

// Lets check programmatically to see if our body is actually the one and only body:
foreach($wordcat->getNodesByTagName("body") as $node) {
    if($wordcat->nodeIs($node, $body)) {
        echo "Body node matches!\n";
    }
}

// We can also specify what the top level is when getting a top level parent; lets
// rebuild $paragraphs, adding each paragraph twice:
$paragraphs = [];
$wordcat->clearSearch()
    ->findText("Paragraph")
    ->andFindText("Paragraph")
    ->forSearch( function($node) use(&$paragraphs, $wordcat, $body) {
        $paragraphs[] = $wordcat->getTopLevelParent($node, $body);
    });
// Now we'll print the paragraphs to ensure we have them; we should see all paragraphs twice:
foreach($paragraphs as $p) {
    echo "{$p->textContent}\n";
}
echo "\n";

// We can remove duplicates in search results; note we are using the same search results here,
// and we rebuild the list by using searchLambda:
echo "Rebuild paragraph list using searchLambda and searchUnique\n";
$paragraphs = $wordcat->clearSearch()
    ->findText("Paragraph")
    ->andFindText("Paragraph")
    ->searchLambda( function($node) use(&$paragraphs, $wordcat, $body) {
        return $wordcat->getTopLevelParent($node, $body);
    })
    ->searchUnique()
    ->getSearch();
// Now we'll print the paragraphs to ensure we have them; we should see all paragraphs twice:
foreach($paragraphs as $p) {
    echo "{$p->textContent}\n";
}
echo "\n";


$wordcat->close();