<?php
namespace WordCat;
use DOMDocument;
use DOMElement;
use DOMNode;
use DOMNodeList;
use DOMText;
use DOMXPath;

/**
 * WordCatXML Class
 * 
 * This class provides an abstracted interface to manipulate individual XML files within
 * a docx document.
 * 
 * These are instantiated by the WordCat class instance, and contain a link back to this
 * instance in order to function - in particular saving changes needs a WordCat object to
 * work.
 */
class WordCatXML {
    private $docPath = null;
    private $wordCat = null;
    private $document = null;
    private $searchResults = [];

    /**
     * Constructor - Takes the owning WordCat object and the XML file path within
     * the archive as parameters. This will then load the XML file, allowing us to
     * operate on the file using this object.
     *
     * @param WordCat $instance
     * @param string $docPath
     */
    function __construct(WordCat &$instance, string $docPath) {
        $this->docPath = $docPath;
        $this->wordCat = &$instance;
        $this->document = $instance->readXML($docPath);
    }

    /**
     * Return the underlying DOMDocument object for the XML
     *
     * @return DOMDocument
     */
    function getDOMDocument() {
        return $this->document;
    }

    /**
     * Return the XML file path within the archive
     *
     * @return string
     */
    function getPath() {
        return $this->docPath;
    }

    /**
     * Store the XML file to the archive. This is named "store" to disambiguate from
     * the WordCat::save function.
     *
     * @return void
     */
    function store() {
        $this->wordCat->writeFile($this->docPath, $this->document->saveXML());
    }

    /**
     * Clone a given node within the same document.
     *
     * @param DOMNode $node
     * @return DOMNode
     */
    function cloneNode($node) {
        return $node->cloneNode(true);
    }

    /**
     * Import a given node (from another document) into the current one.
     *
     * @param DOMNode $node
     * @return DOMNode
     */
    function importNode($node) {
        return $this->document->importNode($node, true);
    }

    /**
     * Import a given node (from another document) and append it to the current
     * document directly after the given "sibling" node.
     *
     * @param DOMNode $node
     * @param DOMNode $sibling
     * @return DOMNode
     */
    function importNodeAfter($node, $sibling) {

        $newNode = $this->importNode($node);
        if($sibling->nextSibling) {
            $sibling->parentNode->insertBefore($newNode, $sibling->nextSibling);
        } else {
            $sibling->parentNode->appendChild($newNode);
        }
        return $newNode;
    }

    /**
     * Import a given node (from another document) and append it to the current
     * document directly before the given "sibling" node.
     *
     * @param DOMNode $node
     * @param DOMNode $sibling
     * @return DOMNode
     */
    function importNodeBefore(&$node, &$sibling) {
        $newNode = $this->importNode($node);
        $sibling->parentNode->insertBefore($newNode, $sibling);
        return $newNode;
    }

    /**
     * Import a given node (from another document) and append it to the current
     * document as the last child of the given "parent" node
     *
     * @param DOMNode $node
     * @param DOMNode $parent
     * @return DOMNode
     */
    function importNodeInside(&$node, &$parent) {
        $newNode = $this->importNode($node);
        $parent->appendChild($newNode);
        return $newNode;
    }

    function createNode($tagname, $attrs=[], $text=null) {
        $node = $this->document->createElement($tagname, $text);
        if(is_array($attrs)) {
            foreach($attrs as $key=>$value) {
                $node->setAttribute($key, $value);
            }
        }
        return $node;
    }

    function createNodeTree($tree, $target, $callback) {
        foreach($tree as $element) {
            $attrs = null;
            $text = null;
            if(isset($element["attributes"])) {
                $attrs = $element["attributes"];
            }
            if(isset($element["text"])) {
                $text = $element["text"];
            }
            $node = $this->createNode($element["tag"], $attrs, $text);
            $callback($node, $target);

            if(isset($element["children"])) {
                $this->createNodeTree($element["children"], $node, function($a, $b) { $this->insertNodeInside($a,$b); });
            }

        }
        return $node;
    }

    function insertNodeAfter($node, $sibling) {
        if($sibling->nextSibling) {
            return $sibling->parentNode->insertBefore($node, $sibling->nextSibling);
        } else {
            return $sibling->parentNode->appendChild($node);
        }
    }

    function insertNodeBefore($node, $sibling) {
        return $sibling->parentNode->insertBefore($node, $sibling);
    }

    function insertNodeInside($node, $parent) {
        return $parent->appendChild($node);
    }

    /**
     * Duplicates the given node within the document. This makes an identical copy of
     * the node and inserts it directly before the original node. The original node will
     * be returned. This approach means we can continue to duplicate the node as many times
     * as needed as cheaply as possible.
     *
     * You can specify a $count, in which case $count duplicates will be created.
     *
     * @param DOMNode $node
     * @param integer $count
     * @return DOMNode
     */
    function duplicateNode($node, $count=1) {
        $parent = $node->parentNode;
        for($i=0;$i<$count;$i++) {
            $parent->insertBefore($this->cloneNode($node), $node);
        }
        return $node;
    }

    /**
     * Remove a given node from the XML tree. This does not work for the root node!
     *
     * @param DOMNode $node
     * @return void
     */
    function removeNode($node) {
        $node->parentNode->removeChild($node);
    }

    /**
     * Return an XML string for the given node (and its children). This is mostly
     * useful for debugging, but it could find other applications.
     *
     * @param DOMNode $node
     * @return string
     */
    function getNodeXML($node) {
        return simplexml_import_dom($node)->asXML();
    }

    /**
     * Return an XML string for the entire document
     *
     * @return void
     */
    function getDocumentXML() {
        return $this->document->saveXML();
    }

    /**
     * Internal function to store the internal search results. This is used to optionally
     * append results to the results, or to replace the results with a new search
     *
     * @param array $elements
     * @param bool $append - if false, the current search results will be cleared
     * @return void
     */
    private function setSearch(array $elements, bool $append = false) {
        $this->searchResults = $append ? array_merge($this->searchResults, $elements) : $elements;
    }

    /**
     * Clear the internal search results. This can be chained with other search functions
     * in order to compile complicated searches.
     *
     * @return WordCatXML
     */
    function clearSearch() {
        $this->searchResults = [];
        return $this;
    }

    /**
     * Text search function, which searches the whole document for the given text content.
     *
     * This will add any nodes which contain the given text content to the internal search results.
     *
     * You can optionally do a regex search, or append details to the search results rather than
     * replacing them; for convenience there are several wrapper functions below which are more
     * symantically appropriate, so it is recommended you only use the $find argument to this
     * function and use the wrappers if you require regex or append functionality.
     *
     * This function is chainable with other search functions.
     *
     * @param string $find
     * @param bool $regex
     * @param bool $append
     * @return WordCatXML
     */
    function findText(string $find, bool $regex = false, bool $append = false) {
        if(!$append) {
            $this->clearSearch();
        }
        $elements = [];
        $all = $this->document->getElementsByTagName("*");
        foreach($all as $node) {
            $found=false;
            foreach($node->childNodes as $childNode) {
                if($childNode instanceof DOMText) {
                    if($regex && preg_match($find, $childNode->data)) {
                        $found=true;
                    } elseif(!$regex && strpos($childNode->data, $find,0) !== false) {
                        $found=true;
                    }
                }
            }
            if($found) {
                $elements[] = $node;
            }
        }
        $this->setSearch($elements, $append);
        return $this;
    }

    /**
     * Do a fresh search for a regular expression. If the regular expression matches any
     * DOMText content, the containing elements will be added to the search results.
     *
     * You can optionally append to the search results rather than clearing them, but it is
     * recommended that you use "andFindRegex" instead.
     *
     * @param string $find
     * @param boolean $append
     * @return WordCatXML
     */
    function findRegex($find, $append = false) {
        return $this->findText($find, true, $append);
    }

    /**
     * Append any elements to the search list which contain the given string.
     * You can optionally use a regular expression, but it is recommended to use
     * "andFindRegex" instead (see below)
     *
     * @param string $find
     * @param boolean $regex
     * @return void
     */
    function andFindText($find, $regex=false) {
        return $this->findText($find, $regex, true);
    }

    /**
     * Append any elements to the search list which match the given regular expression
     *
     * @param string $find
     * @param boolean $regex
     * @return void
     */
    function andFindRegex($find) {
        return $this->findText($find, true, true);
    }

    /**
     * Find and replace text within XML nodes. This will operate on the search results,
     * replacing each instance of $find with $replace. If no search has been performed,
     * we will automatically perform the search first.
     *
     * You can also specify $replace as a function, in which case the function should
     * accept $find as a parameter, and should return a string.
     * 
     * You can optionally use regex to search/replace, although it's recommended to use
     * the replaceRegex method below.
     * 
     * As with all search functions this can be chained as it returns $this
     *
     * @param string $find
     * @param string|function $replace
     * @param boolean $regex
     * @return WordCatXML
     */
    function replaceText(string $find, $replace, $regex = false) {
        if(count($this->searchResults) < 1) {
            $this->findText($find, $regex);
        }
        foreach($this->searchResults as $node) {
            foreach($node->childNodes as $childNode) {
                if($childNode instanceof DOMText) {
                    if(is_callable($replace)) {
                        // As we have a callback, we want to call this for each time we find the string.
                        // The implementation for this varies depending on whether we are doing plain text or regex replacement:
                        if($regex) {
                            $pos = 0;
                            // The implementation for regex uses preg_match_all and preg_replace
                            if(preg_match_all($find, $childNode->data, $matches, PREG_OFFSET_CAPTURE, $pos)) {
                                $len = strlen($childNode->data);
                                foreach($matches[0] as $match) {
                                    $found = $match[0];
                                    $pre = substr($childNode->data, 0, $pos);
                                    $childNode->data = $pre . preg_replace($find, $replace($found), substr($childNode->data,$pos),1);
                                    $pos = $match[1] + (strlen($childNode->data) - $len);
                                }
                            }
                        } else {
                            // The implementation for plain text uses strpos and substr_replace:
                            $pos=0;
                            while($pos !== false) {
                                $pos = strpos($childNode->data, $find, $pos);
                                if($pos !== false) {
                                    $replacement = $replace($find);
                                    $childNode->data = substr_replace($childNode->data, $replacement, $pos, strlen($find));
                                    $pos = $pos + strlen($replacement);
                                }
                            }
                        }
                    } else {
                        // If we do not have a callback, replacement is simple:
                        if($regex) {
                            $childNode->data = preg_replace($find, $replace, $childNode->data);
                        } else {
                            $childNode->data = str_replace($find, $replace, $childNode->data);
                        }
    
                    }
                }
            }
        }
        return $this;
    }

    /**
     * Find and replace text within XML nodes. This is a wrapper around the replaceText
     * method.
     * 
     * This will operate on the search results, replacing each instance of $find with 
     * $replace. If no search has been performed, we will automatically perform the 
     * search first.
     *
     * You can also specify $replace as a function, in which case the function should
     * accept $find as a parameter, and should return a string.
     * 
     * You can optionally use regex to search/replace, although it's recommended to use
     * the replaceRegex method below.
     * 
     * As with all search functions this can be chained as it returns $this
     *
     * @param string $find
     * @param string|function $replace
     * @return WordCatXML
     */
    function replaceRegex($find, $replace) {
        return $this->replaceText($find, $replace, true);
    }

    /**
     * Get the search results. This returns an array (not a DOMNodeList!) of
     * elements which have been found and compiled using the findText function
     * and its wrappers
     *
     * @return array[DOMNode]
     */
    function getSearch() {
        return $this->searchResults;
    }

    /**
     * Perform an xpath query on the XML document.
     *
     * It's worth noting there are currently namespace issues with XML files
     * which have a namespace set but do not use a prefix for tag names.
     *
     * @param string $query
     * @return DOMNodeList
     */
    function xpath($query) {
        $xpath = new DOMXPath($this->document);
        if($namespaces = simplexml_import_dom($this->document->documentElement)) {
            foreach($namespaces as $ns=>$uri ) {
                if($ns != "") {
                    $xpath->registerNamespace($ns, $uri);
                }
            }
            $xpath->query($query);
        }
        return $xpath->query($query);
    }

    /**
     * Run a callback for each node in the search results.
     *
     * @param function $callback
     * @return WordCatXML
     */
    function forSearch($callback) {
        foreach($this->searchResults as $node) {
            $callback($node);
        }
        return $this;
    }

    /**
     * Get nodes from the given tag name. Note that this does not use the namespace;
     * use "p" rather than "w:p" for example.
     *
     * @param string $tag
     * @return DOMNodeList
     */
    function getNodesByTagName($tag) {
        return $this->document->getElementsByTagName($tag);
    }

    /**
     * Get nodes from the given tag name within the specified namespace
     *
     * @param string $namespace
     * @param string $tag
     * @return DOMNodeList
     */
    function getNodesByagNameNS($namespace, $tag) {
        return $this->document->getElementsByTagNameNS($namespace, $tag);
    }

    /**
     * Return a list of nodes, converted from a DOMNodeList to an array. If you specify
     * an attribute which the nodes use for an ID, this will be used as the index for the
     * array, otherwise it will be a flat array.
     *
     * @param DOMNodeList $list
     * @param string $idAttr
     * @return array
     */
    function nodesArray(DOMNodeList $list, $idAttr = null) {
        $array = [];
        foreach($list as $element) {
            if(!is_null($idAttr)) {
                if($attr = $element->getAttribute($idAttr)) {
                    $array[$attr] = $element;
                } else {
                    $array[]=$element;
                }
            } else {
                $array[]=$element;
            }
        }
        return $array;
    }
    
    /**
     * Fix document sections after merge or insertion
     * 
     * Document sections are defined by a w:sectPr element which has links to the relevant 
     * page layout and headers/footers. Each section needs to have this inside the last
     * paragraph of a document, except the last one which is a child of w:body.
     * 
     * This function ensures that if there are more than one w:sectPr elements as a child of
     * w:body, all but the last one will be inserted within its own paragraph.
     *
     * @return void
     */
    function fixDocumentSections() {
        $sectPrs = $this->document->getElementsByTagName("sectPr");
        $count = count($sectPrs);
        foreach($sectPrs as $index => $sectPr) {
            if($sectPr->parentNode && $sectPr->parentNode->nodeName == "w:body") {
                if($index < $count-1) {
                    $p = $this->insertNodeAfter($this->createNode("w:p"), $sectPr);
                    $pPr = $this->insertNodeInside($this->createNode("w:pPr"), $p);
                    $newSectPr = $this->cloneNode($sectPr);
                    $pPr->appendChild($newSectPr);
                    $this->removeNode($sectPr);
                }
            }

        }
    }

    /**
     * This function will insert a new section after the specified node.
     * 
     * This works by finding the final sectPr element, and copying it to a new paragraph
     * after the insertion point.
     *
     * @param DOMElement $insertAfter
     * @return DOMElement|void
     */
    function splitSection(DOMElement $insertAfter) {

        $insertAfter->setAttribute("wordcat_id", "splitsection");

        $getAllElements = $this->document->getElementsByTagName("*");
        $found = false;
        foreach($getAllElements as $index => $element) {
            if(!$found) {
                if($element->hasAttribute("wordcat_id") && $element->getAttribute("wordcat_id") == "splitsection") {
                    $element->removeAttribute("wordcat_id");
                    $found = true;
                }
            } else {
                if($element->nodeName == "w:sectPr") {
                    $insertPoint = $insertAfter;
                    while($insertPoint && $insertPoint->parentNode->nodeName != "w:body") {
                        $insertPoint = $insertPoint->parentNode;
                    }
                    if($insertPoint) {
                        $p = $this->insertNodeAfter($this->createNode("w:p"), $insertPoint);
                        $pPr = $this->insertNodeInside($this->createNode("w:pPr"), $p);
                        $newSectPr = $this->cloneNode($element);
                        $pPr->appendChild($newSectPr);
                        return $p;
                    }
                }
            }
        }
    }

    /**
     * Iterate from parent to parent until you find a node matching the given tag name.
     * 
     * This allows us, for example, to find the w:p node that contains a given w:t element.
     *
     * @param string $tagName
     * @param DOMNode $fromNode
     * @return DOMNode|null
     */
    function getClosestTagName(string $tagName, DOMNode $fromNode) {
        $find = $fromNode;
        while($find && $find->nodeName != $tagName) {
            $find = $find->parentNode;
        }
        return $find;
    }
}
