<?php
namespace WordCat;
use DOMDocument;
use DOMElement;
use DOMNode;
use DOMNodeList;
use DOMText;
use DOMXPath;
use Exception;

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
    private $searchFrom = null;
    private $searchTo = null;
    private $searchInside = null;
    private $allElements = null;

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
     * Create a cache of elements (with tagnames); all operations which would normally
     * use a full list of all elements in the document will then use this list instead
     * until this is reset or unset.
     *
     * Note that any nodes you add or remove will not be added or removed from this list,
     * so you will need to either call this method again or call operateOff if you then
     * wish to use these features including the nodes you have added or removed.
     *
     * This function is provided to increase speed when doing filtered searches; in
     * particular this is useful when you need to repeatedly work on a list of nodes
     * which are between two other nodes (using nodeIsAfter, nodeIsBefore, filterNodesBetween),
     * but it applies to anywhere that getElementsByTagName("*") would otherwise be used.
     *
     * @param array|DOMNodeList|null $nodes
     * @return void
     */
    function operateOn($nodes = null) {
        if(is_null($nodes)) {
            $this->allElements = $this->document->getElementsByTagName("*");
        }
        $this->allElements = $nodes;
    }

    /**
     * Remove the cached elements set by operateOn and return to using the full list of nodes
     *
     * @return void
     */
    function operateOff() {
        $this->allElements = null;
    }

    /**
     * Used internally in place of $this->document->getElementsByTagName("*") to provide cacheing
     * when doing searches.
     *
     * @return array|DOMNodeList
     */
    private function getAllElements() {
        if(is_null($this->allElements)) {
            return $this->document->getElementsByTagName("*");
        }
        return $this->allElements;
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
     * You can optionally elect not to clear the range (searchFrom and searchTo) filters
     * by providing false as an argument.
     *
     * @param bool $clearRange
     * @return WordCatXML
     */
    function clearSearch(bool $clearRange = true) {
        $this->searchResults = [];
        if($clearRange) {
            $this->searchFrom = null;
            $this->searchTo = null;
            $this->searchInside = null;
        }
        return $this;
    }

    /**
     * Set the starting node from which (and including) search results will be returned.
     *
     * Note that this filter is applied when you get the search results; the find routines
     * are not aware of this filter so changing this filter after you have search results
     * will effect the results of the search without the need to search again.
     *
     * @param DOMNode|null $node
     * @return WordCatXML
     */
    function searchFrom(?DOMNode $node) {
        $this->searchFrom = $node;
        return $this;
    }

    /**
     * Set the starting node up to which (not including) search results will be returned
     *
     * Note that this filter is applied when you get the search results; the find routines
     * are not aware of this filter so changing this filter after you have search results
     * will effect the results of the search without the need to search again.
     *
     * @param DOMNode|null $node
     * @return WordCatXML
     */
    function searchTo(?DOMNode $node) {
        $this->searchTo = $node;
        return $this;
    }

    /**
     * Set the node we want all search results to be within. This will not include the
     * specified node iteself in the search results.
     *
     * Note that this filter is applied when you get the search results; the find routines
     * are not aware of this filter so changing this filter after you have search results
     * will effect the results of the search without the need to search again.
     *
     * @param DOMNode|null $node
     * @return WordCatXML
     */
    function searchInside(?DOMNode $node) {
        $this->searchInside = $node;
        return $this;
    }

    /**
     * This function iterates over the search results to remove any duplicate entries for the same
     * node. This effects the search results, and ignores any additional filters, so it should be
     * called after any find functions.
     *
     * @return WordCatXML
     */
    function searchUnique() {
        $searchResults = [];
        foreach($this->searchResults as $node) {
            $found=false;
            foreach($searchResults as $check) {
                if($this->nodeIs($node, $check)) {
                    $found = true;
                }
            }
            if(!$found) {
                $searchResults[]=$node;
            }
        }
        $this->searchResults = $searchResults;
        return $this;
    }

    /**
     * Manipulate the search results by running a given callback on each item in the search
     * results list.
     *
     * It's important to note that the function will be run on all items in the search results
     * which does not include searchFrom, searchTo or searchInside filters - this means that if
     * your callback modifies the elements themselves you may modify more items than you had
     * intended!
     *
     * The callback should accept a DOMNode argument (the item in the search results),
     * and return a DOMNode (the item to replace it with). A common use case for this is
     * to find parent elements, or top level parents, of the elements within the search
     * results.
     *
     * @param callable $callback
     * @return WordCatXML
     */
    function searchLambda(callable $callback) {
        $searchResults = [];
        foreach($this->searchResults as $node) {
            $result = $callback($node);
            if($result instanceof DOMNode) {
                $searchResults[]=$result;
            }
        }
        $this->searchResults = $searchResults;
        return $this;
    }

    /**
     * Get the search results. This returns an array (not a DOMNodeList!) of
     * elements which have been found and compiled using the findText function
     * and its wrappers
     *
     * @return array
     */
    function getSearch() {
        return array_values(
            array_filter(
                $this->filterNodesBetween($this->searchResults, $this->searchFrom, $this->searchTo),
                function( $node ) {
                    if(!$node instanceof DOMNode) {
                        return false;
                    }
                    if(!is_null($this->searchInside) && !$this->nodeIsDescendant($node, $this->searchInside)) {
                        // Node is not a descendant of the given node
                        return false;
                    }
                    return true;
                }
            )
        );
    }

    /**
     * Run a callback for each node in the search results.
     *
     * @param function $callback
     * @return WordCatXML
     */
    function forSearch($callback) {
        foreach($this->getSearch() as $node) {
            $callback($node);
        }
        return $this;
    }

    /**
     * Search for nodes with the specified tag name, storing the list as an array
     * in the search results which can then be obtained using getSearch().
     *
     * Like all search features, this can be chained as it returns $this.
     *
     * You can optionally append to search results, but it is better to use the
     * andFindTagName alias as this is semantically clearer.
     *
     * @param string $tagName
     * @param boolean $append
     * @return WordCatXML
     */
    function findTagName(string $tagName, bool $append = false) {
        if(!$append) {
            $this->clearSearch(false);
        }
        $elements = [];
        $all = $this->document->getElementsByTagName($tagName);
        foreach($all as $node) {
            $elements[] = $node;
        }
        $this->setSearch($elements, $append);
        return $this;
    }

    /**
     * An alias for findTagName which appends to the search results.
     *
     * @param string $tagName
     * @return WordCatXML
     */
    function andFindTagName(string $tagName) {
        return $this->findTagName($tagName, true);
    }

    /**
     * Search for nodes with a given attribute, optionally if it has the given value,
     * storing the list as an array in the search results which can then be obtained
     * using getSearch().
     *
     * Like all search features, this can be chained as it returns $this.
     *
     * You can optionally append to search results, but it is better to use the
     * andFindAttribute alias as this is semantically clearer.
     *
     * @param string $attribute
     * @param string|null $value
     * @param boolean $append
     * @return void
     */
    function findAttribute(string $attribute, ?string $value = null, bool $append = false) {
        if(!$append) {
            $this->clearSearch(false);
        }
        $elements = [];
        foreach($this->getAllElements() as $node) {
            if($node->hasAttributes()) {
                if($node->hasAttribute($attribute)) {
                    if(is_null($value) || $node->getAttribute($attribute) == $value) {
                        $elements[] = $node;
                    }
                }
            }
        }
        $this->setSearch($elements, $append);
        return $this;
    }

    /**
     * An alias for findAttribute which appends to the search results.
     *
     * @param string $attribute
     * @param string $value
     * @return WordCatXML
     */
    function andFindAttribute(string $attribute, string $value) {
        return $this->findAttribute($attribute, $value, true);
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
            $this->clearSearch(false);
        }
        $elements = [];
        foreach($this->getAllElements() as $node) {
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
     * @param bool $append
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
     * @param bool $regex
     * @return WordCatXML
     */
    function andFindText($find, $regex=false) {
        return $this->findText($find, $regex, true);
    }

    /**
     * Append any elements to the search list which match the given regular expression
     *
     * @param string $find
     * @param bool $regex
     * @return WordCatXML
     */
    function andFindRegex($find) {
        return $this->findText($find, true, true);
    }

    /**
     * Search for text within paragraphs and merge all text elements together that
     * encompass the searched text.
     * 
     * This is useful if you need to search and replace text, but the text is not
     * stored contiguously in a single w:t element (this is often the case when a
     * user goes back and edits a paragraph; the word processor will often split
     * the text where the edit was made even if only the text is changed).
     * 
     * Note that this function is destructive; it will remove all elements between
     * the first and last elements needed to match the given text, effectively
     * replacing them all with a single w:t element.
     * 
     * The matching paragraphs will be stored in the internal search results, so
     * there is no need to do an additional findText (or similar) afterwards.
     * 
     * You should be able to use replaceText (or similar) to then make any
     * replacement you need to.
     *
     * @param string $find
     * @param boolean $regex
     * @return WordCatXML
     */
    function findAndMergeParagraphText(string $find, $regex = false, $append = false) {
        $elements = [];
        $all = $this->document->getElementsByTagName('p');
        foreach($all as $node) {
            $textNodes = $node->getElementsByTagName('t');
            $text="";
            $chunks = [];
            $lastNode = null;
            foreach($textNodes as $index => $textNode) {
                $chunks[$index] = strlen($text);
                $text.=$textNode->textContent;
                if( ($regex && preg_match($find, $text)) || (!$regex && strpos($text, $find) !== false)) {
                    $lastNode = $textNode;
                    foreach($chunks as $chunkIndex => $position) {
                        if( ($regex && preg_match($find, substr($text,$position))) || (!$regex && strpos(substr($text,$position), $find) !== false)) {
                            $firstNode = $textNodes[$chunkIndex];
                        } else {
                            break;
                        }
                    }
                    if($lastNode && $firstNode) {
                        $filteredNodes = $this->filterNodesBetween(null, $firstNode, $lastNode);
                        $newText = "";
                        foreach($filteredNodes as $filteredNode) {
                            if($filteredNode->nodeName == "w:t") {
                                $newText .= $filteredNode->textContent;
                            }
                            if(!$filteredNode->isSameNode($firstNode)) {
                                $this->removeNode($filteredNode);
                            }
                        }
                        $firstNode->textContent = $newText;
                        $elements[] = $firstNode;
                    }
                }
            }
            
        }
        $this->setSearch($elements, $append);
        return $this;
    }

    function findAndMergeParagraphRegex(string $find, bool $append = false) {
        return $this->findAndMergeParagraphText($find, $append);
    }

    function andFindAndMergeParagraphText(string $find, bool $regex = false) {
        return $this->findAndMergeParagraphText($find, $regex, true);
    }

    function andFindAndMergeParagraphRegex(string $find) {
        return $this->findAndMergeParagraphText($find, true, true);
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
     * @param bool $regex
     * @return WordCatXML
     */
    function replaceText(string $find, $replace, bool $regex = false) {
        $searchResults = $this->getSearch();
        if(count($searchResults) < 1) {
            $searchResults = $this->findText($find, $regex)->getSearch();
        }
        foreach($searchResults as $node) {
            foreach($node->childNodes as $childNode) {
                if($childNode instanceof DOMText) {
                    if(is_callable($replace) && !is_string($replace)) {
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
     * Find out if a given node appears later in the document than another
     * given node.
     *
     * @param DOMNode $node
     * @param DOMNode $afterNode
     * @return bool
     */
    function nodeIsAfter(DOMNode $node, DOMNode $afterNode) {
        $found = false;
        foreach($this->getAllElements() as $index => $element) {
            if(!$found) {
                if( $element->isSameNode($afterNode) ) {
                    $found = true;
                }
            } else {
                if( $element->isSameNode($node) ) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * Find out if a given node appears earlier in the document than another
     * given node.
     *
     * @param DOMNode $node
     * @param DOMNode $beforeNode
     * @return bool
     */
    function nodeIsBefore(DOMNode $node, DOMNode $beforeNode) {
        $found = false;
        foreach($this->getAllElements() as $index => $element) {
            if( $element->isSameNode($beforeNode) ) {
                $found = true;
            }
            if($element->isSameNode($node)) {
                return !$found;
            }
        }
        return false;
    }

    /**
     * Filter an array of nodes down to just the nodes that appear between a start and end
     * node in the tree. The start and end nodes are included in the resulting array of nodes.
     *
     * You can omit the start or end (or specify null) to just filter everything from the start
     * or up to the end node.
     *
     * Because finding a relative position in the tree is an expensive operation, using this
     * function can prove useful when trying to increase performance.
     *
     * If $nodes is set to null, all nodes between the $start and $end are returned.
     *
     * @param array $nodes
     * @param DOMNode|null $start
     * @param DOMNode|null $end
     * @return array[DOMNode]
     */
    function filterNodesBetween(?array $nodes = null, ?DOMNode $start = null, ?DOMNode $end = null) {
        $filtered = [];
        $foundStart = is_null($start);
        foreach($this->getAllElements() as $node) {
            if(!$foundStart && $node->isSameNode($start)) {
                $foundStart = true;
            }
            if($foundStart) {
                if(!is_null($nodes)) {
                    foreach($nodes as $index => $findNode) {
                        if($node->isSameNode($findNode)) {
                            unset($nodes[$index]);
                            $filtered[] = $node;
                        }
                    }
                } else {
                    $filtered[] = $node;
                }
            }
            if(!is_null($end) && $node->isSameNode($end)) {
                return $filtered;
            }
        }
        return $filtered;
    }

    /**
     * Find out if a given node is a descendant of another given node
     *
     * @param DOMNode $node
     * @param DOMNode $parent
     * @return bool
     */
    function nodeIsDescendant(DOMNode $node, DOMNode $parent) {
        $searchNode = $node;
        while($searchNode = $searchNode->parentNode) {
            if($searchNode->isSameNode($parent)) {
                return true;
            }
        }
        return false;
    }

    /**
     * Find out if a given node contains another given node
     *
     * @param DOMNode $container
     * @param DOMNode $contained
     * @return bool
     */
    function nodeContainsNode(DOMNode $container, DOMNode $contained) {
        if($container->isSameNode($contained)) {
            return true;
        }
        if($container->hasChildNodes()) {
            foreach($container->childNodes as $node) {
                if($this->nodeContainsNode($node, $contained)) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * Wrapper to compare two nodes; if they are the same node true is returned
     *
     * @param DOMNode $node
     * @param DOMNode $isNode
     * @return bool
     */
    function nodeIs(DOMNode $node, DOMNode $isNode) {
        return $node->isSameNode($isNode);
    }

    /**
     * Find a node after a given node in the document with the given tag name.
     *
     * @param string $tagName
     * @param DOMNode $afterNode
     * @return void
     */
    function getNextTagName(string $tagName, DOMNode $afterNode) {
        $found = false;
        foreach($this->getAllElements() as $index => $element) {
            if(!$found) {
                if( $element->isSameNode($afterNode) ) {
                    $found = true;
                }
            } else {
                if( $element->nodeName == $tagName ) {
                    return $element;
                }
            }
        }
    }

    /**
     * Find a node before a given node in the document with the given tag name.
     *
     * @param string $tagName
     * @param DOMNode $afterNode
     * @return void
     */
    function getPreviousTagName(string $tagName, DOMNode $beforeNode) {
        $found = null;
        foreach($this->getAllElements() as $index => $element) {
            if( $element->isSameNode($beforeNode) ) {
                return $found;
            }
            if( $element->nodeName == $tagName ) {
                $found = $element;
            }
        }
    }

    /**
     * This function will insert a new section after the specified node.
     *
     * You can specify a specific sectPr to copy and insert; if you omit this,
     * the final sectPr will be used by default.
     *
     * On success, the new paragraph (w:p) element is returned
     *
     * You can also specify a callback, which will be passed the sectPr we are copying
     * and the new sectPr we create (in that order), allowing the sectPrs to be modified.
     *
     * The callback will only be run if we have successfully inserted a sectPr.
     *
     * @param DOMElement $insertAfter
     * @param bool $sectPr
     * @param callable $callback
     * @return DOMElement|void
     */
    function splitSection(DOMElement $insertAfter, ?DOMNode $sectPr = null, ?callable $callback = null) {
        if(!$sectPr) {
            $sectPrs = $this->document->getElementsByTagName("sectPr");
            if($sectPrs->count() > 0) {
                $sectPr = $sectPrs->item($sectPrs->count()-1);
            }
        }
        // Now if we have a sectPr element, we copy it:
        if($sectPr) {
            // This is the correct sectPr
            $p = $this->insertNodeAfter($this->createNode("w:p"), $insertAfter);
            $pPr = $this->insertNodeInside($this->createNode("w:pPr"), $p);
            $newSectPr = $this->cloneNode($sectPr);
            $pPr->appendChild($newSectPr);
            if(is_callable($callback)) {
                $callback($sectPr, $newSectPr);
            }
            return $p;
        }
        return;


        $insertAfter->setAttribute("wordcat_id", "splitsection");
        $found = false;
        foreach($this->getAllElements() as $index => $element) {
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

    function normaliseSectionHeaderFooter(DOMElement $section, string $type = "default") {
        if($section->nodeName != "w:sectPr") {
            $sectPrs = $section->getElementsByTagName("sectPr");
            if($sectPrs->count()>0) {
                $section = $sectPrs->item(0);
            }
        }
        if($section->nodeName != "w:sectPr") {
            throw Exception("No sectPr tag found");
        }
        $headers = $section->getElementsByTagName("headerReference");
        $footers = $section->getElementsByTagName("footerReference");
        // Find given header:
        $header = null;
        foreach($headers as $node) {
            if($node->hasAttribute("w:type") && $node->getAttribute("w:type") == $type) {
                $header = $node;
                break;
            }
        }
        if(!is_null($header)) {
            $firstHeader = $this->cloneNode($header);
            $oddHeader = $this->cloneNode($header);
            $firstHeader->setAttribute("w:type", "first");
            $oddHeader->setAttribute("w:type", "first");
            $this->insertNodeAfter($firstHeader, $header);
            $this->insertNodeAfter($oddHeader, $header);
            foreach($headers as $node) {
                if(!$node->isSameNode($header)) {
                    $this->removeNode($node);
                }
            }
        }
        // Find given footer:
        $footer = null;
        foreach($footers as $node) {
            if($node->hasAttribute("w:type") && $node->getAttribute("w:type") == $type) {
                $footer = $node;
                break;
            }
        }
        if(!is_null($footer)) {
            $firstFooter = $this->cloneNode($footer);
            $oddFooter = $this->cloneNode($footer);
            $firstFooter->setAttribute("w:type", "first");
            $oddFooter->setAttribute("w:type", "first");
            $this->insertNodeAfter($firstFooter, $footer);
            $this->insertNodeAfter($oddFooter, $footer);
            foreach($footers as $node) {
                if(!$node->isSameNode($footer)) {
                    $this->removeNode($node);
                }
            }
        }

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
     * If you specify $replacedNode and this node is replaced, the replacement will be returned;
     * this allows you to ensure that you can maintain an insertion point if the document is
     * changed.
     *
     * @param DOMNode $replacedNode
     * @return DOMNode|null
     */
    function fixDocumentSections(?DOMNode $replacedNode=null) {
        $sectPrs = $this->document->getElementsByTagName("sectPr");
        $count = $sectPrs->count();
        foreach($sectPrs as $index => $sectPr) {
            if($sectPr->parentNode && $sectPr->parentNode->nodeName == "w:body") {
                if($index < $count-1) {
                    $p = $this->insertNodeAfter($this->createNode("w:p"), $sectPr);
                    $pPr = $this->insertNodeInside($this->createNode("w:pPr"), $p);
                    $newSectPr = $this->cloneNode($sectPr);
                    $pPr->appendChild($newSectPr);
                    if(!is_null($replacedNode) && $sectPr->isSameNode($replacedNode)) {
                        $replacedNode = $p;
                    }
                    $this->removeNode($sectPr);
                }
            }

        }
        return $replacedNode;
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

    /**
     * Get the parent node at the top level (root node).
     *
     * You can optionally specify the top level node; for instance you may want the parent
     * that is a child of the w:body node rather than the w:document
     *
     * @param DOMNode $node
     * @param DOMNode|null $topLevel
     * @return DOMNode
     */
    function getTopLevelParent(DOMNode $node, ?DOMNode $topLevel = null) {
        if(is_null($topLevel)) {
            $topLevel = $this->document->documentElement;
        }
        $find = $node;
        while($find && $find->parentNode && !$find->parentNode->isSameNode($topLevel)) {
            $find = $find->parentNode;
        }
        return $find;
    }

}
