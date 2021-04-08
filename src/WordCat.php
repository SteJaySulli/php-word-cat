<?php
namespace WordCat;

use DOMDocument;
use DOMNode;
use Exception;
use WordCat\Exceptions\NoDocumentException;
use ZipArchive;

/**
 * WordCat Class
 *
 * This class provides the API to deal with docx documents as a whole.
 *
 * The methods provided are primarily geared to dealing with the main document itself;
 * The methods of the main document's WordCatXML object are exposed as a convenience.
 *
 * I may remove this behaviour at a later date, as I only did this to expose certain
 * elements which should really have their own wrapper methods instead, so please do
 * not rely on this as a permanent feature!
 *
 */
class WordCat {

    public $archive = null;
    private $parts = [];

    /**
     * Read a file from the docx archive and attempt to open it as an XML DOMDocument.
     * This function is primarily used internally, but it needs to remain public so that
     * WordCatXML objects can use it
     *
     * @param string $path - The path of the file within the archive
     * @return DOMDocument
     */
    function readXML(string $path) {
        if(!$this->archive instanceof ZipArchive) {
            throw new NoDocumentException;
        }
        $xml = new DOMDocument();
        $xml->loadXML($this->archive->getFromName($path));
        return $xml;
    }

    /**
     * Save a modified DOMDocument back to the archive.
     * This function is primarily used internally, but it needs to remain public so that
     * WordCatXML objects can use it
     *
     * @param string $path - The path of the file within the archive
     * @return void
     */
    function writeXML(string $path) {
        $this->writeFile($path, $this->getXML($path)->getDocumentXML());
    }

    /**
     * The constructor. Provide the path of your docx file and it will be loaded.
     * Alternatively you can provide the string content of the docx directly.
     *
     * @param string $wordFile
     */
    function __construct(?string $wordFile = null) {
        $this->tempName = null;
        $this->wordFile = null;
        $this->archive = null;
        if(!is_null($wordFile)) {
            if(preg_match("/.docx$/i",$wordFile)) {
                $this->load($wordFile);
            } else {
                $this->set($wordFile);
            }
        }
    }

    function load(string $wordFile) {
        $this->archive = new ZipArchive();
        // Copy the archive to a temp file so we can work on it without wiping out the original:
        $this->tempName = tempnam(sys_get_temp_dir(), 'php-wordcat-doc');
        $this->wordFile = $wordFile;
        copy($wordFile, $this->tempName);
        $this->archive->open($this->tempName);
        $this->getXML("word/document.xml");
        $this->getXML("word/styles.xml");
        return $this;
    }

    function set(string $content) {
        // Copy the archive to a temp file so we can work on it without wiping out the original:
        $this->tempName = tempnam(sys_get_temp_dir(), 'php-wordcat-doc');
        $this->wordFile = "new-word-cat-document.docx";
        file_put_contents($this->tempName, $content);
        $this->archive = new ZipArchive();
        $this->archive->open($this->tempName);
        $this->getXML("word/document.xml");
        $this->getXML("word/styles.xml");

        return $this;
    }

    /**
     * The destructor. This should ensure that temporary archives are automatically deleted
     */
    function __destruct() {
        if(is_file($this->tempName)) {
            unlink($this->tempName);
        }
    }

    /**
     * Close the current archive; once you do this you can discard this object as it will no
     * longer be useable.
     *
     * @return void
     */
    function close() {
        if(!$this->archive instanceof ZipArchive) {
            try {
                $this->archive->close();
            } catch (Exception $e) {
                // no-op - If we fail to close the archive at this point, we still want to remove
                // the temporary file.
            }
        }
        if(is_file($this->tempName)) {
            unlink($this->tempName);
            $this->tempName = null;
        }

    }

    /**
     * Save the document as a docx file. Internally this closes the temporary archive file,
     * makes a copy of it to the new filename and then opens it again - it also ensures any
     * changed XML content is written back to the archive before it is closed
     *
     * @param string $filename - The file path where the docx file should be stored
     * @return void
     */
    function saveAs(string $filename) {
        if(!$this->archive instanceof ZipArchive) {
            throw new NoDocumentException;
        }
        // Ensure all parts are updated to zip first...
        foreach($this->parts as $part) {
            $part->store();

        }
        $this->archive->close();
        // We need to manually set the general purpose bits within the file header
        // so it can be correctly identified as a docx file:
        $contents = file_get_contents($this->tempName);
        $contents[6] = "\x6";
        $contents[7] = "\x0";
        file_put_contents($filename, $contents);
        $this->archive->open($this->tempName);
        // If we have used a different filename to the one we loaded, update the name so subsequent
        // calls to save() don't overwrite the original file!
        if($this->wordFile != $filename) {
            $this->wordFile = $filename;
        }
    }

    /**
     * Save the docx file; This is the same as saveAs (see above), but it saves to whatever
     * filename you loaded.
     *
     * @return void
     */
    function save() {
        return $this->saveAs($this->wordFile);
    }

    /**
     * Archive stat() wrapper -- This returns the file stat or false if the file doesn't exist.
     *
     * The response will look something like this: [
     *   "name" => "word/document.xml",
     *   "index" => 18,
     *   "crc" => 555398942,
     *   "size" => 31740,
     *   "mtime" => 1616520898,
     *   "comp_size" => 3067,
     *   "comp_method" => 8,
     *   "encryption_method" => 0,
     * ]
     *
     * @param string $path
     * @return false|array
     */
    function statArchive(string $path) {
        if(!$this->archive instanceof ZipArchive) {
            throw new NoDocumentException;
        }
        return $this->archive->statName($path);
    }

    /**
     * Archive reader - this returns the contents (as a string blob) of the given file within
     * the archive
     *
     * @param string $path
     * @return string
     */
    function readFile(string $path) {
        if(!$this->archive instanceof ZipArchive) {
            throw new NoDocumentException;
        }
        return $this->archive->getFromName($path);
    }

    /**
     * Archive writer - this saves the given string blob to the given filename within the archive
     *
     * @param string $path - The filename within the archive
     * @param string $content - The string containing the contents of the file
     * @return boolean
     */
    function writeFile($path, $content) {
        if(!$this->archive instanceof ZipArchive) {
            throw new NoDocumentException;
        }
        if($this->archive->addFromString($path, $content)) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * Get a list of resource IDs within the document's relationships.
     *
     * Any Ids without any numeric digits are listed first in alphabetical order,
     * and any Ids which do have numeric digits are sorted according to the value
     * of the numeric content when all alphas are removed.
     *
     * @return array[string]
     */
    function getResourceIds() {
        // Get the relations XML document
        $relations = $this->getXML("word/_rels/document.xml.rels");
        // Find all the elements with an Id attribute
        $elements = $relations->xpath('//*[@Id]');
        $array = [];
        foreach($elements as $element) {
            $array[] = $element->getAttribute("Id");
        }
        // Sort by Id, ensuring that any Ids without numeric portions are first
        usort($array, function( $a, $b) {
            $numA = preg_replace('/[^0-9]/','',$a);
            $numB = preg_replace('/[^0-9]/','',$b);
            // If neither ID has numeric content, sort alphabetically:
            if($numA == $numB && $numA == "") {
                if($a == $b) {
                    return 0;
                }
                return ($a < $b) ? -1 : 1;
            }
            // If one ID has numeric content but not the other, we want the numeric one to be last:
            if($numA == "") {
                return -1;
            }
            if($numB == "") {
                return 1;
            }
            // Both IDs have numeric content, so sort by the value of the concatenation of the numeric digits:
            if(intval($numA) == intval($numB)) {
                return 0;
            }
            return (intval($numA) < intval($numB)) ? -1 : 1;
        });
        // Return unique values:
        return array_unique($array);
    }

    /**
     * Create a resource ID for the document. This works by attempting to add one
     * or more to the highest numeric in the IDs.
     *
     * While we assume that the IDs all follow the same alphanumeric scheme (eg rId1, rId3, etc)
     * there should be no problem if the ID fields follow different forms - As long as there is
     * some numeric content.
     *
     *
     * @param integer $add
     * @return void
     */
    function newResourceId($add = 0) {
        $ids = $this->getResourceIds();
        $idCount = count($ids);
        // If there are currently no IDs at all, we will return a default of "rId1":
        if($idCount < 1) {
            return "rId1";
        }
        $last = $ids[$idCount - 1];
        $alpha = preg_replace('/[0-9]/', '', $last);
        $num=preg_replace('/[^0-9]/', '', $last);
        // Add some protection against alpha-only
        if(empty($num)) {
            $num="0";
        }
        $num = intval($num) + 1 + $add;
        return "$alpha$num";
    }

    /**
     * Get a resource ID map array.
     *
     * This looks at the resource IDs for the provided "source" document, returning a map
     * of those IDs to suggested replacements for use in the current (destination) document.
     *
     * The array will be in the form "oldId" => "newId", where oldId is the ID as set in
     * the source document, and newId is the ID we will use in the current (destination) document.
     *
     * @param WordCat $source
     * @return array
     */
    function getResourceIdMap(WordCat $source) {
        $map = [];
        foreach($source->getResourceIds() as $index=>$id) {
            $map[$id] = $this->newResourceId($index);
        }
        return $map;

    }

    /**
     * Copy XML content using an xpath, Ignoring any conflicting IDs.
     *
     * This is used by the mergeStyles function, but could find use elsewhere. See mergeStyles
     * below for an example of usage!
     *
     * This function is quite naive; any nodes it does merge from the source document will simply
     * be appended to the end of the node that the first matching node appears in.
     *
     * @param WordCat $source - The document you want to copy from
     * @param string $filename - The path of the xml file in the archive (in both documents)
     * @param string $nodeXpath - The xpath to use to find elements to copy
     * @param string $idAttr - The attribute within the target nodes which specifies the element ID
     * @return integer - The number of items that have been merged
     */
    function mergeXpath(WordCat $source, string $filename, string $nodeXpath, string $idAttr) {
        $dDoc = $this->getXML($filename);
        $sDoc = $source->getXML($filename);
        $dNodes = $dDoc->nodesArray($dDoc->xpath($nodeXpath), $idAttr);
        $nodes = array_filter($sDoc->nodesArray($sDoc->xpath($nodeXpath), $idAttr), function($key) use($dNodes) {
            return !isset($dNodes[$key]);
        }, ARRAY_FILTER_USE_KEY);

        $importCount = -1;
        if(count($nodes) > 0) {
            $importCount++;
            $dNode = array_values($dNodes)[0]->parentNode;
            foreach($nodes as $sNode) {
                $dDoc->importNodeInside($sNode, $dNode);
                $importCount ++;
            }
        }
        $dDoc->store();
        return $importCount;
    }
    /**
     * This function merges any styles which appear in a source document and aren't already set
     * in the current document.
     *
     * This allows us to ensure that any content referring to these styles will have the styles
     * set, while not overwriting the current document's styles.
     *
     * @param WordCat $source
     * @return void
     */
    function mergeStyles(WordCat $source) {
        return $this->mergeXpath($source, "word/styles.xml", "//w:styles/w:style", "w:styleId");
    }

    /**
     * This function merges relationships from a source document to the current document.
     *
     * This works by transposing the id of each relationship in the source document so that
     * it does not conflict with relationships already set in the current document, then
     * copying the relationship nodes from the source document to the current document.
     *
     * We also check if the relationship's Target is a file within the archive; if it is
     * we also copy that file.
     *
     * Finally we go over the source document and update any references to the resource id
     * within the document itself, so we are ready to copy them too by the time this function
     * has executed
     *
     * @param WordCat $source
     * @return void
     */
    function mergeRelationships(WordCat $source) {
        $idMap = $this->getResourceIdMap($source);
        $dDoc = $this->getXML("word/_rels/document.xml.rels");
        $sDoc = $source->getXML("word/_rels/document.xml.rels");
        $dNode = $dDoc->getNodesByTagName("Relationships")[0];

        // Change the IDs for all relationships in the source document
        foreach($sDoc->getNodesByTagName("Relationship") as $node) {
            $idAttr = $node->getAttribute("Id");
            if( $id = $idMap[$idAttr]) {
                $node->setAttribute("Id", $id);

                $sName = $node->getAttribute("Target");
                if($source->statArchive("word/$sName") !== false) {
                    $dName = $sName;
                    // If the file already exists, we need to change it...
                    $n = 0;
                    while($this->statArchive("word/$dName") !== false) {
                        $n++;
                        $dName=$sName;
                        if(($pos = strrpos($sName, '/'))!==false) {
                            $dName=substr($sName, 0, $pos+1) . "i$n-" . substr($sName, $pos+1);

                        } else {
                            $dName = "i$n-$sName";
                        }
                    }
                    // If the filename has changed, we need to change the reference to it...
                    if($dName != $sName) {
                        $node->setAttribute("Target", $dName);
                    }
                    $this->writeFile("word/$dName", $source->readFile("word/$sName"));
                }
                $dDoc->importNodeInside($node, $dNode);
            } else {
            }
        }
        $sDoc = $source->getXML("word/document.xml");
        // Change IDs within source document itself, ready for it to be merged:
        foreach($sDoc->xpath('//*[@r:id]') as $node) {
            $attr = $node->getAttribute("r:id");
            if(isset($idMap[$attr])) {
                $node->setAttribute("r:id", $idMap[$attr]);
            }
        }
        foreach($sDoc->xpath('//*[@r:embed]') as $node) {
            $attr = $node->getAttribute("r:embed");
            if(isset($idMap[$attr])) {
                $node->setAttribute("r:embed", $idMap[$attr]);
            }
        }
        $dDoc->store();
    }

    /**
     * Insert another document after a given node within the current document.
     *
     * This allows us to insert the content of one document into another in as
     * simple a way as possible without the hassle of trying to change layout and
     * styling (as it may be half way through the document).
     *
     * It's worth noting that the document will be inserted at root level after
     * whatever node you specify; that means that if you specify a node that isn't
     * a direct child of the root there may be content left between the "afterNode"
     * node and the inserted document content.
     *
     * This isn't usually a problem when inserting into a placeholder, but it does
     * not allow you to insert the document content within a table cell (for example);
     * it would appear after the block the table is in.
    *
    * @param WordCat $source
    * @param DOMNode $afterNode
    * @param bool $splitSections
    * @return DOMNode
    */
    function insertDocument(WordCat $source, DOMNode $afterNode, bool $splitSections = false) {
        $this->mergeStyles($source);
        $this->mergeRelationships($source);
        $sDoc = $source->getXML("word/document.xml");
        $dDoc = $this->getXML("word/document.xml");
        $sBody = $sDoc->getNodesByTagName("body")[0];

        $after=$afterNode;

        if($splitSections) {
            if($sect = $dDoc->splitSection($afterNode)) {
                $after = $sect;
            }
        }

        // Find the top-level node for the insertion point
        while($after->parentNode && $after->parentNode->nodeName != "w:body") {
            $after = $after->parentNode;
        }
        // Insert all nodes from the source to destination
        foreach($sBody->childNodes as $node) {
            $after = $dDoc->importNodeAfter($node, $after);
        }

        $after = $dDoc->fixDocumentSections($after);

        $dDoc->store();
        return $after;
    }

    /**
     * Insert a text run node after the given node.
     *
     * Specify your target location with $insertionNode. The new run will be inserted in the most
     * appropriate position after this node, depending on its type:
     *
     * w:t (text node) - The node is appended after the insertionNode's parent.
     * w:r (run) - The new run is appended after the insertion node
     * w:p (paragraph) or any other type: The run is appended as the last child of the insertion node
     *
     * This allows maximal flexibility without being too specific
     *
     * @param DOMNode $insertionNode
     * @return DOMNode
     */
    function insertRun(DOMNode $insertionNode) {
        $xml = $this->getXML("word/document.xml");
        $node =$xml->createNode("w:r");
        switch($insertionNode->nodeName) {
            case "w:t":
                return $xml->insertNodeAfter($node, $insertionNode->parentNode);
            case "w:r":
                return $xml->insertNodeAfter($node, $insertionNode);
            default:
                return $xml->insertNodeInside($node, $insertionNode);
        }
    }

    /**
     * Insert text into a run. If the insertionNode you specify is not a run (w:r),
     * this will do nothing at all and return void
     *
     * @param string $text
     * @param DOMNode $insertionNode
     * @return DOMNode|void
     */
    function insertText(string $text, DOMNode $insertionNode) {
        $xml = $this->getXML("word/document.xml");
        if($insertionNode->nodeName == "w:r") {
            $node =$xml->createNode("w:t", null, $text);
            return $xml->insertNodeInside($node, $insertionNode);
        }
    }

    /**
     * Insert a paragraph. Set $contents to null to create an empty paragraph containing a run.
     * Otherwise $contents can be a string, DOMNode, or an array of strings and DOMNodes, which will
     * be added inside the paragraph.
     *
     * The paragraph should be inserted within the w:body element - If a child of w:body is specified
     * as the insertion node, the new paragraph is added at whatever parent node is a child of w:body,
     * otherwise the paragraph is added after the specified insertion node.
     *
     * @param null|string|DOMNode|array[string]|array[DOMNode]|array[DOMNode|string] $contents
     * @param DOMNode $insertionNode
     * @return DOMNode
     */
    function insertParagraph($contents, DOMNode $insertionNode) {
        $xml = $this->getXML("word/document.xml");
        $node =$xml->createNode("w:p");
        while($insertionNode->nodeName != "w:body" && $insertionNode->parentNode && $insertionNode->parentNode->nodeName != "w:body") {
            $insertionNode = $insertionNode->parentNode;
        }
        if($insertionNode->nodeName == "w:body") {
            $returnNode = $xml->insertNodeInside($node, $insertionNode);
        } else {
            $returnNode = $xml->insertNodeAfter($node, $insertionNode);
        }

        if(!is_array($contents)) {
            $contents = [$contents];
        }
        $run = $this->insertRun($returnNode);

        foreach($contents as $child) {
            if(is_string($child)) {
                $text = $xml->createNode("w:t", null, $child);
                $insertedText = $xml->insertNodeInside($text, $run);
            } elseif(!is_null($child)) {
                $insertedNode = $xml->insertNodeInside($child, $run);
            }
        }
        return $returnNode;
    }

    /**
     * This function inserts an image after a given node. This will use a very basic image definition
     * to provide the minimum possible implementation that can insert an image.
     *
     * The new image will be appended to the insertionNode as its last child. It's recommended to use
     * a run (w:r) element as your insertionNode but this isn't compulsory.
     *
     * You should specify a width and height in pt - If you don't a default size of 300×300pt is used.
     * The image will be resized within these constraints while preserving the aspect ratio.
     *
     * You can also specify a title if you wish.
     *
     * @param string $imageFile
     * @param DOMNode $insertionNode
     * @param integer $width
     * @param integer $height
     * @param string $title
     * @return DOMNode
     */
    function insertImage(string $imageFile, DOMNode $insertionNode, $width=null, $height=null, $title="") {
        $xml = $this->getXML("word/document.xml");
        $rel = $this->getXML("word/_rels/document.xml.rels");
        $id = $this->newResourceId();
        // Figure out a filename for the image within the archive
        $sName = $dName = "media/".substr(basename($imageFile), -16);
        if($this->statArchive("word/$sName") !== false) {
            $dName = $sName;
            // If the file already exists, we need to change it...
            $n = 0;
            while($this->statArchive("word/$dName") !== false) {
                $n++;
                $dName=$sName;
                if(($pos = strrpos($sName, '/'))!==false) {
                    $dName=substr($sName, 0, $pos+1) . "i$n-" . substr($sName, $pos+1);

                } else {
                    $dName = "a$n-$sName";
                }
            }
        }
        // Add the file to the archive:
        $this->writeFile("word/$dName", file_get_contents($imageFile));
        // Add the releationship
        $rel->insertNodeInside(
            $rel->createNode("Relationship", [
                "Id" => $id,
                "Type" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                "Target" => "$dName"
            ]),
            $rel->getDOMDocument()->documentElement
        );
        // Set up the image details
        $imgInfo = getimagesize($imageFile);
        $ratio = $imgInfo[0] / $imgInfo[1];
        if(is_null($width) && is_null($height)) {
            $width = 300;
            $height = 300;
        }
        if(is_null($width)) {
            $width = $height;
        }
        if(is_null($height)) {
            $height = $width;
        }
        // Ensure the image is sized correctly while maintaining the aspect ratio
        if($width > $height) {
            $width = round($height * $ratio);
        } else {
            $height = round($width / $ratio);
        }

        // Add the nodes to the main document:
        $newNode = $xml->createNodeTree([
            [
                "tag" => "w:pict",
                "children" => [
                    [
                        "tag" => "v:shape",
                        "attributes" => [
                            "style"=>"width:{$width}pt; height:{$height}pt; margin-left:0pt; margin-top:0pt; mso-position-horizontal:left; mso-position-vertical:top; mso-position-horizontal-relative:char; mso-position-vertical-relative:line;"
                        ],
                        "children" => [
                            [
                                "tag" => "v:imagedata",
                                "attributes" => [
                                    "r:id" => $id,
                                    "o:title" => $title
                                ]
                            ]
                        ]
                    ]
                ]
            ]
        ], $insertionNode, function($a,$b) use($xml) { $xml->insertNodeInside($a,$b); });
        return $newNode;
    }

    /**
     * Get a WordCatXML object for the given XML file within the document's archive.
     * WordCat keeps a record of the XML files you load using this, so it will return
     * the existing copy if there is one, otherwise it loads it from the archive.
     *
     * This is essential to the function of this library, as we use it a lot to ensure
     * we are always operating on the correct XML objects without losing changes in between
     * calls
     *
     * @param string $path
     * @return WordCatXML
     */
    function getXML(string $path) {
        if(!isset($this->parts[$path])) {
            $this->parts[$path] = new WordCatXML($this, $path);
        }
        return $this->parts[$path];
    }

    /**
     * Get a WordCat instance from a docx file.
     *
     * This is essentially an alias for new WordCat($instance), but it will gracefully cope
     * with being passed a WordCat instance.
     *
     * @param string|WordCat $instance
     * @return WordCat
     */
    static function instance($instance) {
        if(is_string($instance)) {
            return new self($instance);
        }
        if($instance instanceof self) {
            return $instance;
        }
    }

    function __call($fnName, $arguments) {
        $xml = $this->getXML("word/document.xml");
        if(in_array($fnName, get_class_methods($xml))) {
            return $xml->{$fnName}(...$arguments);
        } else {
            throw new Exception("Method $fnName does not exist in WordCat");
        }
    }
}

