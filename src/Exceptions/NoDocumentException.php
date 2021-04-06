<?php
namespace WordCat\Exceptions;

use Exception;
use Throwable;

class NoDocumentException extends Exception {
    public function __construct($message = null, $code = 0, Throwable $previous = null) {
        if(is_null($message)) {
            $message = "WordCat instance does not have a docx document open";
            parent::__construct($message, $code, $previous);
        }
    }
}