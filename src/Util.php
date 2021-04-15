<?php
namespace WordCat;

class Util {
    const ID_EXISTS_FROMID = 1;
    const ID_EXISTS_REGEX = 2;
    const ID_EXISTS_NUMERIC_FIELD = 4;

    static function genericNextId(string $fromId, string $regex, int $numericField, callable $idExists, int $flags = 0) {
        $matches = null;
        $idExistsParams = array_filter([$fromId, $regex, $numericField], function($key) use($flags) { return $flags & (2 ** $key) > 0; },ARRAY_FILTER_USE_KEY);
        if(preg_match($regex,$fromId,$matches)) {
            $preString = "";
            $postString = "";
            for($i = 1; $i < $numericField; $i++) {
                if($i == $numericField - 1) {
                    $preString .= '${'.$i.'}';
                } else {
                    $preString .= "\$$i";
                }
            }
            for($i = $numericField + 1; $i < count($matches); $i++) {
                $postString .= "\$$i";
            }
            $number = intval($matches[$numericField]);
            do {
                $toId = preg_replace($regex, $preString . ++$number . $postString, $fromId);
                if(is_null($exists = $idExists($toId, ...$idExistsParams))) {
                    return $toId;
                }
                if($toId == $fromId) {
                    $toId = null;
                }
            } while(!is_null($toId) && $exists );
            if(is_null($exists)) {
                return $fromId;
            }
            return $toId;
        }
    }

    static function sortDirectoriesFirst(&$dir, $ascending = true) {
        if(is_array($dir)) {
            foreach($dir as &$content) {
                self::sortDirectoriesFirst($content, $ascending);
            }
            uksort($dir, function( &$a, $b) use($ascending) {
                if($ascending) {
                    $c=$a;
                    $a=$b;
                    $b=$c;
                }
                if(substr($a,-1) == "/" && substr($b,-1) != "/") {
                    return -1;
                }
                if(substr($a,-1) != "/" && substr($b,-1) == "/") {
                    return 1;
                }
                if($a == $b) return 0;
                return $a < $b ? -1 : 1;
            });
        }
    }

    static function sortDirectoriesLast(&$dir, $ascending = true) {
        if(is_array($dir)) {
            foreach($dir as &$content) {
                self::sortDirectoriesLast($content, $ascending);
            }
            uksort($dir, function( $a, $b)  use($ascending) {
                if($ascending) {
                    $c=$a;
                    $a=$b;
                    $b=$c;
                }
                if(substr($a,-1) == "/" && substr($b,-1) != "/") {
                    return 1;
                }
                if(substr($a,-1) != "/" && substr($b,-1) == "/") {
                    return -1;
                }
                if($a == $b) return 0;
                return $a < $b ? -1 : 1;
            });
        }
    }

    static function customDirectorySort(&$dir, $pathsFirst = [], $pathsLast = []) {
        if(is_array($dir)) {
            uksort($dir, function( $a, $b) use($pathsFirst, $pathsLast) {
                // Sort files in $pathsFirst first...
                $c = array_search($a, $pathsFirst);
                $d = array_search($b, $pathsFirst);
                if($c !== false && $d !== false) {
                    if($c==$d) { return 0; }
                    return $c < $d ? -1 : 1;
                } elseif($c !== false && $d === false) {
                    return -1;
                } elseif($c === false && $d !== false) {
                    return 1;
                }
                // Sort files in $pathsLast last...
                $c = array_search($a, $pathsLast);
                $d = array_search($b, $pathsLast);
                if($c !== false && $d !== false) {
                    if($c==$d) { return 0; }
                    return $c < $d ? 1 : -1;
                } elseif($c !== false && $d === false) {
                    return 1;
                } elseif($c === false && $d !== false) {
                    return -1;
                }
                if($a == $b) return 0;
                return $a < $b ? -1 : 1;
            });
        }
    }

    static function flattenDirectoryArray(array $dir, bool $flat = true, int $showDirs = 0, string $path = "") {
        $out = [];
        if(is_array($dir)) {
            foreach($dir as $key => $value) {
                if(is_array($value)) {
                    $results = self::flattenDirectoryArray($value, $flat, $showDirs, "$path$key");
                    if($flat) {
                        if($showDirs >= 0) {
                            $out["$path$key"] = null;
                        }
                        if($showDirs <= 0) {
                            foreach($results as $k=>$v) {
                                $out["$path$k"] = $v;
                            }
                        }
                    } else {
                        if($showDirs >= 0 || count($results) > 0) {
                            $out["$path$key"] = $results;
                        }
                    }
                } else {
                    if($showDirs <= 0) {
                        $out["$path$key"] = $value;
                    }
                }
            }
        }

        return $out;
    }
}