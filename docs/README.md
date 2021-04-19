# WordCat Documentation

## WordCat's view of DocX

WordCat is naive; it has only a very minimum level of abstraction of the underlying xml structure of docx files. Docx files have a structure of linked XML files, and WordCat largely ignores them; WordCat is aware of the following:



## Creating a WordCat instance

WordCat requires an existing document to work on. We can create a new WordCat object a couple of ways:

### Direct instantiation:
```php
$newWordCatInstance = new WordCat("/path/to/file.docx");
```

### Instantiation via helper:
```php
$newWordCatInstance = WordCat::instance("/path/to/file.docx");
```

### Instantiation behind the scenes
When WordCat is instantiated, it creates a copy of the file you are working on in the temporary directory; this ensures that your original document will not be polluted with changes as you manipulate it with WordCat. Once the file is copied, WordCat creates an internal `ZipArchive` object and loads the temporary file into it, allowing us to manipulate the content within.

While WordCat does try to clean these files up automatically, PHP destructors are not super reliable, so it's always a good idea to make sure you close the file with the `close()` method when you have finished - this should ensure that any temporary files are cleaned up and will close the `ZipArchive` properly.

