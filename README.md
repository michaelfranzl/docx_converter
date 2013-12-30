docx-converter
=================

This Ruby library (gem) parses and translates `.docx` Word documents into kramdown syntax, which allows for easy subsequent translation into `html` or `TeX` code via the excellent `kramdown` library. `kramdown` is a superset of `Markdown`. See http://kramdown.gettalong.org/ for more details.

A `.docx` file as written by modern versions of Microsoft Office is just a `.zip` file in disguise. It contains a directory tree containing XML files. Parsing of these compressed XML trees is rather staightforward, thanks to the `zip` and `nokogiri` Ruby libraries.

`docx-converter` contains a parser which translates all common Word document entities into corresponding `kramdown` syntax. It extracts images and converts them into `.jpg` files with a maximum width or height of 800 pixels.

Output files and directories will be created according to the `webgen` conventions. This is useful when you want to generate a static website with the `webgen` gem after you have converted your `.docx` file into `html`. The file naming is in the format `ss.nnnn.ll.page`, where `ss` is a 2-digit sort number, `nnnn` is the main file name, `ll` is the language code. For more information on `webgen` see http://webgen.gettalong.org/

`docx_converter` was written for our project `publishr_web`, see http://documentation.red-e.eu/publishr/index.html

Supported Word elements:

* Paragraph
* Line break
* Page break
* Bold
* Italic
* paragraph styles like Heading1, Heading2 and Title
* character styles like Strong and Quote
* footnotes
* images including captions
* non-breaking spaces

Installation
----------

`gem install docx-converter`

Usage
----------

From the command line:

`docx-converter` `inputfile` `format` `output_directory`

`format` can be either `kramdown`, `html` or `latex`. For example:

`docx-converter` `~/Downloads/testdoc1.docx` `latex` `/tmp/docxoutput`

`output_directory` will be created if it doesn't exist. A subdirectory `/src` will be created by default, which is merely a convention to be identical with the `webgen` file system standard.

If you want to use `docx_converter` from a Ruby script, you can use the API like this:

    r = DocxConverter::Render.new(options)
    rendered_filepaths = r.render(:html)
    
`options` is a hash with the following keys

* `:output_dir`: The directory to be created for the output files. A subdirectory `/src` will be created by default, which is merely a convention to be identical with the `webgen` file system standard.
* `:inputfile`: The path to the `.docx` file to be parsed
* `:image_subdir_filesystem`: The subdirectory name into which images will be put. It will be created below the `/src` subdirectory.
* `:image_subdir_kramdown`: Usually this is identical to `:image_subdir_filesystem` and should only be different when you do further manual postprocessing with the kramdown output. This string will be added as a prefix for images in the final kramdown output. An example: `![image description](/image_subdir_kramdown/imagename.jpg)`.
* `:language`: The language to be used for the generated file names. See `webgen` conventions above.
* `:split_chapters`: when `true`, the output files will be split between headings which have the Word paragraph style "Heading1". This is useful for large documents. When `false`, no splitting is done and all content will be output to the file `01.chapter01.ll.page`. Footnotes will be split correctly into the various chapters.