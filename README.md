docx-converter
=================

This Ruby library (gem) parses and translates `.docx` Word documents into kramdown syntax, which allows for easy subsequent translation into `html` or `TeX` code via the excellent `kramdown` library. `kramdown` is a superset of `Markdown`. See http://kramdown.gettalong.org/ for more details.

A `.docx` file as written by modern versions of Microsoft Office is just a `.zip` file in disguise. It contains a directory tree containing XML files. Parsing of these compressed XML trees is rather staightforward, thanks to the `zip` and `nokogiri` Ruby libraries.

`docx-converter` contains a parser which translates all common Word document entities into corresponding kramdown syntax. It extracts images and converts them into `.jpg` files with a maximum width or height of 800 pixels.

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

`docx-converter inputfile format output_directory`

`format` can be either `kramdown`, `html` or `latex`. For example:

`docx-converter ~/Downloads/testdoc1.docx latex /tmp/docxoutput`

`output_directory` will be created if it doesn't exist. A subdirectory `/src` will be created, which is merely a convention to be closer to the `webgen` (http://webgen.gettalong.org/) file system standard, in case you want to generate a static webpage via `webgen`.

If you want to use docx_converter from a Ruby script, you can use it like this:

    r = DocxConverter::Render.new(options)
    rendered_filepaths = r.render(:html)
    
`options` is a hash with the following keys

* `:output_dir`: The directory to be created for the output files
* `:inputfile`: The path to the .docx file
* `:image_subdir_filesystem`: The subdirectory name into which images will be put. It will be created below `:output_dir`
* `:image_subdir_kramdown`: Usually this is identical to `:image_subdir_filesystem`. This value is the image url prefix for images in the kramdown output. This syntax looks like this `![]()`. 
* `:language`: The language to be used for the generated file names, which follows the `webgen` format: `ss.nnnn.ll.page`, where `ss` is a 2-digit sort number, `nnnn` is any string, `ll` is the language code.
* `:split_chapters`: when `true`, the output files will be split between headings which have the Word style "Heading1". This is useful for large documents. When `false`, no splitting is done and all content will be in the file `01.chapter01.ll.page`