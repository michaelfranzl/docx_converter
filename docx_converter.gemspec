# -*- encoding: utf-8 -*-
$:.push File.expand_path("../lib", __FILE__)
require "docx_converter/version"

Gem::Specification.new do |s|
  s.name        = "docx_converter"
  s.version     = DocxConverter::VERSION
  s.authors     = ["Michael Franzl"]
  s.email       = ["office@michaelfranzl.com"]
  s.homepage    = ""
  s.summary     = %q{Converts docx files into html or LaTeX via the kramdown syntax}
  s.description = %q{Converts docx files into html or LaTeX via the kramdown syntax. It supports Word's most common paragraph, character and mixed styles (Title, Heading, Strong, Quote), footnotes, citations, line breaks, page breaks, non-breaking spaces and images. The output is in kramdown syntax (see http://kramdown.gettalong.org/) which can be converted into beautiful html and LaTex code.} 
  s.rubyforge_project = "docx_converter"

  s.files         = `git ls-files`.split("\n")
  s.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")
  s.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ["lib"]

  # specify any dependencies here; for example:
  s.add_runtime_dependency "kramdown"
  s.add_runtime_dependency "nokogiri"
end
