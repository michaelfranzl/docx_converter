# encoding: UTF-8

# docx_converter -- Converts docx files into html or LaTeX via the kramdown syntax
# Copyright (C) 2013 Red (E) Tools Ltd. (www.thebigrede.net)
# 
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as
# published by the Free Software Foundation, either version 3 of the
# License, or (at your option) any later version.
# 
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

module DocxConverter
  class Parser
    def initialize(docx_filepath, output_dir, options)
      @output_dir = output_dir
      @docx_filepath = docx_filepath
      
      @image_subdir_filesystem = options[:image_subdir_filesystem]
      @image_subdir_kramdown = options[:image_subdir_kramdown]
      FileUtils.mkdir_p(File.join(@output_dir, @image_subdir_filesystem))
      
      @footnotes_hash = {}
      @relationships_hash = {}
      
      @image_extract_dir = File.dirname(docx_filepath)
      
      @zipfile = Zip::ZipFile.new(@docx_filepath)
      
      docx = nil
    end
    
    def parse
      document_xml = unzip_read("word/document.xml")
      footnotes_xml = unzip_read("word/footnotes.xml")
      relationships_xml = unzip_read("word/_rels/document.xml.rels")
      
      content = Nokogiri::XML(document_xml)
      footnotes = Nokogiri::XML(footnotes_xml)
      relationships = Nokogiri::XML(relationships_xml)
      
      @relationships_hash = parse_relationships(relationships)
      
      puts "Relationships are"
      puts @relationships_hash.inspect

      @footnotes_hash = parse_footnotes(footnotes)
      puts "Footnotes are:"
      puts @footnotes_hash.inspect

      debugger
      output = parse_content(content.elements.first,0)
      output += compose_footnote_definitions
      
      return output
    end
    
    private
    
    def unzip_read(path)
      file = @zipfile.find_entry(path)
      contents = ""
      file.get_input_stream do |f|
        contents = f.read
      end
      return contents
    end
    
    def unzip_extract(zip_path, extract_filepath)
      extract_filepath_full = File.join(@output_dir, extract_filepath)
      file = @zipfile.find_entry(zip_path)
      FileUtils.rm_f(extract_filepath_full)
      ret = file.extract(extract_filepath_full)
      return ret
    end

    def parse_relationships(relationships)
      output = {}
      relationships.children.first.children.each do |rel|
        rel_id = rel.attributes["Id"].value
        rel_target = rel.attributes["Target"].value
        output[rel_id] = rel_target
      end
      return output
    end
    
    def parse_footnotes(node)
      output = {}
      node.xpath("//w:footnote").each do |fnode|
        footnote_number = fnode.attributes["id"].value
        if ["-1", "0"].include?(footnote_number)
          # Word outputs -1 and 0 as 'magic' footnotes
          next
        end
        output[footnote_number] = parse_content(fnode,0).strip
      end
      return output
    end


    def compose_footnote_definitions
      output = ""
      @footnotes_hash.each do |k,v|
        output += "[^#{ k }]: #{ v }\n\n"
      end
      return output
    end


    def parse_content(node,depth)
      output = ""
      depth += 1
      children_count = node.children.length
      i = 0
      
      while i < children_count
        add = ""
        nd = node.children[i]
        
        case nd.name
        when "body"
          # This is just a container element.
          add = parse_content(nd,depth)
          
        when "document"
          # This is just a container element.
          add = parse_content(nd,depth)
          
        when "p"
          # This is a paragraph. In kramdown, paragraphs are spearated by an empty line.
          add = parse_content(nd,depth) + "\n\n"
          
        when "pPr"
          # This is Word's paragraph-level preset
          add = parse_content(nd,depth)
          
        when "pStyle"
          # This is a reference to one of Word's paragraph-level styles
          case nd["w:val"]
            when "Title"
              add = "{: .class = 'title' }\n# "
            when "Heading1"
              add = "# "
            when "Heading2"
              add = "## "
          end
            
        when "r"
          # This corresponds to Word's character/inline node. Word's XML is not nested for formatting, wo we cannot descend recursively and 'close' kramdown's formatting in the recursion. Rather, we have to look ahead if this node is formatted, and if yes, set a formatting prefix and postfix which is required for kramdown (e.g. **bold**).
          prefix = postfix = ""
          first_child = nd.children.first
          
          case first_child.name
          when "rPr"
            # This inline node is formatted. The first child always specifies the formatting of the subsequent 't' (text) node.
            format_node = first_child.children.first
            case format_node.name
            when "b"
              # This is regular (non-style) bold
              prefix = postfix = "**"
            when "i"
              # This is regular (non-style) italic
              prefix = postfix = "*"
            when "rStyle"
              # This is a reference to one of Word's style names
              case format_node.attributes["val"].value
              when "Strong"
                # "Strong" is a predefined Word style
                # This node is missing the xml:space="preserve" attribute, so we need to set the spaces ourselves.
                prefix = " **"
                postfix = "** "
              end
            end
            add = prefix + parse_content(nd,depth) + postfix
          when "br"
            if first_child.attributes.empty?
              # This is a line break. In kramdown, this corresponds to two spaces followed by a newline.
              add = "  \n"
            else first_child.attributes["type"] == "page"
              # this is a Word page break
              add = "<br style='page-break-before:always;'>"
            end
            
          else
            add = parse_content(nd,depth)
          end
          
            
        when "t"
          # this is a regular text node
          add = nd.text
          
        when "footnoteReference"
          # output the Kramdown footnote syntax
          footnote_number = nd.attributes["id"].value
          add = "[^#{ footnote_number }]"
          
        when "tbl"
          # parse the table recursively
          add = parse_content(nd,depth)
          
        when "tr"
          # select all paragraph nodes below the table row and render them into Kramdown syntax
          table_paragraphs = nd.xpath(".//w:p")
          td = []
          table_paragraphs.each do |tp|
            td << parse_content(tp,depth)
          end
          add = "|" + td.join("|") + "|\n"
          
        when "drawing"
          image_nodes = nd.xpath(".//a:blip", :a => 'http://schemas.openxmlformats.org/drawingml/2006/main')
          image_node = image_nodes.first
          image_id = image_node.attributes["embed"].value
          image_path_zip = File.join("word", @relationships_hash[image_id])
          
          image_extract_path_filesystem = File.join(@image_subdir_filesystem, File.basename(image_path_zip))
          image_extract_path_kramdown = File.join(@image_subdir_filesystem, File.basename(image_path_zip))
          image_extract_name_kramdown = File.basename(image_path_zip)
            
          add = "![]{#{ image_extract_name_kramdown }}\n"
          
          unzip_extract(image_path_zip, image_extract_path_filesystem)
          
        when "proofErr", "rPr"
          # ignore those nodes, they don't have a correspondence in kramdown.
        else
          puts ' ' * depth + "ELSE: #{ nd.name }"
        end
        
        output += add
        i += 1
      end

      #puts "#{ nd.name }: #{ text.inspect }\n-----\n"
      depth -= 1
      
      return output
    end
    
  end
end
