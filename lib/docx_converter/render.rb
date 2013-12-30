# encoding: UTF-8

# docx_converter -- Converts Word docx files into html or LaTeX via the kramdown syntax
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
  class Render
    
    def initialize(options)
      @options = options
      @output_dir = options[:output_dir]
      
      @chapters = nil
    end
    
    def render(output_format)
      FileUtils.mkdir_p(File.join(@output_dir, @options[:image_subdir_filesystem]))
      
      output = DocxConverter::Parser.new(@options).parse

      content = output[:content]
      footnote_definitions = output[:footnote_definitions]
      
      p = DocxConverter::PostProcessor.new(content, footnote_definitions)
      p.join_blockquotes

      if @options[:split_chapters] == true
        p.split_into_chapters
      end
      
      p.add_foonote_definitions
      @chapters = p.chapters

      case output_format
      when :kramdown
        filepaths = render_kramdown
      when :html
        filepaths = render_html
      when :latex
        filepaths = render_latex
      else
        filepaths = []
      end
      return filepaths
    end
    
    private
    
    def render_kramdown
      # output is in .page file extension. this is merely a convetion taken from the webgen file system structure
      rendered_kramdown_file_paths = []
      if @options[:split_chapters] == true
        @chapters.size.times do |n|
          filename = "%02i.chapter%02i.%s.page" % [n, n, @options[:language]]
          file_path = File.join(@output_dir, filename)
          File.write(file_path, @chapters[n])
          rendered_kramdown_file_paths << filename
        end
      else
        filename = "01.chapter01.%s.page" % [@options[:language]]
        file_path = File.join(@output_dir, filename)
        File.write(file_path, @chapters[0])
        rendered_kramdown_file_paths << filename
      end
      return rendered_kramdown_file_paths
    end
    
    def render_html
      rendered_kramdown_file_paths = render_kramdown
      rendered_html_file_paths = []
      rendered_kramdown_file_paths.each do |kfp|
        filename = kfp.gsub("page", "html")
        file_path = File.join(@output_dir, filename)
        kramdown = File.read(File.join(@output_dir, kfp))
        html = Kramdown::Document.new(
          kramdown,
          :input => 'kramdown',
          :line_width => 100000
        ).to_html
        File.write(file_path, html)
        rendered_html_file_paths << filename
      end
      return rendered_html_file_paths
    end
    
    def render_latex
      rendered_kramdown_file_paths = render_kramdown
      rendered_latex_file_paths = []
      rendered_kramdown_file_paths.each do |kfp|
        filename = kfp.gsub("page", "tex")
        file_path = File.join(@output_dir, filename)
        kramdown = File.read(File.join(@output_dir, kfp))
        latex = Kramdown::Document.new(
          kramdown,
          :input => 'kramdown',
          :line_width => 100000
        ).to_latex
        File.write(file_path, latex)
        rendered_latex_file_paths << filename
      end
      return rendered_latex_file_paths
    end

  end
end