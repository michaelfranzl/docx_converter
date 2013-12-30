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
  class PostProcessor
    
    def initialize(content, footnote_definitions)
      @content = content
      @footnote_definitions = footnote_definitions
      
      @chapters = []
    end
    
    # getter method
    def chapters
      return @chapters
    end
    
    def join_blockquotes
      lines = @content.split("\n")
      processed_lines = []
      lines.size.times do |i|
        if /^> /.match(lines[i-1]) && /^> /.match(lines[i+1])
          processed_lines << ">" + lines[i]
        else
          processed_lines << lines[i]
        end
      end
      @content = processed_lines.join("\n")
      @chapters[0] = @content
      return @content
    end

    def split_into_chapters
      chapter_number = 0
      @chapters[chapter_number] = ""
      @content.split("\n").each do |line|
        if /^# /.match(line)
          # this is the style Heading1. A new chapter begins here.
          chapter_number += 1
          @chapters[chapter_number] = ""
        end
        @chapters[chapter_number] += line + "\n"
      end
      return @chapters
    end
    
    def add_foonote_definitions
      @chapters.size.times do |n|
        footnote_ids = @chapters[n].scan(/\[\^(.+?)\]/).flatten
        @chapters[n] += "\n\n"
        footnote_ids.each do |i|
          @chapters[n] += "[^#{ i }]: #{ @footnote_definitions[i] }\n\n"
        end
      end
    end
   
  end
end