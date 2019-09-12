# require 'zip/zipfilesystem'
require 'zip'
require 'htmlentities'
require 'docx/argument_combiner'
require 'docx/document_replacer'
require 'docx/newline_replacer'

# Use .docx as reusable templates
#
# Example:
# buffer = DocxTemplater.replace_file_with_content('path/to/mydocument.docx',
#    {
#      :client_email1 => 'test@example.com',
#      :client_phone1 => '555-555-5555',
#    })
# # In Rails you can send a word document via send_data
# send_data buffer.string, :filename => 'REPC.docx'
# # Or save the output to a word file
# File.open("path/to/mydocument.docx", "wb") {|f| f.write(buffer.string) }
class DocxTemplater
  def initialize(opts = {})
    @options = opts
  end

  def replace_file_with_content(file_path, data_provider, logo_url)
    # Rubyzip doesn't save it right unless saved like this: https://gist.github.com/e7d2855435654e1ebc52
    zf = Zip::File.new(file_path) # Put original file name here

    buffer = Zip::OutputStream.write_buffer do |out|
      zf.entries.each do |e|
        process_entry(e, out, data_provider, logo_url)
      end
    end
    # You can save this buffer or send it with rails via send_data
    return buffer
  end

  def generate_tags_for(*args)
    Docx::ArgumentCombiner.new(*args).attributes
  end

  def entry_requires_replacement?(entry)
    entry.ftype != :directory && entry.name =~ /document|header|footer/
  end

  private
  attr_reader :options

  def get_entry_content(entry, data_provider)
    file_string = entry.get_input_stream.read
    if entry_requires_replacement?(entry)
      replacer = Docx::DocumentReplacer.new(file_string, data_provider, options)
      replacer.replaced
    else
      file_string
    end
  end

  def process_entry(entry, output, data_provider, logo_url)
    output.put_next_entry(entry.name)
    entry_content = REXML::Document.new(get_entry_content(entry, data_provider))

    if logo_url != nil
      if entry.name == '[Content_Types].xml'
        entry_content.elements[1].add_element(image_extension('png'))
        entry_content.elements[1].add_element(image_extension('jpg'))
        entry_content.elements[1].add_element(image_extension('jpeg'))
      end

      entry_content.elements[1].add_element(relationship_element(logo_url)) \
        if entry.name.end_with?('document.xml.rels')

      iterate_searching_for_image_tag(entry_content.elements[1][0].elements) \
        if entry.name.end_with?('document.xml')
    end

    output.write entry_content.to_s if entry.ftype != :directory
  end

  def image_extension(extension)
    default_extension = REXML::Element.new('Default')
    default_extension.add_attribute('Extension', extension)
    default_extension.add_attribute('ContentType', "image/#{extension}")

    default_extension
  end

  def relationship_element(logo_url)
    rel_element = REXML::Element.new('Relationship')
    rel_element.add_attribute('Id', 'customer_logo')
    rel_element.add_attribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
    rel_element.add_attribute('Target', logo_url)
    rel_element.add_attribute('TargetMode', 'External')

    rel_element
  end

  def iterate_searching_for_image_tag(elements)
    elements.each do |element|
      new_elements = element.elements
      begin
        if element.text.to_s.include?('||customer_logo||')
          parent_elements = element.parent.parent.parent.elements
          reversed_elements = parent_elements.to_a.reverse
          reversed_elements << image_element


          reversed_elements.reverse.each_with_index do |e, index|
            parent_elements[index + 1] = e
          end

          element.parent.delete(element)
        end
      rescue StandardError => e
        pp e.backtrace
      end

      iterate_searching_for_image_tag(new_elements) unless new_elements.empty?
    end
  end

  def image_element
    wp_element    = REXML::Element.new('w:p')
    wr_element    = REXML::Element.new('w:r')
    wpict_element = REXML::Element.new('w:pict')

    vshape_element = REXML::Element.new('v:shape')
    vshape_element.add_attribute('id', 'customer_logo')
    vshape_element.add_attribute('type', '#_x0000_t75')
    vshape_element.add_attribute('style', 'width:100; height:100; border: 10; float: left')

    vimagedata_element = REXML::Element.new('v:imagedata')
    vimagedata_element.add_attribute('r:id', 'customer_logo')

    vshape_element.add_element(vimagedata_element)
    wpict_element.add_element(vshape_element)
    wr_element.add_element(wpict_element)
    wp_element.add_element(wr_element)

    wp_element
  end
end
