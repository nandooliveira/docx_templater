# frozen_string_literal: true

module Docx
  class NodesToFix
    attr_accessor :node_list, :current_node, :value
    def initialize
      forget
    end

    def forget
      @current_node = nil
      @node_list = []
      @value = ''
    end

    def remember(node, index)
      new_node = current_node.nil? || current_node != node
      if new_node
        @current_node = node
        @node_list << { node: node, range: index..index }
      else
        @node_list.last[:range] = (node_list.last[:range].min)..index
      end
    end

    def fix
      @node_list.each do |obj|
        node = obj[:node]
        range = obj[:range]

        new_val = node.value

        key = obj[:node].to_s[range].gsub('||', '')

        if key.start_with?('**') && key.end_with?('**')
          # the text before bolded tag
          first_part = node.value[0..(range.first - 1)]
          node.parent.parent.parent.add_element(common_element(first_part))

          # bolded tag
          node.parent.parent.parent.add_element(bold_element(value.to_s || ''))

          # the text after the bolded tag
          last_part = node.value.to_s[(range.last + 1)..-1]
          node.parent.parent.parent.add_element(common_element(last_part))

          node.parent.delete(node)
        elsif key&.include?('logo_cliente')
          # the text before logo image
          first_part = node.value[0..(range.first - 1)]
          node.parent.parent.parent.add_element(common_element(first_part))

          # image tag
          tags_to_include = %w[body hdr]
          parent = first_parent_of_type(node, tags_to_include)
          parent_elements = parent.elements
          reversed_elements = parent_elements.to_a.reverse
          reversed_elements << image_element if tags_to_include.include?(parent.name)

          reversed_elements.reverse.each_with_index do |e, index|
            next if e == node.parent

            parent_elements[index + 1] = e
          end

          # the text after the logo image
          last_part = node.value.to_s[(range.last + 1)..-1]
          node.parent.parent.parent.add_element(common_element(last_part))

          node.parent.parent.delete(node.parent)
        elsif value&.include?('**')
          tokenized_value = value.split('**')
          tokenized_value.each do |token|
            element = if token.start_with?('[') && token.end_with?(']')
                        # come here this way '||**name already changed**||'
                        token = token&.gsub('[', '')&.gsub(']', '')
                        bold_element(token.to_s || '')
                      else
                        common_element(token)
                      end

            node.parent.parent.parent.add_element(element)
          end

          node.parent.delete(node)
        else
          new_val[range] = value.to_s || ''
          node.value = new_val

          if new_val =~ /^\s+/ && node.parent
            node.parent.add_attribute('xml:space', 'preserve')
          end
        end

        self.value = nil
      end
    end

    def first_parent_of_type(node, element_types)
      puts "===>> #{node.parent.name}" unless node.parent.nil?
      return node        if node.parent.nil?
      return node.parent if element_types.include?(node.parent.name)

      first_parent_of_type(node.parent, element_types)
    end

    def common_element(content)
      wr_element   = REXML::Element.new('w:r')
      wrpr_element = REXML::Element.new('w:rPr')
      wt_element   = REXML::Element.new('w:t')
      wt_element.add_text(content)

      wr_element.add_element(wrpr_element)
      wr_element.add_element(wt_element)

      wt_element.add_attribute('xml:space', 'preserve')

      wr_element
    end

    def bold_element(new_value)
      wr_element   = REXML::Element.new('w:r')
      wrpr_element = REXML::Element.new('w:rPr')
      wb_element   = REXML::Element.new('w:b')
      wbcs_element = REXML::Element.new('w:bCs')
      wt_element   = REXML::Element.new('w:t')

      value = new_value.gsub('\*\*', '')
      wt_element.add_text(value)

      wrpr_element.add_element(wb_element)
      wrpr_element.add_element(wbcs_element)
      wr_element.add_element(wrpr_element)
      wr_element.add_element(wt_element)

      wt_element.add_attribute('xml:space', 'preserve')

      wr_element
    end

    def image_element
      wp_element    = REXML::Element.new('w:p')
      wr_element    = REXML::Element.new('w:r')
      wpict_element = REXML::Element.new('w:pict')

      vshape_element = REXML::Element.new('v:shape')
      vshape_element.add_attribute('id', 'logo_cliente')
      vshape_element.add_attribute('type', '#_x0000_t75')
      vshape_element.add_attribute('style', 'width:70; height:70; border: 10; float: left')

      vimagedata_element = REXML::Element.new('v:imagedata')
      vimagedata_element.add_attribute('r:id', 'logo_cliente')

      vshape_element.add_element(vimagedata_element)
      wpict_element.add_element(vshape_element)
      wr_element.add_element(wpict_element)
      wp_element.add_element(wr_element)

      wp_element
    end
  end
end
