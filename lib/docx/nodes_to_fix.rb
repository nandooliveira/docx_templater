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
        elsif new_val.include?('**')

          # # bolded tags
          # bolded_tags = data[:contratada].split('||').select { |a| a.start_with?('**') && a.end_with?('**') }.uniq
          # bolded_tags.each do |tag|
          #   cleaned_tag = tag.gsub('**', '')
          #   next unless ::Document::HIRED_TAGS.include?(cleaned_tag.to_sym)

          #   value = hired.send(cleaned_tag)
          #   data[cleaned_tag.to_sym] = "<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>#{value}</w:t></w:r>"
          #   data[:contratada] = data[:contratada]&.gsub(
          #     "||**#{cleaned_tag}**||",
          #     "<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>#{value}</w:t></w:r>"
          #   )
          # end
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
  end
end
