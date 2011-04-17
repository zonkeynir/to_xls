
module ToXls

  class ArrayRenderer
    def initialize(array, options)
      @array = array
      @options = options
    end

    def render
      fill_up_sheet
      return @book || @sheet
    end

    def fill_up_sheet
      if columns.any?
        row_index = 0

        if headers_should_be_included?
          sheet.row(0).concat headers
          row_index = 1
        end

        @array.each do |item|
          row = sheet.row(row_index)
          columns.each {|column| aux_to_xls(item, column, row)}
          row_index += 1
        end
      end
    end

    def sheet
      return @sheet if @sheet
      @sheet = @options[:sheet]
      unless @sheet
        @book = Spreadsheet::Workbook.new
        @sheet = @book.create_worksheet
        @sheet.name = @options[:name] || 'Sheet 1'
      end
      @sheet
    end

    def columns
      return  @columns if @columns
      @columns = @options[:columns]
      raise ArgumentError, ":columns (#{columns}) must be an array or nil" unless (@columns.nil? || @columns.is_a?(Array))
      if !@columns && can_get_columns_from_first_element?
        @columns = get_columns_from_first_element
      end
      @columns = (@columns || []).collect{|c| c.to_s}
    end

    def can_get_columns_from_first_element?
      @array.first && 
      @array.first.respond_to?(:attributes) &&
      @array.first.attributes.respond_to?(:keys) &&
      @array.first.attributes.keys.is_a?(Array)
    end

    def get_columns_from_first_element
      @array.first.attributes.keys.sort.collect
    end

    def headers
      return  @headers if @headers
      @headers = @options[:headers]
      if @headers
        raise ArgumentError, ":headers (#{@headers}) must be an array or nil" unless @headers.is_a? Array
        @headers = @headers.collect{|h| h.to_s}
      else
        @headers = columns
      end
      @headers
    end

    def headers_should_be_included?
      @options[:headers] != false
    end


    def aux_headers_to_xls(item, column, row)
      if item.nil?
        row.push(nil)
      elsif column.is_a?(String) or column.is_a?(Symbol)
        row.push(column.to_s)
      elsif column.is_a?(Hash)
        column.each{|key, values| aux_headers_to_xls(item.send(key), values, row)}
      elsif column.is_a?(Array)
        column.each{|value| aux_headers_to_xls(item, value, row)}
      end
    end

    def aux_to_xls(item, column, row)
      if item.nil?
        row.push(nil)
      elsif column.is_a?(String) or column.is_a?(Symbol)
        row.push(item.send(column))
      elsif column.is_a?(Hash)
        column.each{|key, values| aux_to_xls(item.send(key), values, row)}
      elsif column.is_a?(Array)
        column.each{|value| aux_to_xls(item, value, row)}
      end
    end

  end
end
