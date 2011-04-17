require 'spreadsheet'

module ToXls

  class ArrayRenderer
    def initialize(array, options)
      @array = array
      @options = options
    end

    def render
      fill_sheet
      return @book || @sheet
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
      @columns = @columns || []
    end

    def can_get_columns_from_first_element?
      @array.first && 
      @array.first.respond_to?(:attributes) &&
      @array.first.attributes.respond_to?(:keys) &&
      @array.first.attributes.keys.is_a?(Array)
    end

    def get_columns_from_first_element
      @array.first.attributes.keys.sort_by {|sym| sym.to_s}.collect
    end

    def headers
      return  @headers if @headers
      @headers = @options[:headers]
      if @headers
        raise ArgumentError, ":headers (#{@headers}) must be an array" unless @headers.is_a? Array
      else
        @headers = columns
      end
      @headers
    end

    def headers_should_be_included?
      @options[:headers] != false
    end

private

  def fill_sheet
    if columns.any?
      row_index = 0

      if headers_should_be_included?
        fill_row(sheet.row(0), headers)
        row_index = 1
      end

      @array.each do |model|
        row = sheet.row(row_index)
        fill_row(row, columns, model)
        row_index += 1
      end
    end
  end

  def fill_row(row, column, model=nil)
    case column
    when String, Symbol
      row.push(model ? model.send(column) : column)
    when Hash
      column.each{|key, values| fill_row(row, values, model && model.send(key))}
    when Array
      column.each{|value| fill_row(row, value, model)}
    else
      raise ArgumentError, "column #{column} has an invalid class (#{ column.class })"
    end
  end

  end
end
