require 'rubygems'
require 'stringio'
require 'spreadsheet'
require 'set'

module ToXls

  class ArrayWriter
    def initialize(array, options = {})
      @array = array
      @options = options
    end

    def apply_styles (options, sheet)
      if options.has_key?(:style)
        if options[:style].has_key?(:locations) && options[:style].has_key?(:styles)
          styles = register_styles(@options[:style][:styles])
          render_styles(styles, sheet)
        else
          raise ArgumentError, "To_xls error => :style was included in options, but it does not contain the proper keys (:styles and :locations)"
        end
      end
    end

    def register_styles(_styles)
      styles = {}
      _styles.each do |name, style|
        styles[name] = Spreadsheet::Format.new style
      end
      styles
    end

    def render_styles(styles, sheet)
      locations  = @options[:style][:locations]
      locations[:rows].each do |row, style|
        sheet.row(row).default_format = styles[style]
      end
      locations[:columns].each do |column, style|

        if(column.is_a? Integer)
          sheet.column(column).default_format = styles[style]
        else
          sheet.column(find_column(column)).default_format = styles[style]
        end
      end
    end

    def find_column(column_name)
      column_names = headers

      if(column_names)
        if column_names[0].is_a?(String)
          column_index = column_names.index(column_name.to_s)
          if column_index
            return column_index
          else
            raise ArgumentError, "The inputted column name does not exist in provided data"
          end
        elsif column_names[0].is_a?(Symbol)
          column_index = column_names.index(column_name)
          if column_index
            return column_index
          else
            raise ArgumentError, "The inputted column name does not exist in provided data"
          end
        end
      else
        raise "No headers specified or found"
      end

    end

    def write_string(string = '')
      io = StringIO.new(string)
      write_io(io)
      io.string
    end

    def write_io(io)
      book = Spreadsheet::Workbook.new
      write_book(book)
      book.write(io)
    end

    def write_book(book)
      sheet = book.create_worksheet
      sheet.name = @options[:name] || 'Sheet 1'
      write_sheet(sheet)
      return book
    end

    def write_sheet(sheet)
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

        apply_styles(@options, sheet)
      end
    end

    def columns
      return  @columns if @columns
      @columns = @options[:columns]
      raise ArgumentError.new(":columns (#{columns}) must be an array or nil") unless (@columns.nil? || @columns.is_a?(Array))
      @columns ||=  can_get_columns_from_first_element? ? get_columns_from_first_element : []
    end

    def can_get_columns_from_first_element?
      @array.first &&
      @array.first.respond_to?(:attributes) &&
      @array.first.attributes.respond_to?(:keys) &&
      @array.first.attributes.keys.is_a?(Array)
    end

    def get_columns_from_first_element
      @array.first.attributes.keys.sort_by {|sym| sym.to_s}.collect.to_a
    end

    def headers
      return  @headers if @headers
      @headers = @options[:headers] || columns
      raise ArgumentError, ":headers (#{@headers.inspect}) must be an array" unless @headers.is_a? Array
      @headers
    end

    def headers_should_be_included?
      @options[:headers] != false
    end

private

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
