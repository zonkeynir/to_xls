require 'rubygems'
require 'spreadsheet'
require 'stringio'

class Hash
  def to_xls(options = {})
    book = Spreadsheet::Workbook.new

    each do |key, value|
      sheet = book.create_worksheet
      sheet.name = key.to_s
      value.to_xls({ :sheet => sheet }.merge(options[key] || {}))
    end
    
    return book
  end
  
  def to_xls_data(options = {})
    data = StringIO.new('')
    self.to_xls(options).write(data)
    return data.string
  end
end

class Array
  # Options for to_xls: columns, name, header, sheet
  def to_xls(options = {})
    sheet = options[:sheet]
    unless sheet
      book = Spreadsheet::Workbook.new
      sheet = book.create_worksheet
      sheet.name = options[:name] || 'Sheet 1'
    end

    if self.any?
      columns = get_columns(options)

      if columns.is_a?(Array) && columns.any?
        line = 0
        
        unless options[:headers] == false
          if options[:headers].is_a?(Array)
            sheet.row(0).concat options[:headers].collect(&:to_s)
          else
            aux_headers_to_xls(self.first, columns, sheet.row(0))
          end
          line = 1
        end
        
        self.each do |item|
          row = sheet.row(line)
          columns.each {|column| aux_to_xls(item, column, row)}
          line += 1
        end
      end
    end

    return book || sheet
  end
  
  def to_xls_data(options = {})
    data = StringIO.new('')
    self.to_xls(options).write(data)
    return data.string
  end
  
  private  

  def get_columns(options)
    columns = options[:columns]
    raise ArgumentError, ":columns (#{columns}) must be an array or nil" unless (columns.nil? || columns.is_a?(Array))
    if !columns && can_get_columns_from_first_element?
      columns = get_columns_from_first_element
    end
    return columns
  end

  def can_get_columns_from_first_element?
    self.first && 
    self.first.respond_to?(:attributes) &&
    self.first.attributes.respond_to?(:keys) &&
    self.first.attributes.keys.is_a?(Array)
  end

  def get_columns_from_first_element
    self.first.attributes.keys.sort
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
  
end
