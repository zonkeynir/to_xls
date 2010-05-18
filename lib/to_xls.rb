require 'rubygems'
require 'spreadsheet'
require 'stringio'

class Array
  # Options for to_xls: columns, name, header
  def to_xls(options = {})
    
    book = Spreadsheet::Workbook.new
    sheet = book.create_worksheet
    
    sheet.name = options[:name] || 'Sheet 1'    

    if self.any?
      columns = options[:columns] || self.first.attributes.keys.sort

      if columns.any?
        line = 0
        
        unless options[:headers] == false
          if options[:headers].is_a?(Array)
            sheet.row(0).concat options[:headers]
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

    return book
  end
  
  def to_xls_data(options = {})
    data = StringIO.new('')
    self.to_xls(options).write(data)
    return data.string
  end
  
  private  
  def aux_to_xls(item, column, row)
    if column.is_a?(String) or column.is_a?(Symbol)
      row.push(item.send(column))
    elsif column.is_a?(Hash)
      column.each{|key, values| aux_to_xls(item.send(key), values, row)}
    elsif column.is_a?(Array)
      column.each{|value| aux_to_xls(item, value, row)}
    end
  end
  
  def aux_headers_to_xls(item, column, row)
    if column.is_a?(String) or column.is_a?(Symbol)
      row.push("#{item.class.name.underscore}_#{column}")
    elsif column.is_a?(Hash)
      column.each{|key, values| aux_headers_to_xls(item.send(key), values, row)}
    elsif column.is_a?(Array)
      column.each{|value| aux_headers_to_xls(item, value, row)}
    end
  end
  
end
