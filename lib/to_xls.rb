require 'rubygems'
require 'spreadsheet'
require 'stringio'
require 'to_xls/array_renderer.rb'

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

    renderer = ToXls::ArrayRenderer.new(self, options)
    return renderer.render
  end
  
  def to_xls_data(options = {})
    data = StringIO.new('')
    self.to_xls(options).write(data)
    return data.string
  end

end
