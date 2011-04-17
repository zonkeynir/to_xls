require 'rubygems'
require 'stringio'
require 'to_xls/array_renderer.rb'

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
