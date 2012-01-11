
require 'to_xls/writer.rb'

module Enumerable
  # Options for to_xls: columns, name, header, sheet
  def to_xls(options = {})
    ToXls::Writer.new(self, options).write_string
  end
end
