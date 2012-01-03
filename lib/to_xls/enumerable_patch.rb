
require 'to_xls/enumerable_writer.rb'

module Enumerable
  # Options for to_xls: columns, name, header, sheet
  def to_xls(options = {})
    ToXls::EnumerableWriter.new(self, options).write_string
  end
end
