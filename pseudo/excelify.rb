# Users provide us...
Excelify::GO({@users}, {@styles}^, {@style_locations}^)

# @users is just a typical database hash from Activerecord

# styles is a hash specifiying styles based on MS Excel XML API

# style_locations is a hash specifying where the style should be applied

#----- Notes ------- #
# We need to typecheck data to make sure numbers are defined as numbers (by excel) and not strings
  # example: if (typeof cell != number) type = 'String' else type = 'Number'
  # xml.Data column.human_name, 'ss:Type' => type

# Method(s)

module Excelify
  def make_cell_type(value, type)
    # excel mumbo jumbo here
    xml.Data record.send(value), 'ss:Type' => 'String'
  end
  
  def set_style (style)
    'ss:StyleID'=>style
  end
  
  # make a new style for each user style
  def set_user_styles @styles
    @styles.each do |style|
    xml.Style 'ss:ID'=> style.id do
      xml.Font 'ss:Bold'=>'1' , 'ss:Color'=>'#ff0000', 'ss:Size'=>'30'
      xml.Borders do
        xml.Border 'ss:Position'=>'Bottom' , 'ss:Color'=>'#0000ff', 'ss:LineStyle'=>'DashDotDot','ss:Weight'=>'3'
      end
    end
  end
  end
  
  def self.GO @datas, @styles, @style_locations
    @datas.each do |data|
      xml.Row do
        data.attributes.each_pair do |prop, value|
          xml.Cell set_style(style) do
            make_cell_type(value, type)
          end
        end
      end
    end
  end  
  
end



# Example excel lingo

for record in @users
  xml.Row  do # you could highlight the whole row here: 'ss:StyleID'=>'hightlighRow'
    for column in User.content_columns do
      xml.Cell 'ss:StyleID'=>'hightlightRow' do # i chose to limit the hightlight to the cell
        xml.Data record.send(column.name), 'ss:Type' => 'String'
      end
    end
  end
end