module Excelicious
  @styles = {}
  # maybe make some default styles as well
  @locations = {} # do not know if this will be used

  def load_styles @user_styles
    @user_styles.each do |style_name, style|
      @styles[style_name] = Spreadsheet::Format.new style
    end
  end
  
  def apply_styles @locations
    @locations.each do |location, style|
      if location.is_a?(Integer)
        sheet1.row(location).default_format = @styles[style] # example
      else
        sheet1.column(location.to_column_number).default_format = @styles[style] # example
      end
    end
  end
  
  def to_column_number
    # need a way to turn the user inputted column names into column numbers
    @data_array.each_with_index do |index, name|
      
    end
  end
  
end

@data = {
  'user_one' => {
    'email' => 'blah@blah.com',
    'name' => "blah"
  }
}

#Example user input
@user_styles = {
  'blue_bold_big' => {
    :color => :blue,
    :weight => :bold,
    :size => 30
  }
}

# note that the user can either specify a column name they can also specify a row by number
# it may be a design flaw, but for now we are going to assume the user enters columns as strings and rows by number
# these can be typechecked like so: 3.is_a?(Integer)
@locations = {
  'email' => 'blue_bold_big',
  3 => 'blue_bold_big'
}
@lo = {
  :rows => [ ..... ],
  :columns => [ .... ]
}

# combined input into one hash
@styles = {
    :locations => {
        :rows => {
            0 => 'blue_bold_big',
            10 => 'blue_bold_big'
        },
        :columns => { # for now, let's stay away from using column names. stick to numbers
            1 => 'blue_bold_big',
            2 => 'blue_bold_big'
        }
    },
    :styles => {
        'blue_bold_big' => {
            :color => :blue,
            :weight => :bold,
            :size => 30
        }
    }
}