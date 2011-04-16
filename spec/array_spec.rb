require File.expand_path(File.dirname(__FILE__) + '/spec_helper')

describe Array do

  def mock_user(name, age, email)
    attributes = { 'name' => name, 'age' => age, 'email' => email }
    attributes[:attributes] = attributes.clone
    mock(name, attributes)
  end

  def mock_users
    [ mock_user('Peter', 20, 'peter@gmail.com'),
      mock_user('John',  25, 'john@gmail.com'),
      mock_user('Day9',  27, 'day9@day9tv.com')
    ]
  end

  def check_sheet(sheet, array)
    sheet.rows.each_with_index do |row, i|
      row.should == array[i]
    end
  end

  it "should throw no error without data" do
    lambda { [].to_xls }.should_not raise_error
  end
  it "should throw no error without columns" do
    lambda { [1,2,3].to_xls }.should_not raise_error
  end
  it "should throw an error if columns exists but it isn't an array" do
    lambda { [1,2,3].to_xls(:columns => :foo) }.should raise_error
  end
  it "should use the attribute keys as columns if it exists" do
    xls = mock_users.to_xls
    check_sheet( xls.worksheets.first,
      [ ['age', 'email',           'name'],
        [   20, 'peter@gmail.com', 'Peter'],
        [   25, 'john@gmail.com',  'John'],
        [   27, 'day9@day9tv.com', 'Day9']
      ]
    )
  end

  it "should work properly when you provide it with both data and column names" do
    xls = [1,2,3].to_xls(:columns => [:to_s])
    check_sheet( xls.worksheets.first, [ ['to_s'], ['1'], ['2'], ['3'] ] )
  end
end
