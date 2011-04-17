require File.expand_path(File.dirname(__FILE__) + '/spec_helper')

describe Array do


  def mock_model(name, attributes)
    attributes[:attributes] = attributes.clone
    mock(name, attributes)
  end

  def mock_company(name, address)
    mock_model( name, 'name' => name, 'address' => address )
  end

  def mock_user(name, age, email, company)
    user = mock_model(name, 'name' => name, 'age' => age, 'email' => email)
    user.stub!(:company).and_return(company)
    user
  end

  def mock_users
    acme = mock_company('Acme', 'One Road')
    eads = mock_company('EADS', 'Another Road')

    [ mock_user('Peter', 20, 'peter@gmail.com', acme),
      mock_user('John',  25, 'john@gmail.com', acme),
      mock_user('Day9',  27, 'day9@day9tv.com', eads)
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


  describe ":columns option" do
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
    it "should allow re-sorting of the columns by using the :columns option" do
      xls = mock_users.to_xls(:columns => [:name, :email, :age])
      check_sheet( xls.worksheets.first,
        [ ['name',  'email',        'age'],
          ['Peter', 'peter@gmail.com', 20],
          ['John',  'john@gmail.com',  25],
          ['Day9',  'day9@day9tv.com', 27]
        ]
      )
    end

    it "should work properly when you provide it with both data and column names" do
      xls = [1,2,3].to_xls(:columns => [:to_s])
      check_sheet( xls.worksheets.first, [ ['to_s'], ['1'], ['2'], ['3'] ] )
    end
  end

  describe ":headers option" do

    it "should use the headers option if it exists" do
      xls = mock_users.to_xls(
        :columns => [:name, :email, :age],
        :headers => ['Nombre', 'Correo', 'Edad']
      )
      check_sheet( xls.worksheets.first,
        [ ['Nombre', 'Correo',      'Edad'],
          ['Peter',  'peter@gmail.com', 20],
          ['John',   'john@gmail.com',  25],
          ['Day9',   'day9@day9tv.com', 27]
        ]
      )
    end

    it "should include no headers if the headers option is false" do
      xls = mock_users.to_xls(
        :columns => [:name, :email, :age],
        :headers => false
      )
      check_sheet( xls.worksheets.first,
        [ ['Peter',  'peter@gmail.com', 20],
          ['John',   'john@gmail.com',  25],
          ['Day9',   'day9@day9tv.com', 27]
        ]
      )
    end

  end


end
