require File.expand_path(File.dirname(__FILE__) + '/spec_helper')

describe Array do

  it "should throw no error without data" do
    lambda { [].to_xls }.should_not raise_error
  end

  describe ":name option" do
    it "should default to 'Sheet 1' for sheets with no name" do
      [].to_xls.worksheets.first.name.should == 'Sheet 1'
    end

    it "should use the :name option" do
      [].to_xls(:name => 'Empty').worksheets.first.name.should == 'Empty'
    end
  end

  describe ":columns option" do
    it "should throw no error without columns" do
      lambda { [1,2,3].to_xls }.should_not raise_error
    end
    it "should throw an error if columns isn't an array" do
      lambda { [1,2,3].to_xls(:columns => :foo) }.should raise_error
    end
    it "should use the attribute keys as columns if it exists" do
      xls = mock_users.to_xls
      check_sheet( xls.worksheets.first,
        [ [:age,  :email,           :name],
          [   20, 'peter@gmail.com', 'Peter'],
          [   25, 'john@gmail.com',  'John'],
          [   27, 'day9@day9tv.com', 'Day9']
        ]
      )
    end
    it "should allow re-sorting of the columns by using the :columns option" do
      xls = mock_users.to_xls(:columns => [:name, :email, :age])
      check_sheet( xls.worksheets.first,
        [ [:name,   :email,          :age],
          ['Peter', 'peter@gmail.com', 20],
          ['John',  'john@gmail.com',  25],
          ['Day9',  'day9@day9tv.com', 27]
        ]
      )
    end

    it "should work properly when you provide it with both data and column names" do
      xls = [1,2,3].to_xls(:columns => [:to_s])
      check_sheet( xls.worksheets.first, [ [:to_s], ['1'], ['2'], ['3'] ] )
    end

    it "should pick data from associations" do
      xls = mock_users.to_xls(:columns => [:name, {:company => [:name]}])
      check_sheet( xls.worksheets.first,
        [ [:name,  :name],
          ['Peter', 'Acme'],
          ['John',  'Acme'],
          ['Day9',  'EADS']
        ]
      )
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

    it "should pick data from associations" do
      xls = mock_users.to_xls(
        :columns => [:name, {:company => [:name]}],
        :headers => [:name, :company_name]
      )
      check_sheet( xls.worksheets.first,
        [ [:name,  :company_name],
          ['Peter', 'Acme'],
          ['John',  'Acme'],
          ['Day9',  'EADS']
        ]
      )
    end

  end


end
