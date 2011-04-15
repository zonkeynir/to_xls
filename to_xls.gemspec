Gem::Specification.new do |s|
  s.name = %q{to_xls}
  s.version = "0.2.1"

  s.required_rubygems_version = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.authors = ["Enrique Garcia Cota", "Francisco de Juan", "Alan Larkin"]
  s.date = %q{2010-10-31}
  s.description = %q{Transform an Array or Hash into a excel file using the spreadsheet gem.}
  s.email = %q{egarcia@splendeo.es}
  s.extra_rdoc_files = [
    "MIT-LICENSE",
    "README.rdoc"
  ]
  s.files = [
    "MIT-LICENSE",
    "README.rdoc",
    "to_xls.gemspec",
    "init.rb",
    "lib/to_xls.rb"
  ]
  s.homepage = %q{http://github.com/splendeo/to_xls}
  s.rdoc_options = ["--charset=UTF-8"]
  s.require_paths = ["lib"]
  s.rubygems_version = %q{1.3.5}
  s.summary = %q{To xls}
  
  s.add_dependency("spreadsheet", ">= 0")

end
