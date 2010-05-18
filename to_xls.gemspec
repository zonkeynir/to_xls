Gem::Specification.new do |s|
  s.name = %q{to_xls}
  s.version = "0.0.1"

  s.required_rubygems_version = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.authors = ["Enrique Garcia Cota", "Francisco de Juan"]
  s.date = %q{2010-05-18}
  s.description = %q{Transform an Array into a excel file using the spreadsheet gem.}
  s.email = %q{github@splendeo.es}
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

  if s.respond_to? :specification_version then
    current_version = Gem::Specification::CURRENT_SPECIFICATION_VERSION
    s.specification_version = 3

    if Gem::Version.new(Gem::RubyGemsVersion) >= Gem::Version.new('1.2.0') then
      s.add_development_dependency(%q<spreadsheet>, [">= 0"])
    else
      s.add_dependency(%q<spreadsheet>, [">= 0"])
    end
  else
    s.add_dependency(%q<spreadsheet>, [">= 0"])
  end
end
