# -*- encoding: utf-8 -*-
$:.push File.expand_path("../lib", __FILE__)
require "to_xls/version"

Gem::Specification.new do |s|
  s.name        = "to_xls"
  s.version     = ToXls::VERSION
  s.authors     = ["Enrique Garcia Cota", "Francisco de Juan"]
  s.email       = %q{egarcia@splendeo.es}
  s.homepage    = "https://sendorf@github.com/splendeo/to_xls.git"
  s.summary     = %q{to_xls for Arrays and Hashes}
  s.description = %q{Adds a to_xls method to arrays, which can be used to generate excel files conveniently. Can rely on ActiveRecord sugar for obtaining attribute names.}

  s.rubyforge_project = "to_xls"

  s.files         = `git ls-files`.split("\n")
  s.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")

  s.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ["lib"]

  # specify any dependencies here; for example:
  # s.add_development_dependency "rspec"
  # s.add_runtime_dependency "rest-client"

  if s.respond_to? :specification_version then
    s.specification_version = 3

    if Gem::Version.new(Gem::VERSION) >= Gem::Version.new('1.2.0') then
      s.add_runtime_dependency(%q<spreadsheet>, [">= 0"])
      s.add_development_dependency(%q<rspec>, ["~> 2.3.0"])
      s.add_development_dependency(%q<bundler>, ["~> 1.0.0"])
      s.add_development_dependency(%q<jeweler>, ["~> 1.5.2"])
      s.add_runtime_dependency(%q<spreadsheet>, ["> 0"])
      s.add_development_dependency(%q<rspec>, ["> 1.2.3"])
    else
      s.add_dependency(%q<spreadsheet>, [">= 0"])
      s.add_dependency(%q<rspec>, ["~> 2.3.0"])
      s.add_dependency(%q<bundler>, ["~> 1.0.0"])
      s.add_dependency(%q<jeweler>, ["~> 1.5.2"])
      s.add_dependency(%q<spreadsheet>, ["> 0"])
      s.add_dependency(%q<rspec>, ["> 1.2.3"])
   end
  else

    s.add_dependency(%q<spreadsheet>, [">= 0"])
    s.add_dependency(%q<rspec>, ["~> 2.3.0"])
    s.add_dependency(%q<bundler>, ["~> 1.0.0"])
    s.add_dependency(%q<jeweler>, ["~> 1.5.2"])
    s.add_dependency(%q<spreadsheet>, ["> 0"])
    s.add_dependency(%q<rspec>, ["> 1.2.3"])
  end
end
