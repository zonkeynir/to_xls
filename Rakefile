
#Bundler::GemHelper.install_tasks

require 'rake'
require "rubygems"
require "bundler"
require "bundler/gem_tasks"

Bundler.setup(:default, :test)

task :spec do
  begin
    require 'rspec/core/rake_task'

    desc "Run the specs under spec/"
    RSpec::Core::RakeTask.new do |t|
      #undefined method `spec_files=' for #<RSpec::Core::RakeTask:0x951189c>
      t.spec_files = FileList['spec/**/*_spec.rb']
    end
  rescue NameError, LoadError => e
    puts e
  end
end
