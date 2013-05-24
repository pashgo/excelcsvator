$:.push File.expand_path("../lib", __FILE__)

# Describe your gem and declare its dependencies:
Gem::Specification.new do |s|
  s.name        = "excelcsvator"
  s.version     = "0.0.1"
  s.authors     = ["gap intelligence"]
  s.email       = ["info@gapintelligence.com"]
  s.homepage    = "http://www.gapintelligence.com"
  s.summary     = "Excelcsvator gem."
  s.description = "Excel to CSV converter."

  s.files = Dir["{ext,lib}/**/*{.rb,.c,.h}"] + ["MIT-LICENSE", "Rakefile", "README.rdoc"]
  s.extensions = ['ext/excelcsvator_ext/extconf.rb']
end
