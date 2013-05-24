class Excelcsvator
  attr_accessor :file
  
  def initialize(file)
    @file = file
  end

  def path
    @file.path
  end

  def to_csv
    to_c_csv
  end
end

require 'excelcsvator_ext/excelcsvator_ext'