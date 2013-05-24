#!/usr/bin/env rake
require 'rake/testtask'
require 'rake/extensiontask'

Bundler::GemHelper.install_tasks

Rake::TestTask.new(:test) do |t|
  t.libs << 'lib'
  t.libs << 'test'
  t.pattern = 'test/**/*_test.rb'
  t.verbose = false
end

Rake::ExtensionTask.new('excelcsvator_ext')


task :default => :test
