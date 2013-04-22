#!/usr/bin/env ruby
# encoding: utf-8

require 'csv'
require 'pp'
require 'optparse'

def usage
  puts @opts
  exit 1
end

@options = {}

@opts = OptionParser.new
@opts.banner = "Usage: #{File.basename($0)} [options]"
@opts.on("-s", "--source FILE", String, "source CSV file") { |s| @options[:source] = s }
@opts.on_tail("-h", "--help", "Show this message") { puts @opts }
@opts.parse! rescue usage

source = @options[:source]
target = source + '.out'

data = CSV.read(source, {encoding: 'UTF-8'})

out = []
data.each { |p| out << [p[0].gsub("\u0092", "'"),p[1]] }
#pp out

CSV.open(target, 'wb') do |csv|
  out.each do |l|
    csv << l
  end
end
