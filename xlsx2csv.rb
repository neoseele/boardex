#!/usr/bin/env ruby
# encoding: utf-8

require 'csv'
require 'pp'
require 'optparse'
require 'creek'

def usage
  puts @opts
  exit 1
end

@options = {}

@opts = OptionParser.new
@opts.banner = "Usage: #{File.basename($0)} [options]"
@opts.on("-s", "--source DIR", String, "xlsx DIR") { |s| @options[:src] = s }
@opts.on("-d", "--destination DIR", String, "csv DIR") { |d| @options[:dest] = d }
@opts.on_tail("-h", "--help", "Show this message") { puts @opts }
@opts.parse! rescue usage

src_dir = @options[:src]
dest_dir = @options[:dest]

usage if src_dir.nil? or not File.directory?(src_dir)
usage if dest_dir.nil? or not File.directory?(dest_dir)

Dir.glob(File.join(src_dir,'*.xlsx')) do |xlsx|
  csv_name = File.basename(xlsx,'.*').gsub(/\W/, '') + '.csv'
  csv_file = File.join(dest_dir, csv_name)

  if File.exist? csv_file
    puts "[skip]: #{csv_file} exists"
    next
  end

  creek = Creek::Book.new xlsx
  sheet= creek.sheets[0]

  CSV.open(csv_file, 'wb') do |csv|
    sheet.rows.each_with_index do |row, i|
      next if i == 0 # skip the first line
      if i == 1
        csv << row.values.map.with_index do |c,i|
          header = c.to_s.downcase.gsub(/(\*|\'|\(|\))/,'').gsub(/(\W|\/)/,'_') + "_#{i}"
        end
      else
        csv << row.values.map {|c| c.to_s}
      end
    end
  end
end
