#!/usr/bin/env ruby
# encoding: UTF-8

require 'net/http'
require 'net/https'
require 'uri'
require 'pp'
require 'erb'
require 'csv'
require 'optparse'
require 'ostruct'
require 'logger'
require 'nokogiri'
require 'iconv'

USERAGENT = 'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.0.1) Gecko/20060111 Firefox/1.5.0.1'

BASE_PATH = 'http://search.eowa.gov.au'
SEARCH_PATH = BASE_PATH + '/Results.asp'

def view_response(resp)
  if @options[:debug]
    puts '------------------'
    puts 'Code = ' + resp.code
    puts 'Message = ' + resp.message
    resp.each {|key, val| puts key + ' = ' + val}
    puts '------------------'
    puts "\n"
  else
    pp resp
  end
end

def connect()
  url = URI.parse(BASE_PATH)
  http = Net::HTTP.new url.host, url.port
  http
end

def fetch_list(http)
  @log.debug("fetching company list")
  puts '* fetching company list'
  
  data = "Matters=ALL&"+
    "Industry=ALL&"+
    "Area=ALL&"+
    "ReportingPeriodID=ALL&"+
    "Emp_From=ANY&"+
    "Emp_To=ANY&"+
    "OrgName=&"+
    "submit1=Search"
  
  resp = http.post(SEARCH_PATH, data, @headers)
  view_response resp
  
  if resp.is_a? Net::HTTPSuccess
    
    ic = Iconv.new('UTF-8//IGNORE', 'UTF-8')
    output = []
    
    doc = Nokogiri::HTML(resp.body)
    doc.encoding = 'UTF-8'

    doc.css('tr[@bgcolor="silver"]').each do |tr|
      
      name = nil
      id = nil
      detail_url = nil
      report_period = nil
      
      tr.css('td:first a').each do |a|
        detail_url = BASE_PATH + '/' + a.attr('href')
        id = /=(\d+)$/.match(detail_url)[1] 
        name = ic.iconv(a.content)
      end
      
      report_period = tr.css('td:last')[0].content.strip
      
      output << [name, id, report_period, detail_url] unless detail_url.nil?
    end
    
    pp output
    write_to_csv(output, @list, true)
    
    @log.info("#{File.basename(@list)} generated")
  end
end

def fetch_detail(http, c)
  @log.debug("fetching company details")
  
  # read data from csv line
  name = c[0]
  id = c[1]
  report_period = c[2]
  detail_url = c[3]
  
  code = id+'_'+report_period
  
  puts "* fetching company #{id}_#{report_period} details"
  
  # fetch detail page
  resp = http.get2(detail_url, @headers)
  view_response resp
  
  unless resp.is_a? Net::HTTPSuccess
    @log.error("#{detail_url} cannot be opened (#{id}_#{report_period})")
    log_to_exception(code, name, detail_url, 'detail')
    
    return false
  end
  
  # parse html doc to get download link
  doc = Nokogiri::HTML(resp.body)  
  link = doc.css('a[@href^=ReportFiles]')[0]
    
  if link.nil?
    @log.error("no download link found in #{detail_url} (#{id}_#{report_period})")
    log_to_exception(code, name, detail_url, 'detail')
      
    return false
  end
      
  # reformat the download link:
  # * remove whitespaces
  # * convert \ to /
      
  download_url = BASE_PATH + '/' + link.attr('href').gsub(/\s+/,'').gsub(/\\+/, '/')
  # compile output filename
  output_file = File.join(@data_dir, "#{id}_#{report_period}#{File.extname(download_url)}")
      
  resp = http.get2(download_url, @headers)
  view_response resp
  
  unless resp.is_a? Net::HTTPSuccess
    @log.error("#{download_url} download failed (#{id}_#{report_period})")
    log_to_exception(code, name, download_url, 'download')
    
    return false
  end
  
  File.open(output_file, 'wb') do |f|
    f.write(resp.body)
  end
    
  @log.info("#{download_url} saved as #{File.basename(output_file)}")
  log_to_success(code, name)
       
end

def write_to_csv(output, output_file, override=false)
  return if File.exist?(output_file) unless override
  
  CSV.open(output_file, 'wb') do |csv|
    output.each do |line|
      csv << line
    end
  end
end

#def id_exist?(id, file)
#  sf = CSV.read(file, {encoding: "UTF-8"})
#  
#  sf.each do |l|
#    if l[1].to_i == id.to_i
#      puts id + ' found in ' + file
#      return true
#    end
#  end
#  
#  false
#end

def id_exist?(code, file)
  system("grep #{code} #{file} 2>&1 > /dev/null")
end

def log_to_exception(code, name, url, type)
  CSV.open(@exception_log, 'ab') do |csv|
    csv << [name, code, url, type]
  end
end

def log_to_success(code, name)
  CSV.open(@success_log, 'ab') do |csv|
    csv << [name, code]
  end
end

def usage
  puts @opts
  exit 1
end

@options = {}

@opts = OptionParser.new
@opts.banner = "Usage: #{File.basename($0)} [options]"
@opts.on("-s", "--source [FILE]", String, "Source file") do |s|
  @options[:source] = s
end
@opts.on("-o", "--output [DIRECTORY]", String, "(required) Output directory") do |o|
  @options[:output] = o
end
@opts.on("-f", "--fetch [TYPE]", [:list, :detail], "(required) Fetch data type (list, detail)") do |t|
  @options[:type] = t
end
@opts.on("-d", "--debug", "Debug mode") do |d|
  @options[:debug] = d
end
@opts.on_tail("-h", "--help", "Show this message") do
  puts @opts
end
@opts.parse! rescue usage

# sanity checks 
usage if @options[:type].nil?
usage unless @options[:output] != nil and File.directory? @options[:output]
usage if @options[:type] == :detail and @options[:source].nil?

@log_dir = File.join(@options[:output], 'log')
@data_dir = File.join(@options[:output], 'data')

@success_log = File.join(@log_dir, 'success.csv')
@exception_log = File.join(@log_dir, 'exception.csv')
@run_log = File.join(@log_dir, 'run.log')
@list = File.join(@data_dir, 'list.csv')

@headers = {
  'User-Agent' => USERAGENT,
}



Dir.mkdir(@log_dir) unless File.directory?(@log_dir)
Dir.mkdir(@data_dir) unless File.directory?(@data_dir)

unless File.exist?(@success_log)
  File.open(@success_log, 'w') {|f| f.close()}
end

unless File.exist?(@exception_log)
  File.open(@exception_log, 'w') {|f| f.close()} 
end

unless File.exist?(@run_log)
  File.open(@run_log, 'w') {|f| f.close()} 
end

# init log
@log = Logger.new(@run_log)
@log.level = Logger::INFO

# init http connection
http = connect()

case @options[:type]
when :list then
  fetch_list(http)
when :detail then
  # read the company list
  data = CSV.read(@options[:source], {encoding: "ISO-8859-1"})
  data.each do |c|
    id = c[1]
    report_period = c[2]
    
    code = id+'_'+report_period
    
    next if id_exist?(code, @success_log)
    next if id_exist?(code, @exception_log)
    
    puts ">> fetching #{code}"
    fetch_detail(http, c)
  end
end
