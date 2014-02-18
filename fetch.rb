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
require 'yaml'

### constants

CONFIG = 'config.yaml'
USERAGENT = 'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.0.1) Gecko/20060111 Firefox/1.5.0.1'

SEARCH_PATH = '/search/quicksearch/individual/default.aspx?menuCat=1&pCategory=6&d=1&q='
DETAIL_PATH = '/director/profile/default.aspx?menuCat=1&pCategory=6&dir_id='
CONNECTION_PATH = '/director/associations/default.aspx?menuCat=1&pCategory=6&menuGrp=net'

CONNECTION_HEADER = 'Individual,Organisation,Organisation Type,Current / Historic,Duration,Role of Root Individual,Overlap Start Date,Overlap End Date,Role of Connected Individual'
POSITION_HEADER = 'Id,Name,Organisation,Role,Role Description,Start Date,End Date,Type'
EDUCATION_HEADER = 'Id,Name,Date,Institute,Qualification'

### classes

class Position
  attr_accessor :id, :name, :org_name, :role, :role_description, :start_data, :end_data,  :type

  def initialize(type)
    @type = type
  end

  def is_brd?
    @role =~ /\(Brd - /
  end

  def to_array
    [@id, @name, @org_name, @role, @role_description, @start_data, @end_data, @type]
  end
end

class Education
  attr_accessor :id, :name, :date, :institute, :qualification
  def to_array
    [@id, @name, @date, @institute, @qualification]
  end
end

### functions

# Load in the YAML configuration file, check for errors, and return as hash 'cfg'
#
# Ex config.yaml:
#---
#login:
#   username: xxx
#   pasword: xxx
#
def load_config
  cfg = File.open(CONFIG)  { |yf| YAML::load( yf ) } if File.exists?(CONFIG)
  # => Ensure loaded data is a hash. ie: YAML load was OK
  if cfg.class != Hash
     raise "ERROR: Configuration - invalid format or parsing error."
  else
    if cfg['login'].nil?
      raise "ERROR: Configuration: login not defined."
    end
  end

  return cfg
end

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

def login

  url = URI.parse('https://www.boardex.com')
  http = Net::HTTP.new url.host, url.port
  http.verify_mode = OpenSSL::SSL::VERIFY_NONE
  http.use_ssl = true
  http.read_timeout = 60 # seconds
  path = '/Login.aspx'

  resp = http.get2(path, {'User-Agent' => USERAGENT})
  cookie = resp.response['set-cookie'].split('; ')[0]

  eventvalidatoin = ''
  previouspage = ''
  viewstate = ''

  resp.body.each_line do |line|
    eventvalidatoin = /value=\"(.*)\"/.match(line)[1] if line =~ /__EVENTVALIDATION/
    previouspage = /value=\"(.*)\"/.match(line)[1] if line =~ /__PREVIOUSPAGE/
    viewstate = /value=\"(.*)\"/.match(line)[1] if line =~ /__VIEWSTATE/
  end

  data = "__EVENTARGUMENT=&" +
    "__EVENTTARGET=&" +
    "__EVENTVALIDATION=#{ERB::Util.url_encode(eventvalidatoin)}&" +
    "__PREVIOUSPAGE=#{ERB::Util.url_encode(previouspage)}&" +
    "__VIEWSTATE=#{ERB::Util.url_encode(viewstate)}&" +
    "_txtLoginName=#{@username}&" +
    "_txtPassword=#{@password}&" +
    "_btnLogin=Log%20in"

  headers = {
    'User-Agent' => USERAGENT,
    'Cookie' => cookie,
    'Referer' => 'https://www.boardex.com/Login.aspx',
    'Content-Type' => 'application/x-www-form-urlencoded'
  }

  resp = http.post(path, data, headers)

  # Output on the screen -> we should get either a 302 redirect (after a successful login) or an error page
  view_response resp
  cookie = resp.response['set-cookie'].split('; ')[0]
  puts "* Logged in as #{@username}"
  return http, cookie
end

def search(name, http, cookie)
  puts '* Search Person: ' + name

  headers = {'User-Agent' => USERAGENT,'Cookie' => cookie}
  path = SEARCH_PATH + URI.encode_www_form_component(name)
  resp = http.get2(path, headers)
  view_response resp

  case resp
  when Net::HTTPSuccess then
    log_to_exception(name)
    @log.warn(name + ' matches none or more than one records')
  when Net::HTTPRedirection then
    id = /dir_id=(\d+)$/.match(resp['location'])[1]

    if id != nil and id =~ /\d+/
      @log.info("#{name} found (#{id})")
      return id
    end
  else
    @log.error("http error occured: #{resp.code}")
  end

  return nil
end

def fetch_details(id, http, cookie)
  puts '* Fetching Details for PID: ' + id

  headers = {'User-Agent' => USERAGENT,'Cookie' => cookie}
  path = DETAIL_PATH + id
  resp = http.get2(path, headers)
  view_response resp
  resp
end

def fetch_connections(id, name, http, cookie)
  puts '* Fetching Conncetions'

  headers = {'User-Agent' => USERAGENT,'Cookie' => cookie}
  path = CONNECTION_PATH
  resp = http.get2(path, headers)
  view_response resp

  viewstate = ''
  resp.body.each_line do |line|
    if line =~ /__VIEWSTATE/
      viewstate = /value=\"(.*)\"/.match(line)[1]
      break
    end
  end

  data = "__EVENTARGUMENT=&" +
    "__EVENTTARGET=&" +
    "__LASTFOCUS=&" +
    "__VIEWSTATE=#{ERB::Util.url_encode(viewstate)}&" +
    "_ctl0%3A_ctl0%3AContentPlaceholder%3AContentHeaderMultipleContainers%3AContentHeaderBottomPanel%3ADownloadControls%3AbtnViewDrivenDownload=Download&" +
    "_ctl0%3A_ctl0%3AContentPlaceholder%3AContentHeaderMultipleContainers%3AContentHeaderBottomPanel%3ADownloadControls%3ADocumentTypes=8&" + 
    "_ctl0%3A_ctl0%3AContentPlaceholder%3AStandardContentPlaceholder%3AFilters%3AddlRelationshipType=32767&" +
    "_ctl0%3A_ctl0%3AContentPlaceholder%3AStandardContentPlaceholder%3AFilters%3AddlRouteToTarget=119&" +
    "_ctl0%3A_ctl0%3AContentPlaceholder%3AStandardContentPlaceholder%3AFilters%3AchxIncludeMembers=on&" + 
    "_ctl0%3A_ctl0%3AContentPlaceholder%3AStandardContentPlaceholder%3AFilters%3AfilterState=0&" +
    "_ctl0%3A_ctl0%3AContentPlaceholder%3AStandardContentPlaceholder%3AFilters%3AMultipleConnectionsAppliedValue=0&" +
    "_ctl0%3A_ctl0%3AContentPlaceholder%3AStandardContentPlaceholder%3AFilters%3ACurrentConnectionsAppliedValue=0&" +
    "_ctl0%3A_ctl0%3AContentPlaceholder%3AStandardContentPlaceholder%3AFilters%3ARelationshipTypeAppliedValue=32767&" +
    "_ctl0%3A_ctl0%3AContentPlaceholder%3AStandardContentPlaceholder%3AFilters%3AIncludeMembersAppliedValue=1&" +
    "_ctl0%3A_ctl0%3AContentPlaceholder%3AStandardContentPlaceholder%3AFilters%3AOverlapsOnlyAppliedValue=0&" + 
    "_ctl0%3A_ctl0%3AContentPlaceholder%3AStandardContentPlaceholder%3AFilters%3ARouteToTargetAppliedValue=119"

  headers = {
    'User-Agent' => USERAGENT,
    'Cookie' => cookie,
    'Referer' => 'https://www.boardex.com/director/associations/default.aspx?menuCat=1&pCategory=6&menuGrp=net',
    'Content-Type' => 'application/x-www-form-urlencoded'
  }

  resp = http.post(path, data, headers)

  puts '* Downloading CSV'
  view_response resp

  output_file = File.join(@data_dir, "#{id}_connections.csv")

  if resp.is_a? Net::HTTPSuccess
    body_ary = resp.body.split("\r\n")

    full_name = body_ary.shift

    header_ary = CONNECTION_HEADER.split(',')

    output = [['Person of Interest'] + header_ary]

    body_ary.each do |line|
      line_ary = line.strip.parse_csv

      next if line_ary.nil?  # next if the line is empty
      next if line_ary.join(',') == CONNECTION_HEADER  # next if the line is the header
      next if line_ary.length != header_ary.length  # next if the line not matching the header

      output << [full_name] + line_ary
    end

    ## write to file
    write_to_csv(output, output_file)

    ## log to success
    log_to_success(id, name)

    @log.info("#{File.basename(output_file)} download success")
  else
    @log.error("#{File.basename(output_file)} download failed: (#{resp.code})")
  end

end

def write_to_csv(output, output_file)
  unless File.exist?(output_file)
    CSV.open(output_file, 'wb') do |csv|
      output.each do |arr|
        csv << arr
      end
    end
  end
end

def extract_positions(id, name, resp)
  @log.debug("extracting position data")

  doc = Nokogiri::HTML(resp.body)

  current_div_id = '_ctl0__ctl0_ContentPlaceholder_StandardContentPlaceholder_currentPositionsGrid_gridView'
  past_div_id = '_ctl0__ctl0_ContentPlaceholder_StandardContentPlaceholder_pastPositionsGrid_gridView'

  current_positions = doc.css("#"+current_div_id)
  past_positions  = doc.css("#"+past_div_id)

  positions = []

  unless current_positions.nil?
    current_positions.css('tr').each do |tr|
      next unless tr.css('th').empty?

      data = tr.css('td').collect { |obj| obj.content.strip }

      position = Position.new('current')
      position.start_data = data[0]
      position.org_name = data[1]
      position.role = data[2]
      position.role_description = data[3]

      positions << position if position.is_brd?
    end
  end

  unless past_positions.nil?
    past_positions.css('tr').each do |tr|
      next unless tr.css('th').empty?

      data = tr.css('td').collect { |obj| obj.content.strip }

      position = Position.new('past')
      position.start_data = data[0]
      position.end_data = data[1]
      position.org_name = data[2]
      position.role = data[3]
      position.role_description = data[4]

      positions << position if position.is_brd?
    end
  end

  if positions.empty?
    @log.warn("No position data found for person: #{name} (#{id})")
    ## log to success 
    log_to_exception(name)
  else
    output = [POSITION_HEADER.split(',')]

    positions.each do |p|
      p.id = id
      p.name = name
      output << p.to_array
    end

    output_file = File.join(@data_dir, "#{id}_positions.csv")

    ## write to file
    write_to_csv(output, output_file)

    ## log to success 
    log_to_success(id, name)
    @log.info("#{File.basename(output_file)} saved")
  end
end

def extract_education(id, name, resp)
  doc = Nokogiri::HTML(resp.body)
  e_doc = doc.css('#_ctl0__ctl0_ContentPlaceholder_StandardContentPlaceholder_educationGrid_educationGridView')

  educations = []
  unless e_doc.nil?
    e_doc.css('tr').each do |tr|
      next unless tr.css('th').empty?
      data = tr.css('td').collect { |obj| obj.content.strip }
      e = Education.new
      e.date = data[0]
      e.institute = data[1]
      e.qualification = data[2]
      educations << e
    end
  end

  if educations.empty?
    @log.warn("No education data found for person: #{name} (#{id})")
    ## log to success
    log_to_exception(name)
  else
    output = [EDUCATION_HEADER.split(',')]

    educations.each do |e|
      e.id = id
      e.name = name
      output << e.to_array
    end

    output_file = File.join(@data_dir, "#{id}_educations.csv")

    ## write to file
    write_to_csv(output, output_file)

    ## log to success
    log_to_success(id, name)
    @log.info("#{File.basename(output_file)} saved")
  end
end

def name_exist?(name, file)
  if @grep_exist
    if system('grep "' + name + '" ' + file + ' 1>/dev/null')
      puts name + ' found in ' + file
      return true
    end
  else
    CSV.read(file, {encoding: "UTF-8"}).each do |l|
      if l[0].eql? name
        puts name + ' found in ' + file
        return true
      end
    end
  end
  false
end

def id_exist?(id, file)
  if @grep_exist
    if system('grep ",' + id + '$" ' + file + ' 1>/dev/null')
      puts id + ' found in ' + file
      return true
    end
  else
    CSV.read(file, {encoding: "UTF-8"}).each do |l|
      if l[1].to_i == id.to_i
        puts id + ' found in ' + file
        return true
      end
    end
  end
  false
end

def log_to_exception(name)
  CSV.open(@exception_log, 'ab') {|csv| csv << [name]}
end

def log_to_success(id, name)
  CSV.open(@success_log, 'ab') {|csv| csv << [name, id]}
end

def usage
  puts @opts
  exit 1
end

### main

@options = {}

@opts = OptionParser.new
@opts.banner = "Usage: #{File.basename($0)} [options]"
@opts.on("-s", "--source [FILE]", String, "(required) Source file") do |s|
  @options[:source] = s
end
@opts.on("-o", "--output [DIRECTORY]", String, "(required) Output directory") do |o|
  @options[:output] = o
end
@opts.on("-f", "--fetch [TYPE]", [:id, :connection, :position, :education], "(required) Fetch data type (id, connection, position, education)") do |t|
  @options[:type] = t
end
@opts.on("-i", "--by-id", "Fetch data by ID (the second column of source CSV)") do |i|
  @options[:by_id] = i
end
@opts.on("-d", "--debug", "Debug mode") do |d|
  @options[:debug] = d
end
@opts.on_tail("-h", "--help", "Show this message") do
  puts @opts
  exit
end
@opts.parse! rescue usage

# sanity checks 
usage unless @options[:source] != nil and File.exists? @options[:source]
usage unless @options[:output] != nil and File.directory? @options[:output]
usage unless @options[:type] != nil

# load conifg
cfg = load_config
@username = cfg['login']['username']
@password = cfg['login']['password']

# check if grep exist
@grep_exist = system('which grep 2>&1 1>/dev/null')

@log_dir = File.join(@options[:output], 'log')
@data_dir = File.join(@options[:output], 'data')

@success_log = File.join(@log_dir, 'success.csv')
@exception_log = File.join(@log_dir, 'exception.csv')
@run_log = File.join(@log_dir, 'run.log')

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

# login to BoardEx
begin
  http,cookie = login
rescue => e
  message = "login failed: ("+e.message+")"
  puts message
  @log.error(message)
  exit 1
end

# read the PID list
data = CSV.read(@options[:source], :headers => true, :encoding => "ISO-8859-1")

case @options[:type]
when :id then
  data.each do |p|
    name = p[0].strip.encode('UTF-8')
    begin
      next if name_exist?(name, @success_log)
      next if name_exist?(name, @exception_log)
      id = search(name, http, cookie)
      ## log to success
      log_to_success(id, name) unless id
    rescue => e
      @log.error("something when wrong prcessing id:#{id} ("+e.message+")")
      @log.error(e.backtrace)
      next
    end
  end
when :connection then
  # csv format:
  # column0 => name
  # column1 => id
  data.each do |p|

    name = p[0].strip.encode('UTF-8')
    id = p[1]

    #  puts name
    #  puts name.encoding
    #  puts name.length

    begin
      if @options[:by_id] and id
        next if id_exist?(id, @success_log)
        fetch_details(id, http, cookie)
        fetch_connections(id, name, http, cookie)
      else
        next if name_exist?(name, @success_log)
        next if name_exist?(name, @exception_log)
        id = search(name, http, cookie)
        next unless id
        fetch_details(id, http, cookie)
        fetch_connections(id, name, http, cookie)
      end
    rescue => e
      @log.error("something when wrong processing id:#{id} ("+e.message+")")
      @log.error(e.backtrace)
      next
    end
  end
when :position, :education then
  unless @options[:by_id]
    puts 'option -i is required'
    exit 1
  end

  data.each do |p|
    name = p[0].strip.encode('UTF-8')
    id = p[1]
    begin
      next if id_exist?(id, @success_log)
      next if name_exist?(name, @exception_log)
      resp = fetch_details(id, http, cookie)
      extract_positions(id, name, resp) if @options[:type] == :position
      extract_education(id, name, resp) if @options[:type] == :education
    rescue => e
      @log.error("something when wrong processing id:#{id} ("+e.message+")")
      @log.error(e.backtrace)
      next
    end
  end
end
