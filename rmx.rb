#!/usr/bin/ruby
########################################################
# RealtXpress - sandbox2.3.rb   3/1/10               #
#     v2.3                                             #
#                                                      #
#             Written by Charles Thompson              #
#               charles@kinetic-it.com                 #
#                 www.kinetic-it.com                   #
#                                                      #
#          Built with Ruby v1.8.6p111 mswin32          #
#                                                      #
#  Licensed to : Tom Griffey of Anchor Properties      #
#                                                      #
#  UNPUBLISHED PROPRIETARY SOURCE CODE!                #
#  DISTRIBUTION OF THIS PRODUCT WITHOUT LICENSING      #
#  IS ILLEGAL AND PUNISHABLE BY LAW.                   #
#                                                      #
#  (c) 2008 Charles Thompson, all rights reserved      #
#  Source code modifications permitted to license      #
#  holder ONLY provided this notice is kept intact.    #
#  Resale or distribution prohibited by law.           #
#                                                      #
########################################################
#                                                      #
#   Modified March 2010 - Josh Shupack                   #
#                                                      #
########################################################
#                                                      #
#   Modified June 2010                                 #
#   Jeff Cook, Deseret Technology, Inc.                #
#   jeff@deserettechnology.com                         #
#   Updated for new login layout                       #
#                                                      #
########################################################
#
# Compiled using rubyscript2exe - http://www.erikveen.dds.nl/rubyscript2exe/
# Ruby version - One-Click Installer (old) - 1.8.6-26 Final Release
#    http://rubyforge.org/frs/?group_id=167
#
# rubyscript2exe hasn't worked since mid-2009.
# Using ocra now instead. - jcook
#
# Gem versions:
# hpricot (0.6)
#
# facets has been replaced by hashery since facets 2.9, which shed most supplemental libraries.
# hashery (1.3.0)
#
# parseexcel (0.5.2)
# rake (0.7.3)
# watir (1.5.6)
# activesupport (2.3.3)
# hoe (1.3.0)
#

$KCODE = 'u'
# add fox16

I_KNOW_I_AM_USING_AN_OLD_AND_BUGGY_VERSION_OF_LIBXML2 = ''
%w[ csv pp ostruct set  fileutils
  rubygems watir parseexcel hpricot hashery/dictionary
].each{|lib| require lib }
require 'json'
require 'net/http'
require 'win32/registry'

unless defined?(Ocra)

  module ExcelBot
    COUNTIES = {
      "RI" => "RIVERSIDE",
      "SD" => "SAN DIEGO",
      "LA" => "LOS ANGELES",
      "OC" => "ORANGE",
      "SB" => "SAN BERNARDINO",
      "VEN" => "VENTURA",
      "IV" => "IMPERIAL",
      "SBR" => "SANTA BARBARA",
    }
    COUNTY_CODES = {
      "LOS ANGELES" => '153',
      "RIVERSIDE" => '157',
      "SAN BERNARDINO" => '158',
      "ORANGE" => '154',
      "SAN DIEGO" => '159',
      "IMPERIAL" => '2810',
      "SAN LUIS OBISPO" => '2812',
      "SANTA BARBARA" => '2813',
      "VENTURA" => '1292',
      "KERN" => '2811',
    }
    STATES = {
      "ALABAMA" => "AL",
      "ALASKA" => "AK",
      "ARIZONA" => "AZ",
      "ARKANSAS" => "AR",
      "CALIFORNIA" => "CA",
      "COLORADO" => "CO",
      "CONNECTICUT" => "CT",
      "DELAWARE" => "DE",
      "DISTRICT OF COLUMBIA" => "DC",
      "FLORIDA" => "FL",
      "GEORGIA" => "GA",
      "HAWAII" => "HI",
      "IDAHO" => "ID",
      "ILLINOIS" => "IL",
      "INDIANA" => "IN",
      "IOWA" => "IA",
      "KANSAS" => "KS",
      "KENTUCKY" => "KY",
      "LOUISIANA" => "LA",
      "MAINE" => "ME",
      "MARYLAND" => "MD",
      "MASSACHUSETTS" => "MA",
      "MICHIGAN" => "MI",
      "MISSISSIPPI" => "MS",
      "MISSOURI" => "MO",
      "MONTANA" => "MT",
      "NEBRASKA" => "NE",
      "NEVADA" => "NV",
      "NEW HAMPSHIRE" => "NH",
      "NEW JERSEY" => "NJ",
      "NEW MEXICO" => "NM",
      "NEW YORK" => "NY",
      "NORTH CAROLINA" => "NC",
      "NORTH DAKOTA" => "ND",
      "OHIO" => "OH",
      "OKLAHOMA" => "OK",
      "OREGON" => "OR",
      "PENNSYLVANIA" => "PA",
      "RHODE ISLAND" =>  "RI",
      "SOUTH CAROLINA" => "SC",
      "SOUTH DAKOTA" => "SD",
      "TENNESSEE" => "TN",
      "TEXAS" => "TX",
      "UTAH" => "UT",
      "VERMONT" => "VT",
      "VIRGINIA" => "VA",
      "WASHINGTON" => "WA",
      "WEST VIRGINIA" => "WV",
      "WISCONSIN" => "WI",
      "WYOMING" => "WY"
    }
    DEFAULTS = {
      :output => "excel-output.csv",
      :login_pf => 'login-pf.txt',
      #:login_matrix => 'login-matrix.txt', #we  no longer log in
      :csv_pf => 'pf.csv',
      :csv_matrix => 'matrix.csv',
    }
    module WatirHelper

      def initialize_agent
        #@watir = Watir::IE.attach(:title, 'CRMLS Matrix - Windows Internet Explorer');
		#@watir = Watir::IE.attach(:url, 'CRMLS Matrix')
    @watir = Watir::IE.attach(:title, /.*/)
		puts @watir.title();
		
        puts "Opening Internet Explorer...please wait"
        #@watir = Watir::IE.new
        @watir.speed = :zippy
      end

      def safely
        @once = false
        yield
      rescue Exception => ex
        print "Error: ", ex, " - retrying in 10sec...\n"
        exit if ex =~ /wrong status line/i
        if @trapped == 1
          puts "Exiting due to CTRL-C"
          exit
        end
        sleep 10
        (@once = true; retry) unless @once
        raise ex
      end

    end
    class Extractor
      include WatirHelper
      attr_accessor :options

      def initialize(options, verbose = false)
        @options, @verbose = options, verbose
        File.open(out_file, 'w+'){|io|
          # io.puts CSV.generate_line(self.class::CSV_KEYS)
        }
        #creds_from(login_file) #no longer needed as we don't manually log in anymore.

        initialize_agent
        Signal.trap("INT") {
          @trapped = 1
          sleep 5
          exit
        }
      end

      def creds_from(file)
        @name, @pass = File.readlines(file).map{|line| line.strip }
      end

      def xls_each(file = options.xls)
        workbook = Spreadsheet::ParseExcel.parse(file)
        worksheet = workbook.worksheet(0)

        header = nil

        worksheet.each_with_index do |row, y|
          if y == 0
            header = row.map{|col| parse_cell(col) }
          else
            query = Hashery::Dictionary.new
            header.each{|h| query[h] = '' }
            if row == nil
              puts "Program complete - logging out."
              exit
            end
            row.each_with_index do |col, x|
              next unless key = header[x]

              if col
                value = parse_cell(col)

                case key
                when /state/i
                  key = 'MState'
                  value.upcase!
                  value = STATES[value] || value
                  { key => value }.inspect
                end
              else
                value = ''
              end

              query[key] = value
            end


            puts " Record [#{y}/#{worksheet.num_rows}] ".center(80, '-')
            yield query
            puts " End Results ".center(80, '-'), ''
          end
        end
      end

      def parse_cell(cell)
        begin
          string = cell.to_s('latin1')
        rescue ArgumentError
          string = cell.to_s
        end

        case string
        when /^\d+(\.0+$|$)/
          string.to_i
        when /^\d+\.\d+$/
          string.to_f
        else
          string
        end
      end
    end

    class MRMLSMatrix < Extractor
      @methodtmp = 0
      @loggedin = true

      def methodtmp=(methodtmp)
        @methodtmp = methodtmp
      end

#      def login_file
#        options.login_matrix
#      end

      def out_file
        options.csv_matrix
      end
      def login(method)
        @methodtmp = 0
		if  @loggedin == false
			@watir.goto('http://idp.mrmls.safemls.net/idp/Authn/UserPassword')

			success = @watir.text.include?("Click here if you are having trouble logging into your account.") rescue false
			if !success
			  puts "Cannot access login page, could not find login text"
			  return false
			end

			print "\nEnter SafeMLS token: "
			while @secure_key = $stdin.gets
			  if @secure_key.to_s.length < 7
				puts "Invalid token!"
				print "\nEnter SafeMLS token: "
			  else
				break
			  end
			end

			if @secure_key != nil
			  safely do
				@watir.text_field(:name, "j_username2").set(@name)
				@watir.text_field(:name, "j_pin2").set(@pass)
				@watir.text_field(:name, "j_password2").set(@secure_key)
				@watir.button(:id, "login2").click
			  end
			end

			success = @watir.text.include?("Thomas Griffey") rescue false
			if !success
			  puts "Cannot login, couldn't find proper text after entering main page"
			  return false
			end
		end

        puts "Successfully logged in."

        if method == 0 or method == 3
          begin
            safely do
			  @watir.goto("https://realist2.firstamres.com/propertylink?CustomerGroupName=MRMLS&UserID=C12729&UserPW=xuuN3tme&AgentFirstName=THOMAS&AgentLastName=GRIFFEY")
            end
            # What is this for?
            # if continue_form = @watir.form(:name, 'Form1')
            #            @watir.button(:name, "btnContinue").click
            #          end

            success = @watir.button(:name, "CreateGroup") rescue false
            if !success
              puts "Couldn't access http://www.mrmlsmatrix.com/Matrix/Special/Realist.aspx"
              raise "Raising exception, let's try again"
            end
          rescue => e
            puts "Error when logging in using method 0 - retrying: #{e.backtrace}"
            retry
          end
        end
        #### NEW LOGIN METHOD #####
        if method == 1 or method == 2
          safely do
           # @watir.link(:text, "Search").click
           @watir.link(:url, /\/Matrix\/Search/).click;

          end

          #What is this for?
          # if continue_form = @watir.form(:name, 'Form1')
          #             continue_form.button(:name, "btnContinue").click
          #           end
        end

        # Set this so that we can check if its set when processing each property
        if method == 3
          @methodtmp = 3
        else
          @methodtmp = 0
        end

        return true
        ## END NEW METHOD
      end

      def search_result(query)
        # Create our methods and ways of utilizing them
        begin
          # pp query
          @watir.goto("http://realist2.firstamres.com/searchapn.jsp")
          #puts page.form["zipListString"].value

          safely do
            @watir.select_list(:name, "county").select(query["county"])
          end
          sleep 1

          safely do
            @watir.text_field(:name, "apn_entered").set(query["apn_entered"])
          end
          safely do
            @watir.button(:name, "Submit").click
          end

          sleep 2
          if !@watir.text.include?("Subject Property")
            sleep 2
          end
        rescue => e
          puts "Ran into a problem in method 'search_result' - Retrying: #{e.backtrace}\n"
          retry
        end

        return Hpricot(@watir.html)
      end

      def parse_reports(doc, query = {})
        # We store the result so if we get an error along the way we can analyze the document
        # File.open('result.html', 'w+'){|io| io.write(doc.to_s) }
        empty = { 'APN' => query['APN'] }
        doc.search('td.report_results') do |td|
          if td.inner_text =~ /No results were found/
            yield empty if block_given?
            return
          end
        end
        if address_td = doc.at('td.detailaddress')
          add_to_parse = address_td.html.gsub("<label class=\"detailproplocatedat\">Subject Property</label>", "")
          add_to_parse = add_to_parse.gsub("<br /><br />", "")
          address, city, state, zip, county = parse_address(add_to_parse)
          address.strip!
          city.strip!
        else
          return
        end
        notice = 'N'
        # doc.search('span/a.blackformtext') do |a|
        #   notice = 'Y' if a.inner_text =~ /Preforeclosure/i
        #   notice = 'B' if a.inner_text =~ /Bank Owned/i
        #   notice = 'A' if a.inner_text =~ /Auction/i
        # end
        notice = 'B' if @watir.html =~ /greenLink/i
        notice = 'Y' if @watir.html =~ /redLink/i
        notice = 'A' if @watir.html =~ /orangeLink/i
        defaults = {
          'Co'    => (COUNTIES.index(county.upcase) || county).strip,
          'NOD'   => notice,
          'PAdd'  => address,
          'PCity' => city,
          'PZip'  => zip,
        }.merge(empty)
        data = {}.merge(defaults)
        got = Set.new

				# added  Last Market Sale section Recording Date
				data["Purch"] = doc.search('table').find{|x| 'Last Market Sale' == x['id']}.search('td').find{|x| x.innerText =~ /Recording Date/}.next.next.next.next.innerText rescue nil

        doc.search('table/tbody/tr/td.detailreportheader') do |td|
          title = td.inner_text.strip
          # irregular, and i don't think we need them
          # one key can have multiple values for their data
          puts title, '' if @verbose
          if got.include?(title)
            # yield data if block_given?
            #      return # we don't want any more records from here
            # data = {}.merge(defaults)
            # got.clear
          else
            got << title
            key = nil
            td.parent.parent.search('td') do |td|
              c = td[:class]
              next unless c =~ /detailreport/
              value = td.inner_text.strip
              if value.empty? and (key != "Mortgage Date:" and key != "Mortgage Amt:")
                next key = nil
              elsif c == 'detailreport'
                key = value
              elsif key
                # We gotta do a special case for mortgage information since keys are repeating
                if key =~ /Mortgage Date/i
                  if value.empty?
                    data["TD 1 Date"] = td.next.next.inner_text
                    data["TD 2 Date"] = td.next.next.next.next.inner_text
                  else
                    data["TD 1 Date"] = value
                    data["TD 2 Date"] = td.next.next.inner_text
                  end
                end
                if key =~ /Mortgage Amt/i
                  if value.empty?
                    data["TD 1"] = td.next.next.inner_text
                    data["TD 2"] = td.next.next.next.next.inner_text
                  else
                    data["TD 1"] = value
                    data["TD 2"] = td.next.next.inner_text
                  end
                end
                process_report_field(data, key, value)
                key = nil
              end
            end
          end
        end

=begin
        if @methodtmp == 3
					# don't use anymore
          if doc.to_s =~ /window.open\('([^']+)\'/i
            @watir.goto("http://realist2.firstamres.com/#{$1.gsub("&amp;", "&").gsub('#', '')}")
            if @watir.html =~ /<iframe .* src="([^"]+)"/i
              puts "Fetching Value Map from iFrame"
              @watir.goto($1.gsub("&amp;", "&").gsub('#', ''))
            end
            sleep 5
            counter = 0
            puts "Waiting for Estimated Value to load..."
            while !@watir.text.include?('Estimated Value : $')
              counter += 1
              if counter > 12
                puts "Waited too long for Estimated Value, moving to next..."
                break
              end
              sleep 5
            end
            begin
              #File.open('watir_html', 'w') {|f| f.write(@watir.html) }
							# removing price 1/6/2012
              #data["Price"] = Hpricot(@watir.html).search("span[@id=lblValue]").inner_html
              #puts "Estimated Value: "+data["Price"]
            rescue => e
              puts "Couldn't find estimated value"
            end
          end
        end
=end

        # updated on 1/6/2012
				# - P Guardiario pguardiario@gmail.com
				# http://imrmls.com/
				# click on CRMLS Matrix.
				# User is c12729
				# Password is griff123
				#
				data['NA'] = doc.search("//td[@colspan='8']").find{|x| 'Mortgage History:' == x.inner_text}.parent.next.next.search("//td[6]").inner_text rescue ''
				data['Mach'] = doc.search("//td[@colspan='8']").find{|x| 'Mortgage History:' == x.inner_text}.parent.next.next.next.next.search("//td[6]").inner_text rescue ''
				unit = doc.search("//td").find{|x| 'Status:' == x.inner_text}.next.next.next.next.inner_text rescue ''
				data['Unit'] = case unit
					when /active/i then 'A'
					when /pending/i then 'P'
					when /hold/i then 'H'
					when /backup/i then 'B'
					else ''
				end

				if @methodtmp == 3
					# rewritten to skip watir for better speed
					puts 'requesting value map data...'
					license = 'b9b2b1a56ec7481282d8f5c76e1a9123'
					address = [data['PAdd'], data['PCity'], data['MState'], data['PZip']].join(', ')
					host = 'valuemap.facorelogic.com'

					Net::HTTP.start(host, 80) do |http|
						request_json = {"licenseCode" => license,"Address" => address,"propertyType" => "","numBeds" => 0,"numBaths" => 0,"numTotalRooms" => 0,"livingArea" => 0,"yearBuilt" => 0,"currentValue" => 0,"languageCode" => "en-US","renderPropListHTML" => true,"requestType" => "New","leadNumber" => "0"}.to_json

						headers = {
							"User-Agent" => "User-Agent	Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0)",
							'Content-Type' => 'application/json; charset=UTF-8'
						}
						resp, jdata = http.post('/ValueMapService.asmx/GetPropertyInfoReport', request_json, headers)
						json = JSON.parse jdata
						data['Price'] = json['SubjectProperty']['CurrentValue'].to_s
					end
				end

				#
				# end 1/6/2012 updates

        yield data if block_given?
      rescue Object => exception
        puts "", " <ERROR> ".center(60, 'v')
        puts exception, *exception.backtrace
        puts " </ERROR> ".center(60, '^'), ""
        yield empty if block_given?
      end


      def search_resultAPN(query)
        begin
          @watir.link(:text, 'Detail').click
          # make code compatible with san bernardino apn here
          query['apn_entered'].to_s.gsub!(/-/,'')

          safely do
            @watir.select_list(:name, "Fm86_Ctrl9701_LB").select_value(COUNTY_CODES[query['county'].upcase])
          end
          sleep 2

          safely do
            @watir.text_field(:name, "Fm86_Ctrl9849_TextBox").set(query['apn_entered'])
          end
          sleep 2

          safely do
            @watir.select_list(:name, "m_ddPageSize").select_value("100")
          end
          sleep 2

          safely do
            @watir.button(:name, "m_btnSearch").click
          end
        rescue => e
          puts "Ran into a problem in search_resultAPN - retrying: #{e.backtrace}"
          retry
        end

        return Hpricot(@watir.html)
      end

       def run_search_residential(query)
        begin
          #@watir.link(:text, 'Detail').click
          @watir.link(:url, /\/Matrix\/Search\/Residential\/Detail/).click
          # make code compatible with san bernardino apn here
          query['apn_entered'].to_s.gsub!(/-/,'')

          safely do
            @watir.select_list(:name, "Fm9_Ctrl6238_LB").select_value(COUNTY_CODES[query['county'].upcase])
          end
          sleep 2

          safely do
            @watir.text_field(:name, "Fm9_Ctrl6232_TextBox").set(query['apn_entered'])
          end
          sleep 2

          safely do
            #@watir.select_list(:name, "m_ddPageSize").select_value("100")
            @watir.select_list(:id, "m_ucDisplayPicker_m_ddlPageSize").select_value("100")
          end
          sleep 2

          safely do
            #@watir.button(:id, "m_lbSearch").click
            @watir.link(:id, "m_ucSearchButtons_m_lbSearch").click

          end
        rescue => e
          puts "Ran into a problem in run_search_residential - retrying: #{e.backtrace}"
          retry
        end
         sleep 2
       return Hpricot(@watir.html)
      end

      def parse_reports_residential(doc, query = {})
        # File.open('result.html', 'w+'){|io| io.write(doc.to_s) }
        empty = { 'APN' => query['APN'] }
        sleep (2)
        doc.search('td.d673m5') do |td|
          print "APN: ", query['APN'], " Status: ", td.inner_text, "\n"
          empty["Unit"] = td.inner_text;
          empty["Unit"] = case td.inner_text
					  when /[^ABPH]/i then ''
					  else td.inner_text
				  end
          #if td.inner_text =~ /[APBH]/
          #  puts "true td.inner_text = /[APBH]/"
#         #   yield empty if block_given?
          #  return
          # end
          yield (empty) if block_given?
        end
      end

	  def parse_reportsAPN(doc, query = {})
        # File.open('result.html', 'w+'){|io| io.write(doc.to_s) }
        empty = { 'APN' => query['APN'] }
        doc.search('td.d686m7') do |td|
          print "APN: ", query['APN'], " Status: ", td.inner_text, "\n"
          if td.inner_text =~ /[APBH]/
            puts "true td.inner_text = /[APBH]/"
            yield empty if block_given?
            return
          end
        end
      end

      def process_report_field(hash, key, value)
        if key !~ /:/
          puts "Warning: %p is irregular, skip parsing it" % {key => value} if @verbose
          return
        end

        key = key.sub(/:$/, '')
        # puts "%40s : %s" % [key, value] # if @verbose
        placeholder = 0
        case key
        when /Owner Name 2/i
          owner_hash = parse_owner2(value)
          hash.merge!(owner_hash)
        when /Owner Name/i
          owner_hash = parse_owner(value)
          hash.merge!(owner_hash)
        when /Owner Phone/i
          hash['Phone'] = value.sub(/(\d{4})(\d{3})(\d*)/, '\1-\2-\3')
        when /Tax Billing City & State/i
          # ADD MODIFICATION HERE TO FIX THE MULTI WORDED CITY PROBLEM
          # turn it to array
          if value.split.length == 3
            hash['MCity'] = value.split[0] + " " + value.split[1]
            hash['MState'] = value.split[2]
          else
            #state, *city = value.split
            hash['MCity'] = value.split[0]
            hash['MState'] = value.split[1]
          end
        when /Tax Billing Zip/i
          #hash["MZip"] = "##{value}" edited original
          hash["MZip"] = "#{value.to_s}"
        when /Tax Billing Address/i
          address, city, state, zip, county = parse_address(value)
          hash["MAdd"] = address
        when /Recording Date/i
				# skip
		  #if !hash["Purch"] then
			#hash["Purch"] = value
		  #end
        when /Sale Price/i
        # no price returned when using method 1 - added 2/2010
					# removing price 1/6/2012
          #hash["Price"] = (@methodtmp == -1) ? '' : value
        when /Building Sq Ft/i
          hash["SF"] = value
        end

      end

      def parse_address(address)
        city = state = zip = county = nil
        address = address.to_s.split(/<br \/>/)
        address, middle, county = address

        if middle != "" && middle != nil
          if match = middle.match(/(.*),\s+([A-Z]+)\s+([\d-]+)/) then
            city, state, zip = middle.match(/(.*),\s+([A-Z]+)\s+([\d-]+)/).captures
          else
            city, state, zip = "", "", ""
          end
          county.sub!(/\s*County\s*$/, '')
        else
          city, state, zip = "", "", ""
        end

        [address, city, state, zip, county]
      end

      def parse_owner(full)
        atoms = full.split('&').first.split
        fn = mn = ln = nil

        case atoms.size
        when 2
          fn, ln = atoms.reverse
        when 3
          mn, fn, ln = atoms.reverse
        else
          # puts "Cannot parse name: #{full}"
          mn, fn, ln = atoms.reverse
        end

        {'Name' => fn, 'Initial' => mn, 'Last' => ln}
      end

      def parse_owner2(full)
        atoms = full.split('&').first.split
        fn = mn = ln = nil

        case atoms.size
        when 2
          fn, ln = atoms.reverse
        when 3
          mn, fn, ln = atoms.reverse
        else
          # puts "Cannot parse name: #{full}"
          mn, fn, ln = atoms.reverse
        end

        {'Name 2' => fn, 'Last 2' => ln}
      end

      # method argument is 0 or 1 depending on menu choice 3 or 4
      def search_reports(method, query, &block)
        state = (query['state'] || 'CA').upcase
        county = (query['Co'] || query['county'] || query['co']).to_s
        county = county.strip
        county = COUNTIES[county] if county.size <= 4
        return if county == nil
        county = county.upcase
        remarks = ""
        if county =~ /^san bernardino|sb$/i && query['APN'].to_s.length == 12
          #apn = query['APN'].to_s.gsub(/(\d{3})(\d{3})(\d{2})(0{4})?/,'0\1-\2-\3')
          apn = "0".to_s + query['APN'].to_s
          if county =~ /^orange|oc$/i && query['APN'].to_s.length == 7
            apn = "0".to_s + query['APN'].to_s
          end
        else
          apn = query['APN'].to_s
        end
        query = {
          'state' => state, 'apn_entered' => apn, 'county' => county
        }
        puts("Searching: %15s %20s" % [apn, county]) #if @verbose
        if method == 0 or method == 3
          parse_reports(search_result(query),
          'APN' => apn, 'State' => state, 'County' => county,
          &block)
        end
        if method == 1
          #query['Remarks'] = ""
          parse_reports_residential(run_search_residential(query),
          'APN' => apn, 'State' => state, 'County' => county, 'Remarks' => remarks,
          &block)
        end
        if method == 2
          @methodtmp = 1
          parse_reportsAPN(search_resultAPN(query),
          'APN' => apn, 'State' => state, 'County' => county, 'Remarks' => remarks,
          &block)
        end
      end

      # ADD method argument
      def search(method)
        first = true

        xls_each do |query|
          records = 0

          FileUtils.touch(out_file) # make sure it exists
          File.open(out_file, 'a') do |file|
            ################ MODIFIED NEW CODE 06/18/08 FOR USE WITH PROPERTY STATUS CHECK ###############
            # method = 0 will be standard original search, = 1 will be new search                        #
            # to do , add functionality and new argument to search_query(search_type[0 or 1],query)      #
            ##############################################################################################
            #this used to be a block of if statements for each method type
            #but that's pointless because they all did the same thing
            search_reports(method,query) do |report|
                out = query.merge(report)
                file.puts CSV.generate_line(out.keys) if File.size(out_file) == 0
                file.puts CSV.generate_line(out.values)
                records += 1
            end
            # yeah, we're grammatically correct today :)
            print("Fetched #{records} record", records == 1 ? "" : "s", "\n\n")
          end
        end
      end
    end



    class << self
      def start
        $stdout.sync = true
        @options = OpenStruct.new(DEFAULTS)

        parse_args
        print_copyright
        start_menu
      end

      def parse_args
        if xls = ARGV.find{|a| File.file?(a) and File.readable?(a) }
          @options.xls = xls
        else
          print_copyright
          print "No input file specified.\n",
          "Usage: ", File.basename("#$0"), " <path to excel file>\n"
          exit 1
        end
      end

      def print_copyright
        system("cls")
        puts "",
        "--- RealtXpress v2.1 ------------------------------------------",
        "-   Copyright (C) 2008 Charles Thompson, all rights reserved   ",
        "-   Licensed to Thomas Griffey of Anchor Properties.       ",
        "-   License holder has permission to modify source code.     ",
        "---------------------------------------------------------------",
        "-   Press CTRL+C at any time to terminate this application.   ",
        "---------------------------------------------------------------",
        "-   Modified: September 5th, 2009 - Denis Odorcic",
        "-   Modified: March 1st, 2010 - Josh Shupack",
        "-   Modified: 1/12/2012 - P Guardiario",
        "",
        ""
      end

      def start_menu
        menu_head
        choices = %w[csv_from_matrix status_check residential csv_from_matrix_with_map]
        print "  Your choice: "
        while got = $stdin.gets
          choice = got.to_i

          case choice
          when 1..4
            send choices[choice - 1]
          else
            choice == 5 ? exit : puts("Please choose an option 1-5.")
          end

          menu_head
          print "  Your choice: "
        end
      end

      def menu_head
        puts "--- Main Menu -------------------------------------------------",
        "-1. Feed this excel file into mrmlsmatrix.com                  ",
        "-2. Feed this excel file into the mrmlsmatrix status check     ",
        "-3. Feed this excel file into the (broken) status check        ",
        "-4. Feed this excel file into mrmlsmatrix.com with value map   ",
        "-5. Exit                                                       ",
        "---------------------------------------------------------------",
        "",
        ""
      end

      # choices


      def csv_from_matrix
        puts "1. Feed this excel file into mrmlsmatrix.com"
        matrix = MRMLSMatrix.new(@options, verbose = false)
        if matrix.login(0)
          matrix.methodtmp = -1 # not sure how this is being used, so didn't want to set it to 0
          matrix.search(0) 
        end
      end
      def status_check
        puts "2. Feed this excel file into the mrmlsmatrix status check"
        matrix = MRMLSMatrix.new(@options, verbose = false)
        matrix.search(1) if matrix.login(1)
        start_menu
      end
      def residential
        puts "3. Feed this excel file into the residential status check"
        matrix = MRMLSMatrix.new(@options, verbose = false)
        matrix.search(2) if matrix.login(2)
        start_menu
      end
      def csv_from_matrix_with_map
        puts "4. Feed this excel file into mrmlsmatrix.com with value map"
        matrix = MRMLSMatrix.new(@options, verbose = false)
        matrix.search(3) if matrix.login(3)
      end
    end
  end
  ExcelBot.start

end
