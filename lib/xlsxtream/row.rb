# frozen_string_literal: true
require 'date'
require 'xlsxtream/xml'

module Xlsxtream
  class Row
    ENCODING = Encoding.find('UTF-8')

    NUMBER_PATTERN = /\A-?[0-9]+(\.[0-9]+)?\z/.freeze
    # ISO 8601 yyyy-mm-dd
    DATE_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}\z/.freeze
    # ISO 8601 yyyy-mm-ddThh:mm:ss(.s)(Z|+hh:mm|-hh:mm)
    TIME_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,9})?)?(?:Z|[+-][0-9]{2}:[0-9]{2})?\z/.freeze

    TRUE_STRING = 'true'
    FALSE_STRING = 'false'

    DATE_STYLE = 1
    TIME_STYLE = 2
    FLOAT_STYLE = 3

    def initialize(row, rownum, options = {})
      @row = row
      @rownum = rownum
      @sst = options[:sst]
      @auto_format = options[:auto_format]
    end

    def to_xml
      column = +'A'
      xml = +"<row r=\"#{@rownum}\">"

      @row.each do |value|
        cid = "#{column}#{@rownum}"
        column.next!

        value = auto_format(value) if @auto_format && value.is_a?(String)

        case value
        when Float
          xml << %(<c r="#{cid}" s="#{FLOAT_STYLE}" t="n"><v>#{value}</v></c>)
        when Numeric
          xml << %(<c r="#{cid}" t="n"><v>#{value}</v></c>)
        when TrueClass, FalseClass
          xml << %(<c r="#{cid}" t="b"><v>#{value ? 1 : 0}</v></c>)
        when Time
          xml << %(<c r="#{cid}" s="#{TIME_STYLE}"><v>#{time_to_oa_date(value)}</v></c>)
        when DateTime
          xml << %(<c r="#{cid}" s="#{DATE_STYLE}"><v>#{datetime_to_oa_date(value)}</v></c>)
        when Date
          xml << %(<c r="#{cid}" s="#{DATE_STYLE}"><v>#{date_to_oa_date(value)}</v></c>)
        else
          value = value.to_s

          unless value.empty? # no xml output for for empty strings
            value = value.encode(ENCODING) if value.encoding != ENCODING

            xml << if @sst
                     %(<c r="#{cid}" t="s"><v>#{@sst[value]}</v></c>)
                   else
                     %(<c r="#{cid}" t="inlineStr"><is><t>#{XML.escape_value(value)}</t></is></c>)
                   end
          end
        end
      end

      xml << '</row>'
    end

    private

    # Detects and casts numbers, date, time in text
    def auto_format(value)
      case value
      when TRUE_STRING
        true
      when FALSE_STRING
        false
      when NUMBER_PATTERN
        value.include?('.') ? value.to_f : value.to_i
      when DATE_PATTERN
        begin
          Date.parse(value)
        rescue StandardError
          value
        end
      when TIME_PATTERN
        begin
          DateTime.parse(value)
        rescue StandardError
          value
        end
      else
        value
      end
    end

    # Converts Time instance to OLE Automation Date
    def time_to_oa_date(time)
      # Local dates are stored as UTC by truncating the offset:
      # 1970-01-01 00:00:00 +0200 => 1970-01-01 00:00:00 UTC
      # This is done because SpreadsheetML is not timezone aware.
      (time.to_f + time.utc_offset) / 86400 + 25569
    end

    # Converts DateTime instance to OLE Automation Date
    if RUBY_ENGINE == 'ruby'
      def datetime_to_oa_date(date)
        _, jd, df, sf, of = date.marshal_dump
        jd - 2415019 + (df + of + sf / 1e9) / 86400
      end
    else
      def datetime_to_oa_date(date)
        date.jd - 2415019 + (date.hour * 3600 + date.sec + date.sec_fraction.to_f) / 86400
      end
    end

    # Converts Date instance to OLE Automation Date
    def date_to_oa_date(date)
      (date.jd - 2415019).to_f
    end
  end
end
