# frozen_string_literal: true
require "bigdecimal"
require "date"
require "xlsxtream/xml"

module Xlsxtream
  class Row

    ENCODING = Encoding.find('UTF-8')

    NUMBER_PATTERN = /\A-?[0-9]+(\.[0-9]+)?\z/.freeze
    # ISO 8601 yyyy-mm-dd
    DATE_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}\z/.freeze
    # ISO 8601 yyyy-mm-ddThh:mm:ss(.s)(Z|+hh:mm|-hh:mm)
    TIME_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,9})?)?(?:Z|[+-][0-9]{2}:[0-9]{2})?\z/.freeze

    TRUE_STRING = 'true'.freeze
    FALSE_STRING = 'false'.freeze

    STYLES = {
      normal: { general: nil, date: 1, time: 2, float: 6 },
      header: { general: 3, date: 4, time: 5, float: 11 },
      bold: { general: 7, date: 8, time: 9, float: 10 }
    }.freeze

    def initialize(row, rownum, options = {})
      @row = row
      @rownum = rownum
      @sst = options[:sst]
      @auto_format = options[:auto_format]
      @style = options[:is_header] ? STYLES[:header] : STYLES[options[:bold] ? :bold : :normal]
    end

    def to_xml
      column = String.new('A')
      xml = String.new(%Q{<row r="#{@rownum}">})

      @row.each do |value|
        cid = "#{column}#{@rownum}"
        column.next!

        if @auto_format && value.is_a?(String)
          value = auto_format(value)
        end

        case value
        when Float
          xml << numeric_cell(cid, value, :float)
        when BigDecimal
          xml << numeric_cell(cid, value.to_s("F"), :float)
        when Numeric
          xml << numeric_cell(cid, value, :general)
        when TrueClass, FalseClass
          xml << %Q{<c r="#{cid}"#{style(:general)} t="b"><v>#{value ? 1 : 0}</v></c>}
        when Time
          xml << %Q{<c r="#{cid}"#{style(:time)}><v>#{time_to_oa_date(value)}</v></c>}
        when DateTime
          xml << %Q{<c r="#{cid}"#{style(:date)}><v>#{datetime_to_oa_date(value)}</v></c>}
        when Date
          xml << %Q{<c r="#{cid}"#{style(:date)}><v>#{date_to_oa_date(value)}</v></c>}
        else
          value = value.to_s

          unless value.empty? # no xml output for for empty strings
            value = value.encode(ENCODING) if value.encoding != ENCODING

            if value.start_with?('=')
              xml << %Q{<c r="#{cid}"#{style(:general)}><f>#{XML.escape_formula(value.sub('=', ''))}</f></c>}
            elsif @sst
              xml << %Q{<c r="#{cid}"#{style(:general)} t="s"><v>#{@sst[value]}</v></c>}
            else
              xml << %Q{<c r="#{cid}"#{style(:general)} t="inlineStr"><is><t>#{XML.escape_value(value)}</t></is></c>}
            end
          end
        end
      end

      xml << '</row>'
    end

    private

    def numeric_cell(cid, value, type)
      %Q{<c r="#{cid}"#{style(type)} t="n"><v>#{value}</v></c>}
    end

    def style(type)
      style_id = @style.fetch(type)
      style_id ? %Q{ s="#{style_id}"} : ''
    end

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
        Date.parse(value) rescue value
      when TIME_PATTERN
        DateTime.parse(value) rescue value
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
