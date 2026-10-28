# frozen_string_literal: true
require 'test_helper'
require 'stringio'
require 'tempfile'
require 'zip'
require 'xlsxtream/workbook'

module Xlsxtream
  class WorksheetTest < Minitest::Test

    class SpyWriter < Xlsxtream::ZipKitWriter
      def initialize
        @paths_to_file_contents = {}
        @current = nil
      end

      def <<(data)
        @paths_to_file_contents[@current] << data
      end

      def add_file(path)
        @current = path
        @paths_to_file_contents[@current] = String.new
      end

      def [](key)
        @paths_to_file_contents.fetch(key)
      end

      def close
      end
    end

    def test_workbook_from_path
      tempfile = Tempfile.new('xlsxtream')
      Workbook.open(tempfile.path) {}
      refute_equal 0, tempfile.size
    ensure
      tempfile.close! if tempfile
    end

    def test_workbook_from_io
      tempfile = Tempfile.new('xlsxtream')
      Workbook.open(tempfile) {}
      refute_equal 0, tempfile.size
    ensure
      tempfile.close! if tempfile
    end

    def test_empty_workbook
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) {}
      expected = {
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets></sheets>' \
          '</workbook>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
          '</Relationships>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_workbook_with_sheet
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet
      end
      expected = {
        'xl/worksheets/sheet1.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
            '<sheetData></sheetData>' \
          '</worksheet>',
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets>' \
              '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
            '</sheets>' \
          '</workbook>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
          '</Relationships>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_deprecated_add_workbook_with_block
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        silence_warnings do
          wb.add_worksheet {}
        end
      end
      expected = {
        'xl/worksheets/sheet1.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
            '<sheetData></sheetData>' \
          '</worksheet>',
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets>' \
              '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
            '</sheets>' \
          '</workbook>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
          '</Relationships>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_workbook_with_sheet_without_block
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        ws = wb.add_worksheet
        ws << ['foo']
        ws.close
      end
      expected = {
        'xl/worksheets/sheet1.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
            '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>foo</t></is></c></row></sheetData>' \
          '</worksheet>',
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets>' \
              '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
            '</sheets>' \
          '</workbook>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
          '</Relationships>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_workbook_with_sst
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet(nil, use_shared_strings: true) do |ws|
          ws << ['foo']
        end
      end
      expected = {
        'xl/worksheets/sheet1.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
            '<sheetData>' \
              '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' \
            '</sheetData>' \
          '</worksheet>',
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets>' \
              '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
            '</sheets>' \
          '</workbook>',
        'xl/sharedStrings.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">' \
            '<si><t>foo</t></si>' \
          '</sst>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' \
          '</Relationships>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_root_relations
      iow_spy = io_wrapper_spy
      Workbook.new(iow_spy).close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' \
        '</Relationships>'
      actual = iow_spy['_rels/.rels']
      assert_equal expected, actual
    end

    def test_content_types
      iow_spy = io_wrapper_spy
      Workbook.new(iow_spy).close
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' \
          '<Default Extension="xml" ContentType="application/xml"/>' \
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' \
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' \
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' \
        '</Types>'
      actual = iow_spy['[Content_Types].xml']
      assert_equal expected, actual
    end

    def test_write_multiple_worksheets
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet
        wb.write_worksheet
      end

      expected = {
        'xl/workbook.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            '<workbookPr date1904="false"/>' \
            '<sheets>' \
              '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
              '<sheet name="Sheet2" sheetId="2" r:id="rId2"/>' \
            '</sheets>' \
          '</workbook>',
        'xl/_rels/workbook.xml.rels' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' \
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
          '</Relationships>',
        'xl/worksheets/sheet1.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>',
        'xl/worksheets/sheet2.xml' =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>'
      }
      actual = iow_spy
      expected.keys.each do |path|
        assert_equal expected[path], actual[path]
      end
    end

    def test_must_write_sequentially
      iow_spy1 = io_wrapper_spy

      Workbook.open(iow_spy1) do |wb|
        wb.add_worksheet.tap { |ws| ws.close }
        wb.add_worksheet.tap { |ws| ws.close }
      end

      iow_spy2 = io_wrapper_spy
      assert_raises(Xlsxtream::Error) do
        Workbook.open(iow_spy2) do |wb|
          wb.add_worksheet
          wb.add_worksheet # adding a second worksheet without closing
        end
      end
    end

    def test_write_named_worksheet
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet('foo')
      end

      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
          '<workbookPr date1904="false"/>' \
          '<sheets>' \
            '<sheet name="foo" sheetId="1" r:id="rId1"/>' \
          '</sheets>' \
        '</workbook>'
      actual = iow_spy['xl/workbook.xml']
      assert_equal expected, actual
    end

    def test_write_unnamed_worksheet_with_options
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet(:use_shared_strings => true)
      end

      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '\
                  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
          '<workbookPr date1904="false"/>' \
          '<sheets>' \
            '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' \
          '</sheets>' \
        '</workbook>'
      actual = iow_spy['xl/workbook.xml']
      assert_equal expected, actual
    end

    def test_worksheet_name_as_option
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) do |workbook|
        workbook.write_worksheet(name: "foo")
      end
      expected = '<sheet name="foo" sheetId="1" r:id="rId1"/>'
      actual = iow_spy['xl/workbook.xml'][/<sheet [^>]+>/]
      assert_equal expected, actual
    end

    def test_add_columns_via_workbook_options
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy, { :columns => [ {}, {}, { :width_pixels => 42 } ] } ) do |wb|
        wb.write_worksheet {}
      end

      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>' \
          '<col min="1" max="1"/>' \
          '<col min="2" max="2"/>' \
          '<col min="3" max="3" width="42" customWidth="1"/>' \
        '</cols>' \
        '<sheetData></sheetData></worksheet>'

      actual = iow_spy['xl/worksheets/sheet1.xml']
      assert_equal expected, actual
    end

    def test_add_columns_via_workbook_options_and_add_rows
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy, { :columns => [ {}, {}, { :width_pixels => 42 } ] } ) do |wb|
        wb.write_worksheet do |ws|
          ws << ['foo']
          ws.add_row ['bar']
        end
      end

      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>' \
          '<col min="1" max="1"/>' \
          '<col min="2" max="2"/>' \
          '<col min="3" max="3" width="42" customWidth="1"/>' \
        '</cols>' \
        '<sheetData>' \
          '<row r="1"><c r="A1" t="inlineStr"><is><t>foo</t></is></c></row>' \
          '<row r="2"><c r="A2" t="inlineStr"><is><t>bar</t></is></c></row>' \
        '</sheetData></worksheet>'

      actual = iow_spy['xl/worksheets/sheet1.xml']
      assert_equal expected, actual
    end

    def test_styled_archive_preserves_values_and_formulas
      [false, true].each do |use_sst|
        [false, true].each do |auto_format|
          output = StringIO.new
          values = ['A&B _x0041_', true, 42, 1.23456789, BigDecimal('123456789.123456789'),
            Date.new(1900, 1, 1), DateTime.new(1900, 1, 1, 12, 30, 0, '+05:30'),
            Time.new(1900, 1, 1, 12, 30, 0, '+05:30'), '=SUM(C2:E2)', '=IF(C2<50,"A&B","_x0041_")']
          Workbook.open(output, use_shared_strings: use_sst, auto_format: auto_format) do |wb|
            wb.write_worksheet do |ws|
              ws.add_row values, bold: true
              ws << values
              ws.add_row values, bold: true
              ws.add_row values
            end
            wb.write_worksheet('Centered', has_header_row: true) do |ws|
              ws << values
              ws << values
            end
          end
          Zip::File.open_buffer(output.string) do |zip|
            sheet = zip.read('xl/worksheets/sheet1.xml')
            styles = zip.read('xl/styles.xml')
            xfs = styles[/<cellXfs[^>]*>(.*?)<\/cellXfs>/, 1].scan(/<xf\b[^>]*\/>|<xf\b[^>]*>.*?<\/xf>/)
            fonts = styles[/<fonts[^>]*>(.*?)<\/fonts>/, 1].scan(/<font>.*?<\/font>/)
            rows = sheet.scan(/<row\b[^>]*>.*?<\/row>/)
            assert_equal 4, rows.size
            rows.each_with_index do |row, index|
              bold = index.even?
              cells = row.scan(/<c\b[^>]*>.*?<\/c>/)
              expected_formats = [0, 0, 0, 4, 4, 164, 164, 165, 0, 0]
              cells.each_with_index do |cell, column|
                style_id = cell[/\bs="(\d+)"/, 1].to_i
                xf = xfs.fetch(style_id)
                assert_equal expected_formats[column], xf[/numFmtId="(\d+)"/, 1].to_i
                font = fonts.fetch(xf[/fontId="(\d+)"/, 1].to_i)
                assert_equal bold, font.include?('<b val="1"/>')
                refute_includes xf, '<alignment'
              end
              assert_includes cells[1], '<v>1</v>'
              assert_includes cells[2], '<v>42</v>'
              assert_includes cells[3], '<v>1.23456789</v>'
              assert_includes cells[4], '<v>123456789.123456789</v>'
              assert_includes cells[5], '<v>2.0</v>'
              [6, 7].each do |column|
                assert_in_delta 2 + 12.5 / 24, cells[column][/<v>(.*?)<\/v>/, 1].to_f, 1e-10
              end
              assert_includes cells[8], '<f>SUM(C2:E2)</f>'
              assert_includes cells[9], '<f>IF(C2&lt;50,"A&amp;B","_x0041_")</f>'
              refute_match(/\bt="/, cells[8])
              refute_match(/\bt="/, cells[9])
            end
            strings = use_sst ? zip.read('xl/sharedStrings.xml') : sheet
            assert_includes strings, 'A&amp;B _x005f_x0041_'
            centered = zip.read('xl/worksheets/sheet2.xml').scan(/<row\b[^>]*>.*?<\/row>/)
            centered.first.scan(/<c\b[^>]*>.*?<\/c>/).each do |cell|
              xf = xfs.fetch(cell[/\bs="(\d+)"/, 1].to_i)
              assert_includes xf, '<alignment horizontal="center" vertical="center"/>'
            end
            refute_includes centered.last, 's="3"'
          end
        end
      end
    end

    def test_styled_rows_stream_before_source_is_exhausted
      output = StringIO.new
      row_count = 0
      emitted_at = []
      sink = Object.new
      sink.define_singleton_method(:<<) do |chunk|
        emitted_at << row_count
        output << chunk
      end
      Workbook.open(sink) do |wb|
        wb.write_worksheet do |ws|
          ws.add_row ['Header'], bold: true
          2000.times do |index|
            row_count += 1
            # Varied input exceeds both the writer buffer and deflate's output buffer.
            ws << [index, Array.new(40) { |column| "#{index * 7919 + column * 104729}" }.join('-')]
          end
          ws.add_row ['Total', '=SUM(A2:A2001)'], bold: true
        end
      end
      assert emitted_at.any? { |count| count > 0 && count < 2000 }, emitted_at.inspect
      Zip::File.open_buffer(output.string) do |zip|
        sheet = zip.read('xl/worksheets/sheet1.xml')
        assert_equal 2002, sheet.scan(/<row\b/).size
        assert_includes sheet, '<c r="B2002" s="7"><f>SUM(A2:A2001)</f></c>'
        assert_includes sheet, '</sheetData></worksheet>'
        assert_includes zip.read('xl/styles.xml'), '<cellXfs count="12">'
      end
    end

    def test_styles_content
      iow_spy = io_wrapper_spy
      Workbook.open(iow_spy) {}
      expected = \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'"\r\n" \
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
          '<numFmts count="2">' \
            '<numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/>' \
            '<numFmt numFmtId="165" formatCode="yyyy\\-mm\\-dd hh:mm:ss"/>' \
          '</numFmts>' \
          '<fonts count="2">' \
            '<font>' \
              '<sz val="12"/>' \
              '<name val="Calibri"/>' \
              '<family val="2"/>' \
            '</font>' \
            '<font>' \
              '<b val="1"/>' \
              '<sz val="12"/>' \
              '<name val="Calibri"/>' \
              '<family val="2"/>' \
            '</font>' \
          '</fonts>' \
          '<fills count="2">' \
            '<fill>' \
              '<patternFill patternType="none"/>' \
            '</fill>' \
            '<fill>' \
              '<patternFill patternType="gray125"/>' \
            '</fill>' \
          '</fills>' \
          '<borders count="1">' \
            '<border/>' \
          '</borders>' \
          '<cellStyleXfs count="1">' \
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' \
          '</cellStyleXfs>' \
          '<cellXfs count="12">' \
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' \
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' \
            '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' \
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1">' \
              '<alignment horizontal="center" vertical="center"/>' \
            '</xf>' \
            '<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1">' \
              '<alignment horizontal="center" vertical="center"/>' \
            '</xf>' \
            '<xf numFmtId="165" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1">' \
              '<alignment horizontal="center" vertical="center"/>' \
            '</xf>' \
            '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' \
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>' \
            '<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' \
            '<xf numFmtId="165" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' \
            '<xf numFmtId="4" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' \
            '<xf numFmtId="4" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1">' \
              '<alignment horizontal="center" vertical="center"/>' \
            '</xf>' \
          '</cellXfs>' \
          '<cellStyles count="1">' \
            '<cellStyle name="Normal" xfId="0" builtinId="0"/>' \
          '</cellStyles>' \
          '<dxfs count="0"/>' \
          '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>' \
        '</styleSheet>'
      actual = iow_spy['xl/styles.xml']
      assert_equal expected, actual
    end

    def test_custom_font_size
      iow_spy = io_wrapper_spy
      font_options = { :size => 23 }
      Workbook.open(iow_spy, :font => font_options) {}
      expected = '<sz val="23"/>'
      actual = iow_spy['xl/styles.xml'][/<sz [^>]+>/]
      assert_equal expected, actual
    end

    def test_custom_font_name
      iow_spy = io_wrapper_spy
      font_options = { :name => 'Comic Sans' }
      Workbook.open(iow_spy, :font => font_options) {}
      expected = '<name val="Comic Sans"/>'
      actual = iow_spy['xl/styles.xml'][/<name [^>]+>/]
      assert_equal expected, actual
    end

    def test_custom_font_family
      iow_spy = io_wrapper_spy
      font_options = { :family => 'Script' }
      Workbook.open(iow_spy, :font => font_options) {}
      expected = '<family val="4"/>'
      actual = iow_spy['xl/styles.xml'][/<family [^>]+>/]
      assert_equal expected, actual
    end

    def test_font_family_mapping
      tests = {
        nil => 0,
        ''  => 0,
        'ROMAN' => 1,
        :roman => 1,
        'Roman' => 1,
        :swiss => 2,
        :modern => 3,
        :script => 4,
        :decorative => 5
      }
      tests.each do |value, id|
        iow_spy = io_wrapper_spy
        font_options = { :family => value }
        Workbook.open(iow_spy, :font => font_options) {}
        expected = "<family val=\"#{id}\"/>"
        actual = iow_spy['xl/styles.xml'][/<family [^>]+>/]
        assert_equal expected, actual
      end
    end

    def test_invalid_font_family
      iow_spy = io_wrapper_spy
      font_options = { :family => 'Foo' }
      assert_raises Xlsxtream::Error do
        Workbook.open(iow_spy, :font => font_options) {}
      end
    end

    def test_tempfile_is_not_closed
      tempfile = Tempfile.new('workbook')
      Workbook.open(tempfile) {}
      assert_equal false, tempfile.closed?
    ensure
      tempfile && tempfile.close!
    end

    private

    def io_wrapper_spy
      SpyWriter.new
    end

    def silence_warnings
      old_verbose, $VERBOSE = $VERBOSE, nil
      yield
    ensure
      $VERBOSE = old_verbose
    end

  end
end
