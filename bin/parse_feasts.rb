#!/usr/bin/env ruby
require 'csv'
require 'axlsx'

class Feast
  attr_accessor :line, :chunks, :name, :attrs, :mods, :dates

  DATE_RE = %r{\d{1,2}/Août\.|\d{1,2}/Avril\.|\d{1,2}/Dec\.|\d{1,2}/Fev\.|\d{1,2}/Jan\.|\d{1,2}/Jul\.|\d{1,2}/Jun\.|\d{1,2}/Mai\.|\d{1,2}/Mars\.|\d{1,2}/Nov\.|\d{1,2}/Oct\.|\d{1,2}/Sept\.}

  ATTR_REGEX = %r{\bab\.|\babb\.|\bap\.|\bapp\.|\barchang\.|\bcan\.|\bcf\.|\bcff\.|\bd\.|\bdiac\.|\bdisc\.|\bdoct\.|\bduc\.|\bdux|\bep\.|\ber\.|\berem\.|\betc\.|\bev\.|\bfil\.|\bfr\.|\bgerman\.|\bhosp\.|\bimp\.|\binv\.|\blandgr\.|\blev\.|\bm\.|\bmarch\.|\bmatr\.|\bmil\.|\bmm\.|\bmon\.|\bp\.|\bpatr\.|\bpb\.|\bpp\.|\bpph\.|\bprepos\.|\bprotom\.|\br\.|\brecl\.|\breg\.|\bs\.|\bsolit\.|\bsubd\.|\bv\.|\bvel\.|\bvid\.|\bvv\.}

  OTHER_RE = %r{[[:alnum:]]+[[:punct:]]?}

  BRACKET_RE = %r{\[[^\]]+\]}

  MOD_RE = %r{\([^)]+\)}

  def initialize line
    @line   = line.strip
    @chunks = []
    @attrs  = []
    @mods   = []
    @dates  = []
    @name   = ''
  end

  def parse
    lex_line
    extract_name
    extract_attrs
    extract_mods
    extract_dates

    check_unconsumed
  end

  def extract_name
    name = ''
    # name is the first :other chunk
    name = next_chunk.join(' ') if next_chunk_is? :other

    while next_chunk_is? :other, :bracket
      name += " #{next_chunk.join(' ')}"
    end

    @name = name
  end

  def extract_attrs
    attrs = []
    while next_chunk_is? :attr, :other, :bracket
      case next_chunk_type
      when :attr
        next_chunk.each { |a| attrs << a }
      when :other
        attrs[-1] += " (#{next_chunk.join(' ')})"
      when :bracket
        attrs[-1] += " #{next_chunk.join(' ')}"
      end
    end

    @attrs = attrs
  end

  def extract_mods
    # all the :modifier chunks
    @mods = next_chunk if next_chunk_is? :modifier
  end

  def extract_dates
    @dates = next_chunk if next_chunk_is? :date
  end

  # Sample line:
  #
  #       Barbara v. m. Nicomed. (Trans.)  05/Dec.;16/Dec.;04/Dec.
  #
  # chunked into:
  #     :other  => [ Barbara ]
  #     :attrs  => [ v., m. ]
  #     :other  => [ Nicomed. ]
  #     :mods   => [ (Trans.) ]
  #     :dates  => [ 05/Dec., 16/Dec., 04/Dec. ]
  def lex_line
    @chunks = @line.scan(/#{BRACKET_RE}|#{DATE_RE}|#{MOD_RE}|#{ATTR_REGEX}|#{OTHER_RE}/).chunk do |seg|
      case seg
        when BRACKET_RE  then :bracket
        when MOD_RE      then :modifier
        when DATE_RE     then :date
        when ATTR_REGEX  then :attr
        when OTHER_RE    then :other
        else                  :error
      end
    end.to_a
    # it should be impossible to encounter errors, but cry if they happen
    check_for_errors
    @chunks
  end

  # See if there are any :error chunks, and print a warning if so.
  def check_for_errors
    return unless @chunks.any? { |chunk| chunk.first[0] == :error }
    errors = @chunks.find_all { |chunk| chunk[0] == :error  }
    $stderr.puts "WARNING: found the following errors: #{errors.map(&last).inspect}"
  end

  # If we haven't used up the line, something's wrong; print a warning.
  def check_unconsumed
    return if chunks.empty?
    $stderr.puts "WARNING: line not consumed: #{@chunks}; (line: '#{@line}')"
  end

  def next_chunk
    @chunks.shift[1]
  end

  def next_chunk_is? *symbols
    symbols.include? next_chunk_type
  end

  def next_chunk_type
    @chunks.first && @chunks.first[0]
  end
end


outfile = File.expand_path '../../data/bollandistes.xlsx', __FILE__

p = Axlsx::Package.new
wb = p.workbook

wb.add_worksheet name: 'feasts' do |wksh|
  headings = %w{Name Attributes Modifiers Dates Line }
  wksh.add_row headings

  ARGF.each do |line|
    feast = Feast.new(line)
    feast.parse
    (row ||= []) << feast.name
    row << feast.attrs.join(' | ')
    row << feast.mods.join(' | ')
    row << feast.dates.join(' | ')
    row << feast.line
    wksh.add_row row
  end
  widths = (0..headings.size).map { 30 }
  wksh.column_widths(*widths)
end

p.serialize outfile

puts "Wrote #{outfile}"

