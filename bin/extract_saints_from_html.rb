#!/usr/bin/env ruby

require 'nokogiri'

file = ARGV.shift
doc = File.open(file) { |f| Nokogiri::HTML(f) }

doc.xpath('//tr').each do |node|
  cells = node.xpath('td')
  name = cells[0] && cells[0].text
  dates = node.xpath('./td/a[@class="choixDate"]').map { |d| d.text.strip }.join '|'

  puts "#{name}\t#{dates}" unless dates.empty?
end