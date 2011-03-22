require 'csv'

class Excel
  attr_reader :worksheets

  def initialize(spreadsheet)
    if spreadsheet.content_type == Mime::Type.lookup_by_extension(:xls)
      raw_data = `java -jar lib/xls2csv.jar #{spreadsheet.tempfile.path}`
    elsif spreadsheet.content_type == Mime::Type.lookup_by_extension(:xlsx)
      raw_data = `java -jar lib/xlsx2csv.jar #{spreadsheet.tempfile.path}`
    else
      raise 'Only supports Excel documents'
    end

    # Convert to UTF-8
    # See: http://po-ru.com/diary/fixing-invalid-utf-8-in-ruby-revisited/
    ic = Iconv.new('UTF-8//IGNORE', 'UTF-8')
    raw_data = ic.iconv(raw_data + ' ')[0..-2]

    @worksheets = raw_data.split(/^__EOF__\n/).collect do |sheet|
      details, data = sheet.split(/\n/,2)
      name, index = details.match(/.*(?= \[index=(\d+))/).to_a
      rows = CSV.parse(data)
      Excel::Worksheet.new(name, index.to_i, rows)
    end
  end

  def worksheet(id)
    id =
    case id
    when Integer
      @worksheets[id]
    when Regexp
      @worksheets.find {|sheet| sheet.name =~ id }
    when String
      @worksheets.find {|sheet| sheet.name == id }
    end
  rescue
    nil
  end
end

class Excel::Worksheet
  attr_reader :name, :index

  def initialize(name, index, rows)
    @name, @index, @rows = name, index, rows
  end

  def to_a
    @rows
  end

  # Returns an array of hashes using the first row as the keys
  def to_hashed_array
    # Duplicate the rows so that we can modify it without losing data
    rows = @rows.dup

    # Use the header row's values as property keys
    keys = rows.shift.collect {|name| string_to_key(name) }

    # Collect the attributes
    rows.collect do |row|
      Hash[keys.collect {|n| [n, row[keys.index(n)]] }]
    end
  end

  def to_hash
    # Duplicate the rows so that we can modify it without losing data
    rows = @rows.dup

    # The first row are keys, so remove to make easier looping
    header_row = rows.shift
    # Use the header row's values as property names.
    # - Ignore first column since it will be used as keys for each row
    headings = header_row.drop(1).collect {|heading| heading.strip }

    # Use first column of rows as keys
    keys = rows.collect {|row| string_to_key(row.shift) }

    # Array of hashes
    # - Keys are first row
    # - Values are respecitive columns of keys as hashes with first column as key
    headings.inject({}) do |cols, heading|
      cols.merge(heading => Hash[keys.collect {|key| [key, rows[keys.index(key)][headings.index(heading)]] }] )
    end
  end

private

  # Strip, downcase and replace spaces with underscores
  def string_to_key(val)
    val.strip.downcase.gsub(/\s/,'_')
  end
end