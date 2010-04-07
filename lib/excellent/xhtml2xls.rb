class XlsCellStyle
  attr_reader :raw_style, :style

  ATTRIBUTES_MAP = {
    "font-weight" => :weight,
    "color" => :color,
    "font-size" => :size,
    "background-color" => :pattern_fg_color
  }

  def initialize(style)
    return nil unless style
    @raw_style = style.gsub(' ', '')
    @style = Spreadsheet::Format.new(style2hsh)
  end

  def style2hsh
    attrs = @raw_style.split(';')
    h = attrs.inject({}){|acc, attr| acc.merge(encode_attribute(attr))}
    h.merge(:pattern=>h.has_key?(:pattern_fg_color) ? 1 : 0)
  end

  def encode_attribute(attribute)
    attr, value = attribute.split(':')

    value = case attr
    when "color" then
      value.to_sym
    when "background-color" then
      value.to_sym
    when "font-size" then
      value.to_i
    when "font-weight" then
      value.to_sym
    end

    return Hash[attr, value].map_keys(ATTRIBUTES_MAP)
  end
end

class XlsCellMerge
  attr_reader :merge
  def initialize(merge)
    return nil unless merge
    @merge = merge.to_i
  end
end

class XlsCell

  attr_reader :data

  def initialize(cell, kind='normal')
    if cell.is_a? Hash
      cell.symbolize_keys!
      @data = cell[:content] || ''
      @style = XlsCellStyle.new(cell[:style])
      @merge = XlsCellMerge.new(cell[:colspan])
    else
      @data = cell || ''
      @style = nil
      @merge = nil
    end
  end

  def merge
    @merge ? @merge.merge : nil
  end

  def style
    @style ? @style.style : nil
  end

  def style_to_hash
    @style && @style.raw_style ? @style.style2hsh : {}
  end

end

class XlsRow
  attr_reader :cells

  def initialize(row)
    kind = row.keys.include?(:th) ? 'header' : 'normal'
    @cells = row.values.sum.map{|c| XlsCell.new(c, kind)}
  end

end

class XlsTable
  attr_reader :rows

  REQUIRED_FIRST_LEVEL_KEYS = ["name", "tr"]

  def initialize(hsh)
    validate_hash(hsh)
    @rows = hsh["tr"].map{|tr| XlsRow.new(tr)}
  end

  private

  def validate_hash(hsh)
    raise "Table has to have 'name' attribute and a set of rows" if !(REQUIRED_FIRST_LEVEL_KEYS-hsh.keys ).empty?
    raise "tr is to be array" unless hsh["tr"].kind_of?(Array)
  end

end

class Hsh2Xls
  require 'spreadsheet' if defined? Spreadsheet::Format
  require 'xmlsimple'

  attr_reader :xhtml

  def initialize(xhtml)
    @xhtml = xhtml
    hsh = XmlSimple.xml_in(@xhtml, {'ForceArray' => true}).symbolize_keys
    if hsh[:li]
      @hsh = returning Array.new do |result|
        hsh[:li].each do |li|
          result << li["table"].first
        end
      end
    else
      hsh_new = {}
      hsh.each_pair {|k, v| hsh_new[k.to_s] = v}
      @hsh = [hsh_new]
    end
  end

  private

  def render_xls
    xls = Spreadsheet::Workbook.new

    @hsh.each do |hsh|
      @sheet = xls.create_worksheet :name => hsh["name"]
      xls_table = XlsTable.new(hsh)

      xls_table.rows.each_with_index do |row, row_idx|
        row.cells.each_with_index do |cell, cell_idx|
          render_cell(row_idx, cell_idx, cell)
        end
      end
    end

    xls_output = StringIO.new
    xls.write xls_output
    return xls_output.string
  end

  def render_cell(row_idx, cell_idx, cell)
    if cell.merge
      merge = cell_idx +cell.merge
      merge = cell_idx if merge < cell_idx
      style = Spreadsheet::Format.new({:align=>:merge}.merge(cell.style_to_hash))
      cell_idx.upto(merge) do |i|
        @sheet.row(row_idx).set_format(i, style)
      end
    else
      @sheet.row(row_idx).set_format(cell_idx, cell.style) if cell.style
    end
    if cell.data.empty?
      @sheet[row_idx, cell_idx] = cell.data
    else
      @sheet[row_idx, cell_idx] = cell.data.match(/[\D]/) ? cell.data : cell.data.to_i
    end
  end

  public

  def output
    render_xls
  end

end