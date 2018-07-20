require 'rubyXL'
class ExcelAccessor

  # シートを読み込む（シートが未指定の場合は最初のシート）
  # 引数
  #   - path: ファイルの読み込み先のパス
  #   - sheet(Integer / String): シート番号(Integer) または シート名(String)を指定可能
  def load!( path: nil, sheet: nil )

    # 読み込み
    if path.nil?
      raise "\n============ Error!! ============\n読み込むファイルのパスが未指定です。\n============ Error!! ============"
    end

    @read_file_path = path
    @workbook = RubyXL::Parser.parse(@read_file_path)

    # sheetが指定されている場合
    if sheet
      # sheetに指定可能なのはsheetのindex(Integer)またはsheet名(String)
      if [Integer, String].include?(sheet.class)
        @worksheet = @workbook[sheet]
      end

      # sheetの指定に誤りがある場合はエラー
      if @worksheet.nil?
        raise "\n============ Error!! ============\nsheetが見つからないか、指定に誤りがあります。\n============ Error!! ============"
      end

    # sheetが指定されていない場合
    else
      # 最初のシートを取得する
      @worksheet = @workbook.first
    end

  end

  # シートの内容を更新する
  # 引数
  #   - cells(Array of Hash): col(行番号)、row(列番号)、value(値)からなるハッシュの配列
  def update!(cells: [])
    if @worksheet.nil?
      raise "\n============ Error!! ============\nシートがロードされていません。先にload!メソッドを実行する必要があります。\n============ Error!! ============"
    end

    cells.each do |cell|
      @worksheet.add_cell(cell[:row], cell[:col], cell[:value])
    end
  end

  # シートの中身を表示する
  def show
    @worksheet.each do |row|

      values = []

      row && row.cells.each do |cell|

        cell && values.push(cell.value)
      end

      p "| #{values.join(' | ')} |"
    end
  end

  # シートの内容を指定ファイルにダンプする
  # 引数
  #   - path: ファイルの書き込み先のパス
  def dump(path: @read_file_path)
    @workbook.write(path)
  end

  attr_accessor :workbook, :worksheet
end
