require_relative './lib/excel_accessor.rb'

# オブジェクトの初期化
excel_accessor = ExcelAccessor.new

# ファイルの読み込み
excel_accessor.load!( path: './data/sample.xlsx', sheet: 0 )

# セルの更新
cells = [
  {row: 2, col: 0, value: 'value3'},
  {row: 2, col: 1, value: 'value4'},
]
excel_accessor.update!( cells: cells )

# シートの内容を確認
excel_accessor.show

# ファイルの書き出し
excel_accessor.dump( path: './data/out_sample.xlsx' )
